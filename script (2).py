# -*- coding: utf-8 -*-
"""
pyRevit pushbutton: Per-Element Compound Layer BOQ Export

Exports one Excel row for every compound-structure layer of every individual
Walls, Floors, Roofs and Ceilings instance in the active Revit model.

Important behaviour:
- Elements are NOT grouped by type or material.
- Each instance is identified by its Revit Element ID.
- Each layer of that instance gets its own row (Layer Order 1, 2, 3, ...).
- Thickness is converted to metres.
- Wall Length comes from LocationCurve; Wall Height comes from geometry Z extent.
- Floor/Roof/Ceiling Length and Width are derived from the horizontal XY
  geometry envelope. The shorter horizontal dimension is written in the
  "Height / Width" column, as requested.
- Area is taken from Revit's HOST_AREA_COMPUTED parameter where available;
  geometry/footprint estimates are used only as fallbacks.
- Volume is a PER-LAYER material volume estimate: Area x Layer Thickness.
  It is not the whole host element's volume.

The script only reads the model and writes an .xlsx file. It does not modify
Revit elements. It uses only Python standard-library modules for Excel output.
"""

import os
import re
import zipfile
from xml.sax.saxutils import escape

from Autodesk.Revit.DB import (
    BuiltInCategory,
    BuiltInParameter,
    FilteredElementCollector,
    HostObjAttributes,
    Options,
    ViewDetailLevel,
    Solid,
    GeometryInstance,
    PlanarFace,
)
from pyrevit import forms


# -----------------------------------------------------------------------------
# Revit document and constants
# -----------------------------------------------------------------------------

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

CATEGORIES = [
    (u"Walls", BuiltInCategory.OST_Walls),
    (u"Floors", BuiltInCategory.OST_Floors),
    (u"Roofs", BuiltInCategory.OST_Roofs),
    (u"Ceilings", BuiltInCategory.OST_Ceilings),
]

SQFT_TO_SQM = 0.09290304
FEET_TO_M = 0.3048

try:
    text_type = unicode  # IronPython 2
except NameError:
    text_type = str      # CPython 3

try:
    number_types = (int, long, float)
except NameError:
    number_types = (int, float)


def as_text(value):
    if value is None:
        return u""
    try:
        return text_type(value)
    except Exception:
        try:
            return text_type(repr(value))
        except Exception:
            return u""


def safe_id(element_id):
    if element_id is None:
        return u""
    try:
        return as_text(element_id.Value)
    except Exception:
        try:
            return as_text(element_id.IntegerValue)
        except Exception:
            return as_text(element_id)


def get_parameter(element, bip):
    try:
        return element.get_Parameter(bip)
    except Exception:
        return None


def get_double_parameter(element, bip, factor):
    parameter = get_parameter(element, bip)
    if parameter is None:
        return None
    try:
        if not parameter.HasValue:
            return None
        return float(parameter.AsDouble()) * factor
    except Exception:
        return None


def get_type_name(element_type, type_id):
    """Resolve the element type name without silently grouping unknown types."""
    if element_type is not None:
        for bip_name in ("ALL_MODEL_TYPE_NAME", "SYMBOL_NAME_PARAM"):
            try:
                bip = getattr(BuiltInParameter, bip_name)
                parameter = element_type.get_Parameter(bip)
                if parameter is not None and parameter.HasValue:
                    value = parameter.AsString()
                    if value and as_text(value).strip():
                        return as_text(value).strip()
                    value = parameter.AsValueString()
                    if value and as_text(value).strip():
                        return as_text(value).strip()
            except Exception:
                pass
        try:
            value = element_type.Name
            if value and as_text(value).strip():
                return as_text(value).strip()
        except Exception:
            pass

    return u"<Type Name Unavailable - ID {0}>".format(safe_id(type_id) or u"N/A")


def get_compound_structure(element_type):
    if element_type is None:
        return None
    try:
        if isinstance(element_type, HostObjAttributes):
            return element_type.GetCompoundStructure()
    except Exception:
        pass
    try:
        return element_type.GetCompoundStructure()
    except Exception:
        return None


def get_material_name(material_id):
    try:
        if material_id is None:
            return u"<No Material>"
        material = doc.GetElement(material_id)
        if material is None:
            return u"<No Material>"
        return as_text(material.Name)
    except Exception:
        return u"<No Material>"


def iter_solids(geometry_element):
    """Yield non-empty solids, including solids nested in family instances."""
    if geometry_element is None:
        return
    for obj in geometry_element:
        try:
            if isinstance(obj, Solid):
                if obj.Faces.Size > 0:
                    yield obj
            elif isinstance(obj, GeometryInstance):
                nested = obj.GetInstanceGeometry()
                for solid in iter_solids(nested):
                    yield solid
        except Exception:
            continue


def convex_hull(points):
    """Monotonic-chain convex hull for XY points (ordinary Python tuples)."""
    unique_points = sorted(set((round(p[0], 6), round(p[1], 6)) for p in points))
    if len(unique_points) <= 1:
        return unique_points

    def cross(o, a, b):
        return ((a[0] - o[0]) * (b[1] - o[1]) -
                (a[1] - o[1]) * (b[0] - o[0]))

    lower = []
    for point in unique_points:
        while len(lower) >= 2 and cross(lower[-2], lower[-1], point) <= 0:
            lower.pop()
        lower.append(point)

    upper = []
    for point in reversed(unique_points):
        while len(upper) >= 2 and cross(upper[-2], upper[-1], point) <= 0:
            upper.pop()
        upper.append(point)

    return lower[:-1] + upper[:-1]


def oriented_xy_dimensions_m(points):
    """Return long/short dimensions of the minimum-area oriented XY envelope."""
    hull = convex_hull(points)
    if len(hull) < 2:
        return (None, None)

    best_area = None
    best_length = None
    best_width = None

    for index in range(len(hull)):
        p1 = hull[index]
        p2 = hull[(index + 1) % len(hull)]
        dx = p2[0] - p1[0]
        dy = p2[1] - p1[1]
        edge_length = (dx * dx + dy * dy) ** 0.5
        if edge_length <= 1.0e-9:
            continue

        ux = dx / edge_length
        uy = dy / edge_length
        vx = -uy
        vy = ux

        projections_u = [p[0] * ux + p[1] * uy for p in hull]
        projections_v = [p[0] * vx + p[1] * vy for p in hull]
        dim_u = max(projections_u) - min(projections_u)
        dim_v = max(projections_v) - min(projections_v)
        area = dim_u * dim_v

        if best_area is None or area < best_area:
            best_area = area
            best_length = max(dim_u, dim_v)
            best_width = min(dim_u, dim_v)

    if best_length is None:
        return (None, None)
    return (best_length * FEET_TO_M, best_width * FEET_TO_M)


def get_geometry_metrics(element):
    """Return geometry extents and projected footprint area in metric units."""
    options = Options()
    options.DetailLevel = ViewDetailLevel.Fine
    options.IncludeNonVisibleObjects = False
    options.ComputeReferences = False

    points_xyz = []
    points_xy = []
    positive_projected_area_ft2 = 0.0
    negative_projected_area_ft2 = 0.0
    solid_count = 0

    try:
        geometry = element.get_Geometry(options)
    except Exception:
        geometry = None

    for solid in iter_solids(geometry):
        solid_count += 1

        # Collect actual solid edge points for the geometry envelope.
        try:
            for edge in solid.Edges:
                try:
                    for point in edge.Tessellate():
                        x = float(point.X)
                        y = float(point.Y)
                        z = float(point.Z)
                        points_xyz.append((x, y, z))
                        points_xy.append((x, y))
                except Exception:
                    pass
        except Exception:
            pass

        # Project non-vertical planar faces onto the XY plane. Top and bottom
        # faces are accumulated separately so the footprint is not doubled.
        try:
            for face in solid.Faces:
                if not isinstance(face, PlanarFace):
                    continue
                try:
                    normal_z = float(face.FaceNormal.Z)
                    if abs(normal_z) < 0.01:
                        continue
                    mesh = face.Triangulate()
                    projected = 0.0
                    for tri_index in range(mesh.NumTriangles):
                        tri = mesh.get_Triangle(tri_index)
                        a = tri.get_Vertex(0)
                        b = tri.get_Vertex(1)
                        c = tri.get_Vertex(2)
                        cross_z = ((float(b.X) - float(a.X)) * (float(c.Y) - float(a.Y)) -
                                   (float(b.Y) - float(a.Y)) * (float(c.X) - float(a.X)))
                        projected += abs(cross_z) * 0.5
                    if normal_z > 0:
                        positive_projected_area_ft2 += projected
                    else:
                        negative_projected_area_ft2 += projected
                except Exception:
                    pass
        except Exception:
            pass

    # Use an element bounding box only when Revit did not provide solid edge points.
    if not points_xyz:
        try:
            bbox = element.get_BoundingBox(None)
            if bbox is not None:
                for x in (float(bbox.Min.X), float(bbox.Max.X)):
                    for y in (float(bbox.Min.Y), float(bbox.Max.Y)):
                        for z in (float(bbox.Min.Z), float(bbox.Max.Z)):
                            points_xyz.append((x, y, z))
                            points_xy.append((x, y))
        except Exception:
            pass

    if points_xyz:
        xs = [p[0] for p in points_xyz]
        ys = [p[1] for p in points_xyz]
        zs = [p[2] for p in points_xyz]
        z_extent_m = (max(zs) - min(zs)) * FEET_TO_M
    else:
        xs = ys = None
        z_extent_m = None

    long_xy_m, short_xy_m = oriented_xy_dimensions_m(points_xy)
    if long_xy_m is None and xs is not None:
        dx = (max(xs) - min(xs)) * FEET_TO_M
        dy = (max(ys) - min(ys)) * FEET_TO_M
        long_xy_m = max(dx, dy)
        short_xy_m = min(dx, dy)

    projected_area_ft2 = max(positive_projected_area_ft2,
                             negative_projected_area_ft2)
    projected_area_sqm = (projected_area_ft2 * SQFT_TO_SQM
                          if projected_area_ft2 > 1.0e-9 else None)

    return {
        "length_xy_m": long_xy_m,
        "width_xy_m": short_xy_m,
        "height_z_m": z_extent_m,
        "projected_area_sqm": projected_area_sqm,
        "solid_count": solid_count,
    }


def get_location_curve_length_m(element):
    try:
        location = element.Location
        curve = location.Curve if location is not None else None
        if curve is not None:
            return float(curve.Length) * FEET_TO_M
    except Exception:
        pass
    return None


def get_element_area_sqm(element, category_name, length_m, height_or_width_m,
                         geometry_metrics):
    """Use Revit host area first, then category-appropriate geometry fallback."""
    area = get_double_parameter(
        element, BuiltInParameter.HOST_AREA_COMPUTED, SQFT_TO_SQM
    )
    if area is not None and area > 1.0e-9:
        return area

    projected_area = geometry_metrics.get("projected_area_sqm")
    if projected_area is not None and projected_area > 1.0e-9:
        return projected_area

    # Last-resort rectangular estimate. The README and workbook label describe
    # the area source rule; this is not used when Revit provides a host area.
    if length_m is not None and height_or_width_m is not None:
        estimate = length_m * height_or_width_m
        if estimate > 1.0e-9:
            return estimate
    return None


def rounded(value, digits=6):
    if value is None:
        return None
    try:
        return round(float(value), digits)
    except Exception:
        return None


# -----------------------------------------------------------------------------
# Collect one output row per element instance AND per compound layer
# -----------------------------------------------------------------------------


def collect_layer_rows():
    rows = []
    skipped_no_compound = 0
    scanned_instances = 0

    for category_name, built_in_category in CATEGORIES:
        collector = (FilteredElementCollector(doc)
                     .OfCategory(built_in_category)
                     .WhereElementIsNotElementType())

        for element in collector:
            scanned_instances += 1
            try:
                type_id = element.GetTypeId()
                element_type = doc.GetElement(type_id)
            except Exception:
                type_id = None
                element_type = None

            element_id_text = safe_id(element.Id)
            element_name = get_type_name(element_type, type_id)
            compound = get_compound_structure(element_type)

            if compound is None:
                skipped_no_compound += 1
                continue

            try:
                layers = list(compound.GetLayers())
            except Exception:
                layers = []

            if not layers:
                skipped_no_compound += 1
                continue

            geometry = get_geometry_metrics(element)
            location_length_m = get_location_curve_length_m(element)

            if category_name == u"Walls":
                # For walls, the Revit location curve is the actual wall path.
                length_m = location_length_m
                if length_m is None:
                    length_m = geometry.get("length_xy_m")
                height_or_width_m = geometry.get("height_z_m")
            else:
                # For floors, roofs and ceilings, use the plan-view dimensions.
                # The column is explicitly labelled Height / Width because the
                # user requested the second horizontal dimension for these types.
                length_m = geometry.get("length_xy_m")
                height_or_width_m = geometry.get("width_xy_m")

            area_sqm = get_element_area_sqm(
                element, category_name, length_m, height_or_width_m, geometry
            )

            # Keep every instance separate. Do not group by Type, Material, or Layer.
            for layer_index, layer in enumerate(layers):
                try:
                    thickness_m = float(layer.Width) * FEET_TO_M
                except Exception:
                    thickness_m = None

                try:
                    material_id = layer.MaterialId
                except Exception:
                    material_id = None

                material_type = get_material_name(material_id)
                material_id_text = safe_id(material_id)
                if material_type == u"<No Material>":
                    material_id_text = u""

                # Layer volume is the material-specific estimate, not the entire
                # wall/floor/roof/ceiling host solid volume.
                layer_volume_m3 = None
                if area_sqm is not None and thickness_m is not None:
                    layer_volume_m3 = area_sqm * thickness_m

                rows.append([
                    category_name,
                    element_name,
                    element_id_text,
                    material_type,
                    material_id_text,
                    layer_index + 1,
                    rounded(thickness_m),
                    rounded(length_m),
                    rounded(height_or_width_m),
                    rounded(area_sqm),
                    rounded(layer_volume_m3),
                ])

    return {
        "rows": rows,
        "scanned_instances": scanned_instances,
        "skipped_no_compound": skipped_no_compound,
    }


# -----------------------------------------------------------------------------
# XLSX writer (Open XML, standard library only)
# -----------------------------------------------------------------------------


def column_letter(column_number):
    result = u""
    number = int(column_number)
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        result = unichr(65 + remainder) + result
    return result


try:
    unichr
except NameError:
    unichr = chr


def clean_xml_text(value):
    value = as_text(value)
    filtered = []
    for char in value:
        codepoint = ord(char)
        if codepoint in (9, 10, 13) or codepoint >= 32:
            filtered.append(char)
    return escape(u"".join(filtered), {u'"': u"&quot;"})


def cell_xml(cell_reference, value, style_id=0):
    if value is None or value == u"":
        return u""

    style = u' s="{0}"'.format(style_id) if style_id else u""
    if isinstance(value, bool):
        return u'<c r="{0}"{1} t="b"><v>{2}</v></c>'.format(
            cell_reference, style, u"1" if value else u"0"
        )
    if isinstance(value, number_types):
        if isinstance(value, float):
            numeric_text = as_text(format(value, ".12g"))
        else:
            numeric_text = as_text(value)
        return u'<c r="{0}"{1}><v>{2}</v></c>'.format(
            cell_reference, style, numeric_text
        )

    text = clean_xml_text(value)
    return (u'<c r="{0}"{1} t="inlineStr"><is><t xml:space="preserve">'
            u'{2}</t></is></c>').format(cell_reference, style, text)


def worksheet_xml(rows):
    max_columns = max([len(row) for row in rows]) if rows else 1
    last_cell = u"{0}{1}".format(column_letter(max_columns), max(len(rows), 1))
    widths = [15, 32, 14, 30, 14, 13, 16, 14, 20, 14, 14]
    pieces = [
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        u'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        u'<dimension ref="A1:{0}"/>'.format(last_cell),
        u'<sheetViews><sheetView workbookViewId="0">'
        u'<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        u'</sheetView></sheetViews>',
        u'<sheetFormatPr defaultRowHeight="18"/>',
        u'<cols>',
    ]
    for index, width in enumerate(widths[:max_columns], 1):
        pieces.append(u'<col min="{0}" max="{0}" width="{1}" customWidth="1"/>'.format(index, width))
    pieces.append(u'</cols><sheetData>')

    for row_index, row in enumerate(rows, 1):
        pieces.append(u'<row r="{0}">'.format(row_index))
        for column_index, value in enumerate(row, 1):
            # Style 1 = header; style 2 = numeric metres/metres-squared/metres-cubed;
            # style 3 = layer order (integer).
            if row_index == 1:
                style_id = 1
            elif column_index == 6:
                style_id = 3
            elif column_index >= 7:
                style_id = 2
            else:
                style_id = 0
            pieces.append(cell_xml(
                u"{0}{1}".format(column_letter(column_index), row_index),
                value,
                style_id,
            ))
        pieces.append(u'</row>')

    pieces.append(u'</sheetData>')
    if rows and max_columns:
        pieces.append(u'<autoFilter ref="A1:{0}"/>'.format(last_cell))
    pieces.append(u'</worksheet>')
    return u"".join(pieces)


def write_xlsx(file_path, rows):
    content_types = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        u'<Default Extension="xml" ContentType="application/xml"/>'
        u'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        u'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        u'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        u'</Types>'
    )
    root_rels = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        u'</Relationships>'
    )
    workbook_xml = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        u'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        u'<bookViews><workbookView/></bookViews><sheets>'
        u'<sheet name="Material BOQ" sheetId="1" r:id="rId1"/>'
        u'</sheets></workbook>'
    )
    workbook_rels = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        u'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        u'</Relationships>'
    )
    styles_xml = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        u'<numFmts count="2"><numFmt numFmtId="164" formatCode="0.000000"/>'
        u'<numFmt numFmtId="165" formatCode="0"/></numFmts>'
        u'<fonts count="2">'
        u'<font><sz val="10"/><name val="Calibri"/></font>'
        u'<font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Calibri"/></font>'
        u'</fonts>'
        u'<fills count="3">'
        u'<fill><patternFill patternType="none"/></fill>'
        u'<fill><patternFill patternType="gray125"/></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>'
        u'</fills>'
        u'<borders count="2">'
        u'<border><left/><right/><top/><bottom/><diagonal/></border>'
        u'<border><left style="thin"><color rgb="FFD9E2F3"/></left><right style="thin"><color rgb="FFD9E2F3"/></right><top style="thin"><color rgb="FFD9E2F3"/></top><bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border>'
        u'</borders>'
        u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        u'<cellXfs count="4">'
        u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        u'<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" horizontal="center" wrapText="1"/></xf>'
        u'<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
        u'<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
        u'</cellXfs>'
        u'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        u'<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
        u'</styleSheet>'
    )

    archive = zipfile.ZipFile(file_path, "w", zipfile.ZIP_DEFLATED)
    try:
        parts = [
            ("[Content_Types].xml", content_types),
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", workbook_xml),
            ("xl/_rels/workbook.xml.rels", workbook_rels),
            ("xl/styles.xml", styles_xml),
            ("xl/worksheets/sheet1.xml", worksheet_xml(rows)),
        ]
        for part_name, part_content in parts:
            if not isinstance(part_content, bytes):
                part_content = part_content.encode("utf-8")
            archive.writestr(part_name, part_content)
    finally:
        archive.close()


# -----------------------------------------------------------------------------
# pyRevit button execution
# -----------------------------------------------------------------------------


def main():
    if doc is None:
        forms.alert(u"No Revit document is currently open.", title=u"Material BOQ Export")
        return

    forms.alert(
        u"This export creates one row for each compound layer of each individual "
        u"Wall, Floor, Roof or Ceiling instance. Elements are not grouped by type or material.\n\n"
        u"Volume = element area x that layer's thickness.\n"
        u"All dimensions and quantities are exported in metres, square metres and cubic metres.",
        title=u"Material BOQ Export",
    )

    data = collect_layer_rows()
    if not data["rows"]:
        forms.alert(
            u"No compound-structure layers were found in Walls, Floors, Roofs or Ceilings.\n\n"
            u"Instances scanned: {0}".format(data["scanned_instances"]),
            title=u"Material BOQ Export",
        )
        return

    model_title = as_text(doc.Title)
    safe_title = re.sub(r"[^A-Za-z0-9_.-]+", "_", model_title).strip("_.") or "Revit_Model"
    default_name = u"{0}_Material_BOQ.xlsx".format(safe_title)
    output_path = forms.save_file(
        file_ext=u"xlsx",
        default_name=default_name,
        title=u"Save Material BOQ Excel File",
    )
    if not output_path:
        return
    if not output_path.lower().endswith(u".xlsx"):
        output_path += u".xlsx"

    try:
        headers = [
            u"Category",
            u"Name",
            u"Element ID",
            u"Material Type",
            u"Material ID",
            u"Layer Order",
            u"Thickness (m)",
            u"Length (m)",
            u"Height / Width (m)",
            u"Area (m²)",
            u"Volume (m³)",
        ]
        write_xlsx(output_path, [headers] + data["rows"])
    except Exception as export_error:
        forms.alert(
            u"Excel export failed.\n\n{0}".format(as_text(export_error)),
            title=u"Material BOQ Export - Error",
        )
        raise

    forms.alert(
        u"Excel export completed.\n\n"
        u"File: {0}\n\n"
        u"Model instances scanned: {1}\n"
        u"Layer rows exported: {2}\n"
        u"Instances without compound layers: {3}\n\n"
        u"Each row represents one layer of one individual element instance.".format(
            output_path,
            data["scanned_instances"],
            len(data["rows"]),
            data["skipped_no_compound"],
        ),
        title=u"Material BOQ Export",
    )


main()
