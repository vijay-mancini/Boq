# -*- coding: utf-8 -*-
"""
pyRevit pushbutton: Compound Structure + Geometry Export

Reads compound structure layers and solid geometry from Walls, Floors, Roofs and Ceilings in the
active Revit model and exports an Excel .xlsx workbook.

No document changes are made. No third-party Python packages are required.
The XLSX file is written using Python's standard-library zipfile module and
Open XML worksheet parts, so Excel does not need to be installed.

Tested for syntax with Python 3. The code avoids f-strings/type annotations
and is intended to remain compatible with common pyRevit IronPython/CPython
engines. Revit API methods used are available in modern Revit releases.
"""

import os
import re
import zipfile
import datetime
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
# Revit document and export configuration
# -----------------------------------------------------------------------------

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

CATEGORIES = [
    ("Walls", BuiltInCategory.OST_Walls),
    ("Floors", BuiltInCategory.OST_Floors),
    ("Roofs", BuiltInCategory.OST_Roofs),
    ("Ceilings", BuiltInCategory.OST_Ceilings),
]

SQFT_TO_SQM = 0.09290304
CUFT_TO_CUM = 0.028316846592
FEET_TO_MM = 304.8
FEET_TO_M = 0.3048


# -----------------------------------------------------------------------------
# General helpers
# -----------------------------------------------------------------------------

try:
    text_type = unicode  # IronPython 2
except NameError:
    text_type = str      # CPython 3

try:
    number_types = (int, long, float)
except NameError:
    number_types = (int, float)

try:
    unichr
except NameError:
    unichr = chr


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
    """Return an ElementId as text (supports newer 64-bit IDs)."""
    if element_id is None:
        return u""
    try:
        return as_text(element_id.Value)
    except Exception:
        try:
            return as_text(element_id.IntegerValue)
        except Exception:
            return as_text(element_id)


def get_parameter(element, built_in_parameter):
    try:
        return element.get_Parameter(built_in_parameter)
    except Exception:
        return None


def get_string_parameter(element, built_in_parameter):
    parameter = get_parameter(element, built_in_parameter)
    if parameter is None:
        return u""
    try:
        value = parameter.AsString()
        if value:
            return as_text(value)
    except Exception:
        pass
    try:
        return as_text(parameter.AsValueString())
    except Exception:
        return u""


def get_double_parameter(element, built_in_parameter, factor):
    parameter = get_parameter(element, built_in_parameter)
    if parameter is None:
        return None
    try:
        if not parameter.HasValue:
            return None
        return float(parameter.AsDouble()) * factor
    except Exception:
        return None


def get_level_name(element):
    """Read the associated level where Revit exposes one."""
    try:
        level_id = element.LevelId
        level = doc.GetElement(level_id)
        if level is not None:
            return as_text(level.Name)
    except Exception:
        pass

    # Fallback for host elements without a directly exposed LevelId.
    for bip_name in ("LEVEL_PARAM", "FAMILY_LEVEL_PARAM"):
        try:
            bip = getattr(BuiltInParameter, bip_name)
            parameter = element.get_Parameter(bip)
            if parameter and parameter.HasValue:
                level_id = parameter.AsElementId()
                level = doc.GetElement(level_id)
                if level is not None:
                    return as_text(level.Name)
        except Exception:
            pass
    return u""


def get_type_name(element_type):
    try:
        return as_text(element_type.Name)
    except Exception:
        return u"<Unknown Type>"


def clean_xml_text(value):
    """Escape text and remove XML 1.0 control characters."""
    value = as_text(value)
    filtered = []
    for char in value:
        codepoint = ord(char)
        if codepoint in (9, 10, 13) or codepoint >= 32:
            filtered.append(char)
    return escape(u"".join(filtered), {u'"': u"&quot;"})


def safe_filename(value):
    value = re.sub(r"[^A-Za-z0-9_.-]+", "_", as_text(value)).strip("_.")
    return value or "Revit_Model"


def get_compound_structure(element_type):
    """Return CompoundStructure or None for non-compound/system types."""
    try:
        if isinstance(element_type, HostObjAttributes):
            return element_type.GetCompoundStructure()
    except Exception:
        pass
    # A defensive fallback for API proxy classes / alternate Python engines.
    try:
        return element_type.GetCompoundStructure()
    except Exception:
        return None



def get_bounding_box_extents_m(element):
    """Return axis-aligned model bounding-box X/Y/Z extents in metres."""
    try:
        bbox = element.get_BoundingBox(None)
        if bbox is None:
            return (None, None, None)
        dx = abs(float(bbox.Max.X) - float(bbox.Min.X)) * FEET_TO_M
        dy = abs(float(bbox.Max.Y) - float(bbox.Min.Y)) * FEET_TO_M
        dz = abs(float(bbox.Max.Z) - float(bbox.Min.Z)) * FEET_TO_M
        return (dx, dy, dz)
    except Exception:
        return (None, None, None)


def get_location_curve_length_m(element):
    """Return LocationCurve length in metres, or None when not applicable."""
    try:
        location = element.Location
        curve = location.Curve if location is not None else None
        if curve is not None:
            return float(curve.Length) * FEET_TO_M
    except Exception:
        pass
    return None


def get_wall_unconnected_height_m(element):
    """Read the wall's unconnected-height parameter when it is available."""
    try:
        return get_double_parameter(
            element, BuiltInParameter.WALL_USER_HEIGHT_PARAM, FEET_TO_M
        )
    except Exception:
        return None

def iter_solids(geometry_element):
    """Yield non-empty solids, including solids inside family instances."""
    if geometry_element is None:
        return

    for geometry_object in geometry_element:
        try:
            if isinstance(geometry_object, Solid):
                if geometry_object.Faces.Size > 0 and geometry_object.Volume > 1.0e-9:
                    yield geometry_object
            elif isinstance(geometry_object, GeometryInstance):
                # GetInstanceGeometry returns the instance geometry in model space.
                instance_geometry = geometry_object.GetInstanceGeometry()
                for nested_solid in iter_solids(instance_geometry):
                    yield nested_solid
        except Exception:
            # Some geometry objects (for example, curves or meshes) are not solids.
            continue


def get_solid_geometry_metrics(element):
    """
    Read solid geometry and return model-axis extents, areas, volume and edge length.

    Extents are calculated from tessellated solid edges in model coordinates. They
    are true model-axis envelopes, not oriented/local dimensions. Tessellation is
    used only to find extents; edge lengths use Revit Curve.Length.
    """
    options = Options()
    options.DetailLevel = ViewDetailLevel.Fine
    options.IncludeNonVisibleObjects = False
    options.ComputeReferences = False

    min_x = min_y = min_z = None
    max_x = max_y = max_z = None
    solid_volume_ft3 = 0.0
    surface_area_ft2 = 0.0
    top_planar_area_ft2 = 0.0
    bottom_planar_area_ft2 = 0.0
    vertical_planar_area_ft2 = 0.0
    edge_length_ft = 0.0
    solid_count = 0
    edge_count = 0

    try:
        geometry = element.get_Geometry(options)
    except Exception:
        geometry = None

    for solid in iter_solids(geometry):
        solid_count += 1

        try:
            solid_volume_ft3 += float(solid.Volume)
        except Exception:
            pass

        try:
            for face in solid.Faces:
                try:
                    face_area = float(face.Area)
                    surface_area_ft2 += face_area
                    if isinstance(face, PlanarFace):
                        normal_z = float(face.FaceNormal.Z)
                        if normal_z >= 0.9:
                            top_planar_area_ft2 += face_area
                        elif normal_z <= -0.9:
                            bottom_planar_area_ft2 += face_area
                        else:
                            vertical_planar_area_ft2 += face_area
                except Exception:
                    continue
        except Exception:
            pass

        try:
            for edge in solid.Edges:
                try:
                    edge_curve = edge.AsCurve()
                    edge_length_ft += float(edge_curve.Length)
                    edge_count += 1
                except Exception:
                    pass

                # Tessellated points provide an envelope for straight and curved edges.
                try:
                    for point in edge.Tessellate():
                        x = float(point.X)
                        y = float(point.Y)
                        z = float(point.Z)
                        min_x = x if min_x is None else min(min_x, x)
                        min_y = y if min_y is None else min(min_y, y)
                        min_z = z if min_z is None else min(min_z, z)
                        max_x = x if max_x is None else max(max_x, x)
                        max_y = y if max_y is None else max(max_y, y)
                        max_z = z if max_z is None else max(max_z, z)
                except Exception:
                    pass
        except Exception:
            pass

    source = u"Solid geometry" if solid_count > 0 and min_x is not None else u"Unavailable"

    # If no solid edges were returned, retain a clearly labelled bounding-box fallback.
    if min_x is None:
        try:
            bbox = element.get_BoundingBox(None)
            if bbox is not None:
                min_x, min_y, min_z = (float(bbox.Min.X), float(bbox.Min.Y), float(bbox.Min.Z))
                max_x, max_y, max_z = (float(bbox.Max.X), float(bbox.Max.Y), float(bbox.Max.Z))
                source = u"Element bounding box fallback"
        except Exception:
            pass

    if min_x is None:
        x_extent = y_extent = z_extent = None
    else:
        x_extent = abs(max_x - min_x) * FEET_TO_M
        y_extent = abs(max_y - min_y) * FEET_TO_M
        z_extent = abs(max_z - min_z) * FEET_TO_M

    return {
        "x_extent_m": x_extent,
        "y_extent_m": y_extent,
        "z_extent_m": z_extent,
        "surface_area_sqm": surface_area_ft2 * SQFT_TO_SQM if solid_count else None,
        "volume_cum": solid_volume_ft3 * CUFT_TO_CUM if solid_count else None,
        "top_area_sqm": top_planar_area_ft2 * SQFT_TO_SQM if solid_count else None,
        "bottom_area_sqm": bottom_planar_area_ft2 * SQFT_TO_SQM if solid_count else None,
        "vertical_area_sqm": vertical_planar_area_ft2 * SQFT_TO_SQM if solid_count else None,
        "edge_length_m": edge_length_ft * FEET_TO_M if solid_count else None,
        "solid_count": solid_count,
        "edge_count": edge_count,
        "source": source,
    }


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





# -----------------------------------------------------------------------------
# Collect model data
# -----------------------------------------------------------------------------


def collect_model_data():
    element_rows = []
    layer_rows = []
    category_stats = []
    skipped_type_rows = []

    total_instances = 0
    elements_with_compound = 0
    elements_without_compound = 0
    unique_type_ids = set()

    for category_name, built_in_category in CATEGORIES:
        category_count = 0
        category_compound_count = 0
        category_no_compound_count = 0

        collector = (FilteredElementCollector(doc)
                     .OfCategory(built_in_category)
                     .WhereElementIsNotElementType())

        for element in collector:
            total_instances += 1
            category_count += 1

            element_id_text = safe_id(element.Id)
            type_id = None
            element_type = None
            try:
                type_id = element.GetTypeId()
                element_type = doc.GetElement(type_id)
            except Exception:
                element_type = None

            type_id_text = safe_id(type_id)
            type_name = get_type_name(element_type) if element_type else u"<Type Not Found>"
            unique_type_ids.add(type_id_text)

            mark = get_string_parameter(element, BuiltInParameter.ALL_MODEL_MARK)
            comments = get_string_parameter(
                element, BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS
            )
            level_name = get_level_name(element)
            area_sqm = get_double_parameter(
                element, BuiltInParameter.HOST_AREA_COMPUTED, SQFT_TO_SQM
            )
            volume_cum = get_double_parameter(
                element, BuiltInParameter.HOST_VOLUME_COMPUTED, CUFT_TO_CUM
            )

            # Read the actual Revit solid geometry. X/Y/Z are model-axis extents;
            # they are not claimed to be oriented edge dimensions for irregular shapes.
            geometry_metrics = get_solid_geometry_metrics(element)
            geometry_x_m = geometry_metrics["x_extent_m"]
            geometry_y_m = geometry_metrics["y_extent_m"]
            geometry_z_m = geometry_metrics["z_extent_m"]

            # LocationCurve.Length is the preferred true path length for walls/curved walls.
            location_curve_length_m = get_location_curve_length_m(element)
            wall_unconnected_height_m = (
                get_wall_unconnected_height_m(element)
                if category_name == "Walls" else None
            )

            compound = get_compound_structure(element_type) if element_type else None
            layers = []
            vertically_compound = False
            total_thickness_mm = 0.0

            if compound is not None:
                try:
                    layers = list(compound.GetLayers())
                except Exception:
                    layers = []
                try:
                    vertically_compound = bool(compound.IsVerticallyCompound)
                except Exception:
                    vertically_compound = False

            if compound is not None and len(layers) > 0:
                elements_with_compound += 1
                category_compound_count += 1

                for layer_index, layer in enumerate(layers):
                    try:
                        thickness_mm = float(layer.Width) * FEET_TO_MM
                    except Exception:
                        thickness_mm = 0.0
                    total_thickness_mm += thickness_mm

                    material_id = None
                    try:
                        material_id = layer.MaterialId
                    except Exception:
                        pass

                    material_name = get_material_name(material_id)
                    material_id_text = safe_id(material_id)

                    try:
                        layer_function = as_text(layer.Function)
                    except Exception:
                        layer_function = u""

                    try:
                        is_core = bool(compound.IsCoreLayer(layer_index))
                    except Exception:
                        is_core = False

                    if category_name == "Walls":
                        layer_position = "Exterior to Interior"
                    else:
                        layer_position = "Top to Bottom"

                    layer_rows.append([
                        category_name,
                        element_id_text,
                        type_id_text,
                        type_name,
                        layer_index + 1,
                        layer_position,
                        material_name,
                        material_id_text,
                        layer_function,
                        round(thickness_mm, 3),
                        is_core,
                        vertically_compound,
                    ])
            else:
                elements_without_compound += 1
                category_no_compound_count += 1
                skipped_type_rows.append([
                    category_name,
                    element_id_text,
                    type_id_text,
                    type_name,
                    "No compound structure returned for this type",
                ])

            element_rows.append([
                category_name,
                element_id_text,
                type_id_text,
                type_name,
                level_name,
                mark,
                comments,
                round(area_sqm, 4) if area_sqm is not None else None,
                round(volume_cum, 4) if volume_cum is not None else None,
                round(location_curve_length_m, 4) if location_curve_length_m is not None else None,
                round(geometry_x_m, 4) if geometry_x_m is not None else None,
                round(geometry_y_m, 4) if geometry_y_m is not None else None,
                round(geometry_z_m, 4) if geometry_z_m is not None else None,
                round(wall_unconnected_height_m, 4) if wall_unconnected_height_m is not None else None,
                round(geometry_metrics["surface_area_sqm"], 4) if geometry_metrics["surface_area_sqm"] is not None else None,
                round(geometry_metrics["volume_cum"], 4) if geometry_metrics["volume_cum"] is not None else None,
                round(geometry_metrics["top_area_sqm"], 4) if geometry_metrics["top_area_sqm"] is not None else None,
                round(geometry_metrics["bottom_area_sqm"], 4) if geometry_metrics["bottom_area_sqm"] is not None else None,
                round(geometry_metrics["vertical_area_sqm"], 4) if geometry_metrics["vertical_area_sqm"] is not None else None,
                round(geometry_metrics["edge_length_m"], 4) if geometry_metrics["edge_length_m"] is not None else None,
                geometry_metrics["solid_count"],
                geometry_metrics["edge_count"],
                geometry_metrics["source"],
                len(layers),
                round(total_thickness_mm, 3) if layers else None,
                vertically_compound,
            ])

        category_stats.append([
            category_name,
            category_count,
            category_compound_count,
            category_no_compound_count,
        ])

    return {
        "element_rows": element_rows,
        "layer_rows": layer_rows,
        "category_stats": category_stats,
        "skipped_type_rows": skipped_type_rows,
        "total_instances": total_instances,
        "elements_with_compound": elements_with_compound,
        "elements_without_compound": elements_without_compound,
        "unique_type_count": len(unique_type_ids),
    }


# -----------------------------------------------------------------------------
# Minimal XLSX writer (Open XML). No openpyxl / Excel installation required.
# -----------------------------------------------------------------------------


def column_letter(column_number):
    result = u""
    number = int(column_number)
    while number > 0:
        number, remainder = divmod(number - 1, 26)
        result = unichr(65 + remainder) + result
    return result


def cell_xml(cell_reference, value, header=False):
    if value is None or value == u"":
        return u""

    style = u' s="1"' if header else u""

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
    pieces = [
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        u'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
        u'<dimension ref="A1:{0}"/>'.format(last_cell),
        u'<sheetViews><sheetView workbookViewId="0">'
        u'<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        u'</sheetView></sheetViews>',
        u'<sheetFormatPr defaultRowHeight="18"/>',
        u'<sheetData>',
    ]

    for row_index, row in enumerate(rows, 1):
        pieces.append(u'<row r="{0}">'.format(row_index))
        for column_index, value in enumerate(row, 1):
            pieces.append(cell_xml(
                u"{0}{1}".format(column_letter(column_index), row_index),
                value,
                header=(row_index == 1),
            ))
        pieces.append(u'</row>')

    pieces.append(u'</sheetData>')
    if rows and max_columns:
        pieces.append(u'<autoFilter ref="A1:{0}"/>'.format(last_cell))
    pieces.append(u'</worksheet>')
    return u"".join(pieces)


def write_xlsx(file_path, sheets):
    """Write a multi-sheet XLSX package from [(sheet_name, rows), ...]."""
    content_types = [
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        u'<Default Extension="xml" ContentType="application/xml"/>',
        u'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        u'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    ]
    for index in range(len(sheets)):
        content_types.append(
            u'<Override PartName="/xl/worksheets/sheet{0}.xml" '
            u'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'.format(index + 1)
        )
    content_types.append(u'</Types>')

    root_rels = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        u'</Relationships>'
    )

    workbook_parts = [
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        u'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        u'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
        u'<bookViews><workbookView/></bookViews><sheets>',
    ]
    workbook_rels = [
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for index, (sheet_name, unused_rows) in enumerate(sheets, 1):
        workbook_parts.append(
            u'<sheet name="{0}" sheetId="{1}" r:id="rId{1}"/>'.format(
                clean_xml_text(sheet_name), index
            )
        )
        workbook_rels.append(
            u'<Relationship Id="rId{0}" '
            u'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            u'Target="worksheets/sheet{0}.xml"/>'.format(index)
        )
    workbook_parts.append(u'</sheets><calcPr calcId="191029"/></workbook>')
    workbook_rels.append(
        u'<Relationship Id="rId{0}" '
        u'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
        u'Target="styles.xml"/>'.format(len(sheets) + 1)
    )
    workbook_rels.append(u'</Relationships>')

    styles_xml = (
        u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        u'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        u'<fonts count="2">'
        u'<font><sz val="11"/><name val="Calibri"/></font>'
        u'<font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>'
        u'</fonts>'
        u'<fills count="3">'
        u'<fill><patternFill patternType="none"/></fill>'
        u'<fill><patternFill patternType="gray125"/></fill>'
        u'<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>'
        u'</fills>'
        u'<borders count="2">'
        u'<border><left/><right/><top/><bottom/><diagonal/></border>'
        u'<border><left style="thin"><color rgb="FFD9E2F3"/></left>'
        u'<right style="thin"><color rgb="FFD9E2F3"/></right>'
        u'<top style="thin"><color rgb="FFD9E2F3"/></top>'
        u'<bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border>'
        u'</borders>'
        u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
        u'<cellXfs count="2">'
        u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
        u'<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf>'
        u'</cellXfs>'
        u'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
        u'<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>'
        u'</styleSheet>'
    )

    archive = zipfile.ZipFile(file_path, "w", zipfile.ZIP_DEFLATED)
    try:
        parts = [
            ("[Content_Types].xml", u"".join(content_types)),
            ("_rels/.rels", root_rels),
            ("xl/workbook.xml", u"".join(workbook_parts)),
            ("xl/_rels/workbook.xml.rels", u"".join(workbook_rels)),
            ("xl/styles.xml", styles_xml),
        ]
        for index, (sheet_name, rows) in enumerate(sheets, 1):
            parts.append(("xl/worksheets/sheet{0}.xml".format(index), worksheet_xml(rows)))

        for part_name, part_content in parts:
            if not isinstance(part_content, bytes):
                part_content = part_content.encode("utf-8")
            archive.writestr(part_name, part_content)
    finally:
        archive.close()


# -----------------------------------------------------------------------------
# Build workbook sheets
# -----------------------------------------------------------------------------


def build_sheets(data):
    model_name = as_text(doc.Title)
    try:
        revit_version = as_text(doc.Application.VersionNumber)
    except Exception:
        revit_version = u"Unknown"

    summary_rows = [
        ["REVIT COMPOUND STRUCTURE EXPORT", ""],
        ["Model", model_name],
        ["Revit Version", revit_version],
        ["Export Date", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
        ["Categories scanned", "Walls, Floors, Roofs, Ceilings"],
        ["Total element instances", data["total_instances"]],
        ["Instances with compound layers", data["elements_with_compound"]],
        ["Instances without compound layers", data["elements_without_compound"]],
        ["Unique type IDs encountered", data["unique_type_count"]],
        ["", ""],
        ["CATEGORY", "ELEMENT COUNT", "WITH COMPOUND STRUCTURE", "WITHOUT COMPOUND STRUCTURE"],
    ]
    for row in data["category_stats"]:
        summary_rows.append(row)
    summary_rows.extend([
        ["", ""],
        ["IMPORTANT NOTES", ""],
        ["Layer order", "Walls: exterior to interior. Floors/Roofs/Ceilings: top to bottom."],
        ["Instance area", "Read from Revit HOST_AREA_COMPUTED where available; blank means not exposed."],
        ["Instance volume", "Read from Revit HOST_VOLUME_COMPUTED where available; blank means not exposed."],
        ["Location curve length", "Length of the element LocationCurve where Revit exposes one. This is the path length, including curved walls."],
        ["Geometry X/Y/Z extents", "Calculated from tessellated edges of Revit solid geometry in model coordinates. These are model-axis envelopes, not oriented length/width for irregular or rotated elements."],
        ["Geometry source", "Solid geometry is preferred. Element bounding box is used only as a clearly labelled fallback when no solid edge geometry is available."],
        ["Wall height", "Geometry Z extent is the wall's vertical envelope. The wall unconnected-height parameter is also exported when available; constrained/profiled walls may differ."],
        ["Solid geometry area/volume", "Calculated by summing Revit solid face areas and solid volumes. These may differ from host computed parameters due to joins, openings, and category calculation rules."],
        ["Top/bottom planar face area", "Areas of planar solid faces whose normal is nearly vertical (+Z or -Z). For sloped faces, these values are not projected plan area."],
        ["Solid edge length", "Sum of unique edges returned by each Revit solid. This is not a perimeter or a single element length."],
        ["Compound thickness", "Sum of compound layer widths; it is a type assembly thickness, not a geometry-derived BOQ quantity."],
        ["Layer quantities", "This export reports layer definitions and thickness, not calculated material quantities."],
        ["Vertically compound", "Flagged where the type has vertically varying wall regions; inspect the Revit type."],
        ["No compound structure", "Can include curtain, stacked, in-place, or other types without a standard compound structure."],
    ])

    elements_headers = [
        "Category", "Element ID", "Type ID", "Type Name", "Level", "Mark",
        "Comments", "Instance Area (m2)", "Instance Volume (m3)",
        "Location Curve Length (m)",
        "Solid Geometry X Extent (m)", "Solid Geometry Y Extent (m)",
        "Solid Geometry Z Extent (m)", "Wall Unconnected Height (m)",
        "Solid Geometry Surface Area (m2)", "Solid Geometry Volume (m3)",
        "Top-facing Planar Area (m2)", "Bottom-facing Planar Area (m2)",
        "Vertical Planar Face Area (m2)", "Sum of Solid Edge Lengths (m)",
        "Solid Count", "Solid Edge Count", "Geometry Source", "Layer Count",
        "Total Compound Layer Thickness (mm)", "Vertically Compound",
    ]
    elements_rows = [elements_headers] + data["element_rows"]

    layers_headers = [
        "Category", "Element ID", "Type ID", "Type Name", "Layer Number",
        "Layer Order", "Material", "Material ID", "Layer Function",
        "Thickness (mm)", "Is Core Layer", "Vertically Compound",
    ]
    layers_rows = [layers_headers] + data["layer_rows"]

    skipped_headers = [
        "Category", "Element ID", "Type ID", "Type Name", "Status / Note"
    ]
    skipped_rows = [skipped_headers] + data["skipped_type_rows"]

    return [
        ("Summary", summary_rows),
        ("Elements", elements_rows),
        ("Compound Layers", layers_rows),
        ("No Compound Structure", skipped_rows),
    ]


# -----------------------------------------------------------------------------
# Run export
# -----------------------------------------------------------------------------


def main():
    if doc is None:
        forms.alert("No Revit document is currently open.", title="Compound Structure Export")
        return

    forms.alert(
        "This tool will scan Walls, Floors, Roofs and Ceilings in the active model.\n\n"
        "It only reads model data and does not modify the Revit project.",
        title="Compound Structure Export",
    )

    data = collect_model_data()

    if data["total_instances"] == 0:
        forms.alert(
            "No Walls, Floors, Roofs or Ceilings were found in the active model.",
            title="Compound Structure Export",
        )
        return

    default_name = "{0}_Compound_Geometry_Export.xlsx".format(
        safe_filename(doc.Title)
    )
    output_path = forms.save_file(
        file_ext="xlsx",
        default_name=default_name,
        title="Save Compound Structure Excel Export",
    )

    if not output_path:
        return

    if not output_path.lower().endswith(".xlsx"):
        output_path += ".xlsx"

    try:
        write_xlsx(output_path, build_sheets(data))
    except Exception as export_error:
        forms.alert(
            "Excel export failed.\n\n{0}".format(as_text(export_error)),
            title="Compound Structure Export - Error",
        )
        raise

    summary = (
        "Excel export completed successfully.\n\n"
        "File: {0}\n\n"
        "Elements scanned: {1}\n"
        "Elements with compound layers: {2}\n"
        "Elements without compound layers: {3}\n"
        "Layer rows exported: {4}"
    ).format(
        output_path,
        data["total_instances"],
        data["elements_with_compound"],
        data["elements_without_compound"],
        len(data["layer_rows"]),
    )
    forms.alert(summary, title="Compound Structure Export")


main()
