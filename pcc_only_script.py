# -*- coding: utf-8 -*-
# pyRevit PCC Material Extractor
# Extracts PCC-related material quantities from the entire Revit model.

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

import os
import clr

clr.AddReference("System")
from System.Collections.Generic import List

doc = revit.doc
output = script.get_output()

# --------------------------------------------------
# KEYWORDS
# --------------------------------------------------

KEYWORDS = [
    "pcc",
    "plain cement concrete",
    "lean concrete",
    "cement concrete",
    "m10",
    "m15"
]

# --------------------------------------------------
# MATERIAL FILTER
# --------------------------------------------------

def is_pcc_material(material_name):

    if not material_name:
        return False

    name = material_name.lower().strip()

    for keyword in KEYWORDS:
        if keyword in name:
            return True

    return False


# --------------------------------------------------
# UNIT CONVERSION
# --------------------------------------------------

def volume_to_m3(volume_internal):

    try:
        return UnitUtils.ConvertFromInternalUnits(
            volume_internal,
            UnitTypeId.CubicMeters
        )
    except:
        return 0.0


def area_to_m2(area_internal):

    try:
        return UnitUtils.ConvertFromInternalUnits(
            area_internal,
            UnitTypeId.SquareMeters
        )
    except:
        return 0.0


# --------------------------------------------------
# ELEMENT NAME
# --------------------------------------------------

def get_element_name(element):

    try:
        return element.Name
    except:
        try:
            return element.get_Name()
        except:
            return ""


def get_category_name(element):

    try:
        return element.Category.Name
    except:
        return "Unknown"


# --------------------------------------------------
# MATERIAL EXTRACTION
# --------------------------------------------------

elements = FilteredElementCollector(doc)\
    .WhereElementIsNotElementType()\
    .ToElements()

detailed_data = []
summary_data = {}

processed_count = 0

for element in elements:

    try:

        material_ids = element.GetMaterialIds(False)

        for material_id in material_ids:

            material = doc.GetElement(material_id)

            if not material:
                continue

            material_name = material.Name

            if not is_pcc_material(material_name):
                continue

            processed_count += 1

            try:
                volume_internal = element.GetMaterialVolume(material_id)
            except:
                volume_internal = 0.0

            try:
                area_internal = element.GetMaterialArea(
                    material_id,
                    False
                )
            except:
                area_internal = 0.0

            volume_m3 = volume_to_m3(volume_internal)
            area_m2 = area_to_m2(area_internal)

            element_id = str(element.Id.IntegerValue)
            category = get_category_name(element)
            element_name = get_element_name(element)

            detailed_data.append([
                element_id,
                category,
                element_name,
                material_name,
                area_m2,
                volume_m3
            ])

            # Summary by material
            if material_name not in summary_data:

                summary_data[material_name] = {
                    "area": 0.0,
                    "volume": 0.0,
                    "count": 0
                }

            summary_data[material_name]["area"] += area_m2
            summary_data[material_name]["volume"] += volume_m3
            summary_data[material_name]["count"] += 1

    except Exception:
        continue


# --------------------------------------------------
# CHECK RESULTS
# --------------------------------------------------

if not detailed_data:

    forms.alert(
        "No PCC-related materials were found.\n\n"
        "You can modify the keywords in the script.",
        title="PCC Extractor"
    )

    script.exit()


# --------------------------------------------------
# EXCEL EXPORT
# --------------------------------------------------

try:

    import openpyxl

    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter

    wb = openpyxl.Workbook()

    # DETAIL SHEET
    ws = wb.active
    ws.title = "PCC Detailed Data"

    headers = [
        "Element ID",
        "Category",
        "Element Name",
        "Material Name",
        "Area (m2)",
        "Volume (m3)"
    ]

    ws.append(headers)

    for row in detailed_data:
        ws.append(row)

    # SUMMARY SHEET
    ws2 = wb.create_sheet("PCC Summary")

    ws2.append([
        "Material Name",
        "Element Count",
        "Total Area (m2)",
        "Total Volume (m3)"
    ])

    for material_name, values in sorted(summary_data.items()):

        ws2.append([
            material_name,
            values["count"],
            values["area"],
            values["volume"]
        ])

    # FORMATTING
    for sheet in [ws, ws2]:

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        for cell in sheet[1]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="1F4E78"
            )

        for column in sheet.columns:

            max_length = 0

            for cell in column:

                try:
                    max_length = max(
                        max_length,
                        len(str(cell.value))
                    )
                except:
                    pass

            sheet.column_dimensions[
                get_column_letter(column[0].column)
            ].width = min(max_length + 3, 40)

    # --------------------------------------------------
    # SAVE FILE
    # --------------------------------------------------

    desktop = os.path.join(
        os.path.expanduser("~"),
        "Desktop"
    )

    if not os.path.exists(desktop):
        desktop = os.path.expanduser("~")

    file_path = os.path.join(
        desktop,
        "PCC_Material_Extraction.xlsx"
    )

    wb.save(file_path)

    forms.alert(
        "PCC extraction completed!\n\n"
        "Elements processed: {}\n"
        "Materials found: {}\n\n"
        "File saved at:\n{}".format(
            len(detailed_data),
            len(summary_data),
            file_path
        ),
        title="PCC Extractor"
    )

except Exception as ex:

    forms.alert(
        "Excel export failed:\n{}".format(str(ex)),
        title="Export Error"
    )
