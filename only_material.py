# -*- coding: utf-8 -*-
# Generic Material Extractor
# pyRevit | IronPython compatible

from Autodesk.Revit.DB import *
from pyrevit import revit, forms, script

import csv

doc = revit.doc

# --------------------------------------------------
# USER INPUT
# --------------------------------------------------

keyword = forms.ask_for_string(
    prompt="Enter material name or keyword.\n"
           "You can enter multiple keywords separated by commas.",
    title="Generic Material Extractor",
    default="PCC"
)

if not keyword:
    script.exit()

keywords = [
    x.strip().lower()
    for x in keyword.split(",")
    if x.strip()
]

# --------------------------------------------------
# MATERIAL MATCHING
# --------------------------------------------------

def material_matches(name):

    if not name:
        return False

    name = name.lower()

    for keyword in keywords:
        if keyword in name:
            return True

    return False


# --------------------------------------------------
# UNIT CONVERSION
# --------------------------------------------------

def convert_volume(value):

    try:
        return UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.CubicMeters
        )
    except:
        return 0.0


def convert_area(value):

    try:
        return UnitUtils.ConvertFromInternalUnits(
            value,
            UnitTypeId.SquareMeters
        )
    except:
        return 0.0


# --------------------------------------------------
# ELEMENT INFORMATION
# --------------------------------------------------

def get_category(element):

    try:
        return element.Category.Name
    except:
        return "Unknown"


def get_element_name(element):

    try:
        return element.Name
    except:
        return ""


def get_material_volume(element, material_id):

    try:
        return convert_volume(
            element.GetMaterialVolume(material_id)
        )
    except:
        return 0.0


def get_material_area(element, material_id):

    try:
        return convert_area(
            element.GetMaterialArea(material_id, False)
        )
    except:
        return 0.0


# --------------------------------------------------
# COLLECT ELEMENTS
# --------------------------------------------------

elements = FilteredElementCollector(doc)\
    .WhereElementIsNotElementType()\
    .ToElements()

results = []

matched_materials = set()
matched_elements = set()

# --------------------------------------------------
# EXTRACT MATERIALS
# --------------------------------------------------

for element in elements:

    try:

        material_ids = element.GetMaterialIds(False)

        for material_id in material_ids:

            material = doc.GetElement(material_id)

            if material is None:
                continue

            material_name = material.Name

            if not material_matches(material_name):
                continue

            element_id = element.Id.IntegerValue

            area = get_material_area(
                element,
                material_id
            )

            volume = get_material_volume(
                element,
                material_id
            )

            results.append([
                element_id,
                get_category(element),
                get_element_name(element),
                material_name,
                area,
                volume
            ])

            matched_materials.add(material_name)
            matched_elements.add(element_id)

    except:
        continue


# --------------------------------------------------
# RESULT CHECK
# --------------------------------------------------

if not results:

    forms.alert(
        "No matching materials were found.\n\n"
        "Search keywords: {}".format(
            ", ".join(keywords)
        ),
        title="Material Extractor"
    )

    script.exit()


# --------------------------------------------------
# EXPORT CSV
# --------------------------------------------------

file_path = forms.save_file(
    file_ext="csv",
    default_name="Material_Extraction.csv",
    title="Save Material Extraction"
)

if not file_path:
    script.exit()


try:

    with open(file_path, "wb") as csv_file:

        writer = csv.writer(csv_file)

        writer.writerow([
            "Element ID",
            "Category",
            "Element Name",
            "Material Name",
            "Area (m2)",
            "Volume (m3)"
        ])

        for row in results:
            writer.writerow(row)

    # --------------------------------------------------
    # COMPLETION
    # --------------------------------------------------

    forms.alert(
        "Extraction completed successfully!\n\n"
        "Matching materials: {}\n"
        "Matching elements: {}\n"
        "Extracted rows: {}\n\n"
        "Saved at:\n{}".format(
            len(matched_materials),
            len(matched_elements),
            len(results),
            file_path
        ),
        title="Material Extractor"
    )

except Exception as ex:

    forms.alert(
        "Export error:\n{}".format(str(ex)),
        title="Material Extractor"
    )
