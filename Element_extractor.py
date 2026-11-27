
# -*- coding: utf-8 -*-
# pyRevit - Extract All Model Elements
# Exports two CSV files:
# 1. Elements.csv
# 2. ElementParameters.csv

from pyrevit import revit, forms, script
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory,
    Element,
    Material
)
import System
import os
import csv

doc = revit.doc
output = script.get_output()

folder = forms.select_folder(
    title="Select folder to save Revit extraction"
)

if not folder:
    forms.alert("Export cancelled.")
    script.exit()

elements_file = os.path.join(folder, "Elements.csv")
parameters_file = os.path.join(folder, "ElementParameters.csv")


def safe_string(value):
    try:
        if value is None:
            return ""
        return str(value)
    except:
        return ""


def get_category(elem):
    try:
        if elem.Category:
            return elem.Category.Name
    except:
        pass
    return ""


def get_family(elem):
    try:
        if elem.Symbol and elem.Symbol.Family:
            return elem.Symbol.Family.Name
    except:
        pass

    try:
        if elem.GetType():
            return elem.GetType().Name
    except:
        pass

    return ""


def get_type_name(elem):
    try:
        type_id = elem.GetTypeId()

        if type_id and type_id != System.DB.ElementId.InvalidElementId:
            type_elem = doc.GetElement(type_id)

            if type_elem:
                try:
                    return type_elem.Name
                except:
                    pass
    except:
        pass

    return ""


def get_level(elem):
    try:
        level_param = elem.LookupParameter("Level")

        if level_param:
            if level_param.AsElementId():
                level_elem = doc.GetElement(level_param.AsElementId())

                if level_elem:
                    return level_elem.Name
    except:
        pass

    try:
        level_id = elem.LevelId

        if level_id and level_id != System.DB.ElementId.InvalidElementId:
            level = doc.GetElement(level_id)

            if level:
                return level.Name
    except:
        pass

    return ""


def get_workset(elem):
    try:
        if elem.WorksetId:
            workset = doc.GetWorksetTable().GetWorkset(elem.WorksetId)
            if workset:
                return workset.Name
    except:
        pass

    return ""


def get_mark(elem):
    try:
        param = elem.LookupParameter("Mark")

        if param:
            return param.AsString() or param.AsValueString() or ""
    except:
        pass

    return ""


def get_materials(elem):
    names = []
    material_volumes = []

    try:
        material_ids = elem.GetMaterialIds(False)

        for material_id in material_ids:
            material = doc.GetElement(material_id)

            if isinstance(material, Material):
                names.append(material.Name)

                try:
                    volume = elem.GetMaterialVolume(material_id)
                    material_volumes.append(
                        "{}: {}".format(material.Name, volume)
                    )
                except:
                    pass

    except:
        pass

    return "; ".join(names), "; ".join(material_volumes)


# Collect model elements excluding element types
collector = FilteredElementCollector(doc).WhereElementIsNotElementType()

elements = list(collector)

element_rows = []
parameter_rows = []

total = len(elements)

output.print_md("# Revit Model Extraction")
output.print_md("Total elements found: **{}**".format(total))

for index, elem in enumerate(elements):

    try:
        element_id = elem.Id.IntegerValue
    except:
        try:
            element_id = elem.Id.Value
        except:
            element_id = ""

    try:
        element_name = elem.Name
    except:
        element_name = ""

    category = get_category(elem)
    family = get_family(elem)
    type_name = get_type_name(elem)
    level_name = get_level(elem)
    workset_name = get_workset(elem)
    mark = get_mark(elem)

    material_names, material_volumes = get_materials(elem)

    element_rows.append([
        element_id,
        category,
        family,
        type_name,
        element_name,
        level_name,
        workset_name,
        mark,
        material_names,
        material_volumes
    ])

    # Extract all accessible parameters
    try:
        parameters = elem.Parameters

        for param in parameters:

            try:
                param_name = param.Definition.Name
            except:
                param_name = ""

            try:
                param_value = param.AsValueString()
            except:
                param_value = ""

            if not param_value:
                try:
                    if param.StorageType.ToString() == "String":
                        param_value = param.AsString()
                    elif param.StorageType.ToString() == "Integer":
                        param_value = param.AsInteger()
                    elif param.StorageType.ToString() == "Double":
                        param_value = param.AsDouble()
                    elif param.StorageType.ToString() == "ElementId":
                        param_value = param.AsElementId().IntegerValue
                except:
                    pass

            try:
                parameter_type = param.StorageType.ToString()
            except:
                parameter_type = ""

            parameter_rows.append([
                element_id,
                category,
                element_name,
                param_name,
                safe_string(param_value),
                parameter_type
            ])

    except:
        pass

    if index % 500 == 0:
        output.print_md(
            "Processed **{} / {}** elements".format(index, total)
        )


# Write main element file
with open(elements_file, "wb") as f:
    writer = csv.writer(f)

    writer.writerow([
        "Element ID",
        "Category",
        "Family",
        "Type",
        "Name",
        "Level",
        "Workset",
        "Mark",
        "Materials",
        "Material Volumes"
    ])

    writer.writerows(element_rows)


# Write parameter file
with open(parameters_file, "wb") as f:
    writer = csv.writer(f)

    writer.writerow([
        "Element ID",
        "Category",
        "Element Name",
        "Parameter Name",
        "Parameter Value",
        "Storage Type"
    ])

    writer.writerows(parameter_rows)


forms.alert(
    "Extraction completed!\n\n"
    "Elements extracted: {}\n\n"
    "Files saved:\n{}\n{}".format(
        total,
        elements_file,
        parameters_file
    ),
    title="Revit Extraction Complete"
)
