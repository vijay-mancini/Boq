
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInCategory
)


ARCHITECTURAL_CATEGORIES = {
    "Walls": BuiltInCategory.OST_Walls,
    "Floors": BuiltInCategory.OST_Floors,
    "Doors": BuiltInCategory.OST_Doors,
    "Windows": BuiltInCategory.OST_Windows
}


def get_element_type(doc, element):
    type_id = element.GetTypeId()

    if type_id is None:
        return None

    return doc.GetElement(type_id)


def extract_elements(doc):

    extracted_data = []

    for category_name, category in ARCHITECTURAL_CATEGORIES.items():

        collector = (
            FilteredElementCollector(doc)
            .OfCategory(category)
            .WhereElementIsNotElementType()
        )

        for element in collector:

            element_type = get_element_type(doc, element)

            level = None

            try:
                level_id = element.LevelId

                if level_id.IntegerValue != -1:
                    level_obj = doc.GetElement(level_id)
                    level = level_obj.Name if level_obj else None

            except Exception:
                pass

            element_data = {
                "element_id": element.Id.Value,
                "unique_id": element.UniqueId,
                "category": category_name,
                "name": element.Name,
                "type_id": (
                    element_type.Id.Value
                    if element_type else None
                ),
                "type_name": (
                    element_type.Name
                    if element_type else None
                ),
                "level": level
            }

            extracted_data.append(element_data)

    return extracted_data
