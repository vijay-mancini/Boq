
from pyrevit import revit, forms

from lib.boq.extractor import extract_elements
from lib.boq.exporter import export_json


doc = revit.doc

if doc is None:

    forms.alert(
        "Please open a Revit project first.",
        title="BOQ Extractor"
    )

else:

    try:

        elements = extract_elements(doc)

        filepath = export_json(
            elements,
            doc.Title
        )

        forms.alert(
            "Extraction completed!\n\n"
            "Total Elements: {}\n\n"
            "Output: {}".format(
                len(elements),
                filepath
            ),
            title="BOQ Extractor"
        )

    except Exception as error:

        forms.alert(
            str(error),
            title="Extraction Error"
        )
