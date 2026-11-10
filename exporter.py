
import json
import os
from datetime import datetime


def export_json(data, project_name):

    desktop = os.path.join(
        os.path.expanduser("~"),
        "Desktop"
    )

    filename = "BOQ_{}.json".format(
        datetime.now().strftime("%Y%m%d_%H%M%S")
    )

    filepath = os.path.join(desktop, filename)

    output = {
        "project_name": project_name,
        "generated_at": datetime.now().isoformat(),
        "revit_version": "2026",
        "elements": data
    }

    with open(filepath, "w") as file:
        json.dump(
            output,
            file,
            indent=4,
            ensure_ascii=False
        )

    return filepath
