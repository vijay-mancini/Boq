# -*- coding: utf-8 -*-
"""Collects every BOQ page of the active project into one export file and runs boq_export.py (Python 3) to build
the Excel (and PDF if reportlab is installed). Nothing is written to the Revit model."""
import os
import json
import codecs
import datetime
import clr
clr.AddReference("System.Diagnostics.Process")      # Revit 2025+/2026 (.NET 8+): Process lives in its own assembly
clr.AddReference("System.ComponentModel.Primitives")
from System.Diagnostics import Process, ProcessStartInfo

from boq_service import ArchitectureBOQService
from boq_repository import ArchitectureBOQRepository
from data_manager import APP_DATA

BASE_DIR = os.path.dirname(__file__)
PYTHON_CANDIDATES = [
    r"C:\Users\Admin\AppData\Local\Programs\Python\Python312\python.exe",
    r"C:\Python312\python.exe",
    r"C:\Program Files\Python312\python.exe",
]
SECTIONS = [("preliminary", u"Preliminary Work"), ("architecture", u"Architecture"), ("interior", u"Interior"), ("structure", u"Structure"), ("landscape", u"Landscape")]


def _find_python():
    for p in PYTHON_CANDIDATES:
        if os.path.exists(p):
            return p
    for folder in (os.environ.get("PATH") or "").split(";"):
        p = os.path.join(folder.strip('"'), "python.exe")
        if folder and os.path.exists(p) and "WindowsApps" not in p:
            return p
    return None



def _run_book_extraction():
    """Run book.py-style read-only extraction in the active Revit context."""
    try:
        import __main__
        doc = __main__.__dict__.get("doc")
        DB = __main__.__dict__.get("DB")
    except Exception:
        doc = DB = None
    # pyRevit scripts commonly expose doc/DB in script globals, not __main__.
    if doc is None or DB is None:
        try:
            from Autodesk.Revit import DB as _DB
            from pyrevit import revit
            doc, DB = revit.doc, _DB
        except Exception as ex:
            raise ValueError("Could not access active Revit document/API for book extraction: " + str(ex))
    namespace = {"doc": doc, "DB": DB, "PROJECT_DIR": os.path.join(BASE_DIR, "book_ref_detail")}
    extractor = os.path.join(BASE_DIR, "boq2_book_extractor.py")
    if not os.path.exists(extractor):
        raise ValueError("Book extraction helper not found: " + extractor)
    # IronPython 2.7 in pyRevit supports execfile with an injected namespace.
    execfile(extractor, namespace)
    return namespace.get("_snap")


def collect():
    """Plain dict with project, client and the items of every page for the active Revit project."""
    from revit_reader import project_key
    key, title = project_key()
    sections = []
    for disc, name in SECTIONS:
        service = ArchitectureBOQService(ArchitectureBOQRepository(key, disc))
        items = service.list_items()
        for it in items:
            it.pop("model_type_info", None)
            it.pop("model_details", None)
        sections.append({"name": name, "items": items})
    data = {
        "revit_project": title,
        "project": dict(APP_DATA.get("project", {})),
        "client": dict(APP_DATA.get("client", {})),
        "date": datetime.date.today().strftime("%d %b %Y"),
        "sections": sections,
    }
    # Additional read-only snapshot from book.py logic; existing BOQ section items remain unchanged.
    try:
        data["book_extraction"] = _run_book_extraction()
    except Exception as ex:
        data["book_extraction_error"] = str(ex)
    return data


def export_all():
    """Returns (list of written files, message). Raises ValueError with a readable text when nothing can be exported."""
    data = collect()
    if not any(s["items"] for s in data["sections"]):
        raise ValueError("There are no BOQ items yet. Open Architecture, Interior, Structure or Landscape (they load from the model) or tick topics in Preliminary Work first.")
    py = _find_python()
    if py is None:
        raise ValueError("Python 3 was not found. Edit PYTHON_CANDIDATES in exporter.py (Python 3 with openpyxl is required).")
    folder = os.path.join(os.path.expanduser("~"), "Documents", "Mancini BOQ")
    if not os.path.isdir(folder):
        os.makedirs(folder)
    src = os.path.join(folder, "export_data.json")
    with codecs.open(src, "w", "utf-8") as f:
        f.write(json.dumps(data, ensure_ascii=False, indent=1))
    psi = ProcessStartInfo()
    psi.FileName = py
    psi.Arguments = u'-X utf8 "{0}" "{1}" "{2}"'.format(os.path.join(BASE_DIR, "boq_export.py"), src, folder)
    psi.UseShellExecute = False
    psi.CreateNoWindow = True
    psi.RedirectStandardOutput = True
    psi.RedirectStandardError = True
    p = Process.Start(psi)
    out = p.StandardOutput.ReadToEnd()
    err = p.StandardError.ReadToEnd()
    p.WaitForExit()
    if p.ExitCode != 0:
        raise ValueError(u"Export failed (is openpyxl installed? python -m pip install openpyxl):\n" + (err or out)[-900:])
    files = [l[len("written "):].strip() for l in out.splitlines() if l.startswith("written ")]
    notes = [l for l in out.splitlines() if l.startswith("note:")]
    return files, u"\n".join(notes)
