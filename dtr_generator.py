"""
DTR Helper - command-line entry point

Reads a time-records spreadsheet, groups sessions into weekly (Mon-Sat)
DTR forms, and produces:
  1. preview.txt        -> a human-readable dump of the parsed data (for review)
  2. DTR_output.docx     -> the filled Word DTR (one weekly form per week, page-broken)

Run:
    pip install -r requirements.txt
    python dtr_generator.py

You can also pass a specific spreadsheet path:
    python dtr_generator.py "C:\\path\\to\\Sample DTR.xlsx"

For the web version, see run_web.py.
"""

import os
import sys
import datetime as dt

from dtr import engine
from dtr.defaults import DEFAULTS

HERE = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")


def find_file(name):
    for folder in (HERE, DOWNLOADS):
        p = os.path.join(folder, name)
        if os.path.exists(p):
            return p
    return None


def resolve_inputs():
    xlsx = sys.argv[1] if len(sys.argv) > 1 else find_file("Sample DTR.xlsx")
    docx = find_file("EMPTY_DTR.docx")
    return xlsx, docx


def main():
    xlsx, docx = resolve_inputs()
    print("Spreadsheet:", xlsx)
    print("Template:   ", docx)
    if not xlsx or not os.path.exists(xlsx):
        print("ERROR: spreadsheet not found.")
        return

    try:
        name, entries, anomalies = engine.parse_workbook(xlsx)
    except engine.WorkbookReadError as exc:
        print("ERROR:", exc)
        return

    weeks = engine.group_into_weeks(entries)

    preview_path = os.path.join(HERE, "preview.txt")
    engine.write_preview(name, weeks, anomalies, preview_path)
    print("Wrote", preview_path, f"({len(entries)} entries, {len(weeks)} weeks)")

    if not weeks:
        print("No usable Monday-Saturday entries were found; skipped Word generation.")
        print("See preview.txt for details (anomalies, Sunday-only entries, etc.).")
        return

    ident = dict(DEFAULTS)
    ident["name"] = name

    if docx and os.path.exists(docx):
        try:
            out_path = os.path.join(HERE, "DTR_output.docx")
            try:
                engine.generate_docx(docx, ident, weeks, out_path)
            except PermissionError:
                stamp = dt.datetime.now().strftime("%H%M%S")
                out_path = os.path.join(HERE, f"DTR_output_{stamp}.docx")
                engine.generate_docx(docx, ident, weeks, out_path)
                print("(DTR_output.docx was open/locked; saved a new copy.)")
            print("Wrote", out_path)
        except Exception:
            import traceback
            err_path = os.path.join(HERE, "docx_error.txt")
            with open(err_path, "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
            print("DOCX generation FAILED - see", err_path)
    else:
        print("Template EMPTY_DTR.docx not found; skipped Word generation.")


if __name__ == "__main__":
    main()
