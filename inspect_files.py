"""
Standalone inspector for the DTR template (.docx) and the sample time-records (.xlsx).
Uses ONLY the Python standard library, so no pip install is needed.

Run it from the project folder:

    python inspect_files.py

It writes everything to  inspect_output.txt  in the same folder.
"""

import zipfile
import xml.etree.ElementTree as ET
import datetime
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DOWNLOADS = os.path.join(os.path.expanduser("~"), "Downloads")


def find_file(name):
    for folder in (HERE, DOWNLOADS):
        p = os.path.join(folder, name)
        if os.path.exists(p):
            return p
    return os.path.join(HERE, name)


DOCX = find_file("EMPTY_DTR.docx")
XLSX = find_file("Sample DTR.xlsx")

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

out_lines = []


def log(*args):
    out_lines.append(" ".join(str(a) for a in args))


# ---------------------------------------------------------------- DOCX ----
def para_text(p):
    parts = []
    for t in p.iter(W + "t"):
        parts.append(t.text or "")
    return "".join(parts)


def walk_docx():
    log("=" * 70)
    log("DOCX:", DOCX, "exists=", os.path.exists(DOCX))
    log("=" * 70)
    if not os.path.exists(DOCX):
        log("!! EMPTY_DTR.docx not found in", HERE)
        return
    with zipfile.ZipFile(DOCX) as z:
        log("-- files inside docx --")
        for n in z.namelist():
            log("  ", n)
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    body = root.find(W + "body")

    def render(el, depth=0):
        pad = "  " * depth
        for child in el:
            tag = child.tag.replace(W, "")
            if tag == "p":
                txt = para_text(child)
                log(f"{pad}P: {txt!r}")
            elif tag == "tbl":
                log(f"{pad}[TABLE]")
                for ri, row in enumerate(child.findall(W + "tr")):
                    cells = []
                    for cell in row.findall(W + "tc"):
                        ctext = " ".join(
                            para_text(p) for p in cell.findall(W + "p")
                        )
                        cells.append(ctext.strip())
                    log(f"{pad}  row {ri}: " + " | ".join(f"[{c}]" for c in cells))
                log(f"{pad}[/TABLE]")
            elif tag == "sectPr":
                continue
            else:
                render(child, depth)

    render(body)


# ---------------------------------------------------------------- XLSX ----
def col_letters(cellref):
    letters = ""
    for ch in cellref:
        if ch.isalpha():
            letters += ch
        else:
            break
    return letters


def excel_serial_to_date(n):
    try:
        n = float(n)
    except (TypeError, ValueError):
        return None
    # Excel's day 0 = 1899-12-30 (accounting for the 1900 leap bug)
    base = datetime.datetime(1899, 12, 30)
    try:
        return base + datetime.timedelta(days=n)
    except OverflowError:
        return None


def walk_xlsx():
    log("")
    log("=" * 70)
    log("XLSX:", XLSX, "exists=", os.path.exists(XLSX))
    log("=" * 70)
    if not os.path.exists(XLSX):
        log("!! Sample DTR.xlsx not found in", HERE)
        return
    with zipfile.ZipFile(XLSX) as z:
        names = z.namelist()
        log("-- files inside xlsx --")
        for n in names:
            log("  ", n)

        shared = []
        if "xl/sharedStrings.xml" in names:
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sroot.findall(S + "si"):
                txt = "".join(t.text or "" for t in si.iter(S + "t"))
                shared.append(txt)

        # map sheet names
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        log("-- sheets --")
        for sh in wb.iter(S + "sheet"):
            log("  ", sh.attrib)

        sheet_files = [n for n in names if n.startswith("xl/worksheets/sheet")]
        for sf in sheet_files:
            log("")
            log(f"###### {sf} ######")
            sroot = ET.fromstring(z.read(sf))
            data = sroot.find(S + "sheetData")
            if data is None:
                continue
            for row in list(data.findall(S + "row"))[:45]:
                rnum = row.attrib.get("r", "?")
                cells_out = []
                for c in row.findall(S + "c"):
                    ref = c.attrib.get("r", "")
                    col = col_letters(ref)
                    ctype = c.attrib.get("t", "")
                    v = c.find(S + "v")
                    raw = v.text if v is not None else None
                    if ctype == "s" and raw is not None:
                        val = shared[int(raw)]
                        disp = f"{col}={val!r}"
                    elif raw is not None:
                        # numeric: could be a date/time serial
                        dt = excel_serial_to_date(raw)
                        extra = ""
                        try:
                            f = float(raw)
                            if 0 < f < 60000:
                                if dt:
                                    if f < 1:  # time only
                                        secs = int(round(f * 86400))
                                        hh = secs // 3600
                                        mm = (secs % 3600) // 60
                                        extra = f" (time~{hh:02d}:{mm:02d})"
                                    else:
                                        extra = f" (date~{dt.strftime('%Y-%m-%d %H:%M')} = {dt.strftime('%A')})"
                        except ValueError:
                            pass
                        disp = f"{col}={raw}{extra}"
                    else:
                        disp = f"{col}=<empty>"
                    cells_out.append(disp)
                log(f"  row {rnum}: " + "  ".join(cells_out))


def main():
    walk_docx()
    walk_xlsx()
    text = "\n".join(out_lines)
    outpath = os.path.join(HERE, "inspect_output.txt")
    with open(outpath, "w", encoding="utf-8") as f:
        f.write(text)
    print("Wrote", outpath)
    print("---- preview ----")
    print(text[:2000])


if __name__ == "__main__":
    main()
