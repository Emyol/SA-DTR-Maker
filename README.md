# DTR Helper

Automatically fills the **Student Assistant Daily Time Record (DTR)** Word form
from a simple spreadsheet of time logs.

You paste your time-in / time-out entries into a spreadsheet, run one command,
and the tool produces a ready-to-print Word document with one weekly (Mon–Sat)
DTR form per week, two forms per page — with the daily hours, weekly totals, and
"Assigned Work" all filled in for you.

---

## Table of contents

- [What it does](#what-it-does)
- [Repository contents](#repository-contents)
- [Requirements](#requirements)
- [Installation](#installation)
- [Preparing your spreadsheet](#preparing-your-spreadsheet) ← **most important**
- [Running the tool](#running-the-tool)
- [Output files](#output-files)
- [How the DTR gets filled (the rules)](#how-the-dtr-gets-filled-the-rules)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [Notes for developers (turning this into a website)](#notes-for-developers-turning-this-into-a-website)

---

## What it does

- Reads a spreadsheet with two time logs: **ONLINE** (virtual duty) and **ONSITE** (onsite duty).
- Figures out the **exact weekday** of every date automatically.
- Groups all entries into **weekly forms (Monday–Saturday)** — one form per week that has records.
- Computes **hours rendered per session**, and the **weekly total**.
- Handles **multiple sessions on the same day** (they stack inside the same cell).
- Fills the **"Total Hours Rendered"** box at the top-right of each weekly form.
- Outputs a Word file that matches the official template exactly (two forms per page).

---

## Repository contents

| File | Purpose |
| --- | --- |
| `dtr_generator.py` | The main program. |
| `EMPTY_DTR.docx` | The official blank DTR template that gets filled. **Required.** |
| `requirements.txt` | The Python libraries needed. |
| `Sample DTR.xlsx` | An example time-logs spreadsheet you can copy and edit. |
| `inspect_files.py` | Optional helper for inspecting the template/spreadsheet structure. |
| `README.md` | This file. |

> ⚠️ **`EMPTY_DTR.docx` must stay in the same folder** as `dtr_generator.py`.
> It is the template the tool fills in.

---

## Requirements

- **Python 3.9 or newer** (developed on 3.13). Download from
  [python.org](https://www.python.org/downloads/) and, on Windows, tick
  **"Add Python to PATH"** during installation.
- Two Python libraries (installed in the next step):
  - `openpyxl` — reads the spreadsheet
  - `python-docx` — writes the Word document

---

## Installation

Open a terminal (PowerShell on Windows) **in the project folder**, then run:

```powershell
pip install -r requirements.txt
```

That's it — no database, no accounts, nothing else to set up.

---

## Preparing your spreadsheet

The tool reads the **first (active) sheet** of an `.xlsx` file. The layout below
matches the included `Sample DTR.xlsx` — the easiest approach is to **copy that
file and just replace the data**.

### The layout

There are **two side-by-side logs**. The left log is **ONLINE (virtual)** and the
right log is **ONSITE**. Each log has a Date, Time In, and Time Out column.

| Cell / Column | What goes there |
| --- | --- |
| **B2** | Your full name (e.g., `ACUÑA, AMIEL JOSIAH C.`) |
| **Row 3** | Section labels — `ONLINE` in column B, `ONSITE` in column F |
| **Row 4** | Column headers (`DATE (mm/dd/yy)`, `TIME IN`, `TIME OUT`, `DURATION`) |
| **Row 5 and below** | Your actual time entries |

**ONLINE (Virtual Office Duty) — left side, starting at row 5:**

| Column | Meaning |
| --- | --- |
| **B** | Date of the duty (a real date, e.g. `04/21/26`) |
| **C** | Time In (a real time, e.g. `9:00 AM`) |
| **D** | Time Out (a real time, e.g. `1:00 PM`) |
| **E** | Duration — *optional, ignored* (the tool computes hours itself) |

**ONSITE (Onsite Duty) — right side, starting at row 5:**

| Column | Meaning |
| --- | --- |
| **F** | Date of the duty |
| **G** | Time In |
| **H** | Time Out |
| **I** | Duration — *optional, ignored* |

### Rules for entering data

1. **Put virtual duties in the ONLINE columns (B–D)** and **onsite duties in the
   ONSITE columns (F–H).** The tool decides "Assigned Work" based on **which side
   you put the entry on**, not on the weekday:
   - Anything in the ONLINE columns → **"iCARE Virtual Office Duty"**
   - Anything in the ONSITE columns → **"iCARE Onsite Duty"**
2. **Dates must be real dates** and **times must be real times** (formatted as
   dates/times in Excel), not plain text. If you type `9:00 AM` Excel usually
   converts it automatically. The tool figures out the weekday from the date.
3. **The two sides are independent.** A row does not need matching ONLINE and
   ONSITE entries — leave a side blank if there was no duty there that row.
4. **Multiple sessions on one day are fine.** If you worked twice on the same
   date (e.g., 11:00 AM–3:00 PM and 4:00 PM–7:00 PM), just add two rows with the
   same date. They will stack in the same day's cell on the form.
5. **Sundays are not supported** — the DTR form only has Monday–Saturday rows.
   Any Sunday entry is flagged in `preview.txt` and left off the form.
6. **Empty rows are ignored**, so you can leave blank rows between entries.

---

## Running the tool

With your spreadsheet ready, run:

```powershell
python dtr_generator.py
```

By default the tool looks for a file named **`Sample DTR.xlsx`** in the project
folder or your **Downloads** folder. To use a specific file, pass its path:

```powershell
python dtr_generator.py "C:\path\to\your file.xlsx"
```

---

## Output files

After running, you'll find these in the project folder:

| File | What it is |
| --- | --- |
| **`DTR_output.docx`** | The finished DTR, ready to open in Word, review, and print. |
| **`preview.txt`** | A plain-text summary of everything the tool parsed — great for double-checking the data and spotting anomalies before printing. |
| `DTR_output_HHMMSS.docx` | A fallback copy, created automatically only if `DTR_output.docx` is currently open in Word (locked). |
| `docx_error.txt` | Only appears if Word generation failed — contains the error details. |

> Tip: open `preview.txt` first. It lists each week, each day's sessions, the
> hours, the assigned work, and any anomalies (missing times, Sunday entries, etc.).

To get a **PDF**, open `DTR_output.docx` in Word and use **File → Save As → PDF**
(automatic PDF export is planned for the web version — see the developer notes).

---

## How the DTR gets filled (the rules)

- **Weekday detection:** computed from each date automatically.
- **Assigned Work:** ONLINE column → *iCARE Virtual Office Duty*; ONSITE column → *iCARE Onsite Duty*.
- **Weekly grouping:** entries are grouped by the week (Monday–Saturday) they fall in.
  Only weeks that have at least one entry get a form. Weeks are laid out **two forms per page**.
- **Hours per session:** `Time Out − Time In` (the spreadsheet's DURATION columns are ignored).
- **Long sessions get split:** any session longer than `SPLIT_THRESHOLD_HOURS` (default 4 hours) is broken into same-day chunks of at most that many hours each, so no single line ever shows more hours than the threshold.
- **Multiple sessions/day:** stacked on separate lines within the same cell, with each session's hours on its own line.
- **Weekly total:** sum of all session hours that week, shown both in the totals row
  and in the top-right "Total Hours Rendered: _N_ Hours" box (value shown in 12pt bold).
- **Header fields:** Name (from cell B2), plus Term, School Year, Office, and Supervisor
  (see [Configuration](#configuration)). The Month / from / to / Year line is filled per week.

---

## Configuration

Simple settings live near the top of `dtr_generator.py`:

```python
ENTRY_FONT_PT = 9    # font size for the filled-in time/date/hours/assigned cells

SPLIT_THRESHOLD_HOURS = 4    # sessions longer than this get split into same-day chunks of at most this many hours

DEFAULTS = {
    "term": "3rd",
    "school_year": "'25 - '26",
    "office": "iCARE",
    "supervisor": "Ms. Cherelyn Joy Padasas",
}

ASSIGNED_LABELS = {
    "ONLINE": "iCARE Virtual Office Duty",
    "ONSITE": "iCARE Onsite Duty",
}
```

- Change **`DEFAULTS`** to update the Term, School Year, Office, or Supervisor
  printed on every form.
- The **Name** is always read from cell **B2** of the spreadsheet.
- Change **`ENTRY_FONT_PT`** to make the entry text bigger/smaller.
- Change **`SPLIT_THRESHOLD_HOURS`** to adjust how long a session can be before it gets split into same-day chunks.
- Change **`ASSIGNED_LABELS`** if the duty wording ever changes.

---

## Troubleshooting

| Problem | Fix |
| --- | --- |
| `'python' is not recognized` | Python isn't installed or not on PATH. Reinstall from python.org and tick "Add Python to PATH". Try `py dtr_generator.py`. |
| `PermissionError: ... DTR_output.docx` | The output file is open in Word. Close it and re-run (the tool will otherwise save a timestamped copy). |
| `ERROR: spreadsheet not found` | Make sure your file is named `Sample DTR.xlsx` (or pass the full path as an argument), and is in the project or Downloads folder. |
| `Template EMPTY_DTR.docx not found` | Keep `EMPTY_DTR.docx` in the same folder as the script. |
| Times show as text / weekday looks wrong | Ensure dates and times are entered as real Excel dates/times, not plain text. |
| A duty is missing from the form | Check `preview.txt` for anomalies (e.g., missing Time In/Out, or a Sunday entry, which the form can't hold). |
| `ModuleNotFoundError: openpyxl` / `docx` | Run `pip install -r requirements.txt` again. |

---

## Notes for developers (turning this into a website)

The engine is already structured for reuse — the web layer only needs to call a
few functions and stream back the result. No database is required; process
uploads in memory / a temp folder and discard them.

**Reusable pieces in `dtr_generator.py`:**

- `parse_workbook(xlsx_path)` → returns `(name, entries, anomalies)`.
- `group_into_weeks(entries)` → returns the list of weekly structures.
- `generate_docx(template_path, ident, weeks, out_path)` → writes the filled Word file.
  (`ident` is a dict: `name`, `term`, `school_year`, `office`, `supervisor`.)

**Suggested web flow (no database):**

1. `GET /` — a page with a file upload + editable fields (Name auto-fills from B2
   after parsing; Term / School Year pre-filled from `DEFAULTS`).
2. *(Optional)* a review screen to tweak **Assigned Work** per entry before generating.
3. `POST /generate` — parse the upload, build the `.docx` (and PDF), return download links.

**Recommended stack:** Flask (Python) so the existing engine imports directly;
a single HTML page for the UI. Refactor `parse_workbook`/`generate_docx` to also
accept file-like objects (e.g., `io.BytesIO`) so nothing needs to touch disk.

**PDF export:**

- **Windows with Microsoft Word installed:** use [`docx2pdf`](https://pypi.org/project/docx2pdf/) (drives Word via COM).
- **Linux / cloud hosting (e.g., Render, Railway, Fly.io):** install **LibreOffice**
  and convert with `soffice --headless --convert-to pdf`. On serverless platforms
  that can't run LibreOffice (e.g., plain Vercel functions), use a Docker image
  that bundles LibreOffice, or a conversion service.

**Things to keep in mind:**

- The template stores the "Total Hours Rendered" box twice (a modern + a fallback
  copy); the code already updates both. Preserve that behavior.
- The whole two-form page is cloned as the repeating unit to keep the exact
  layout — don't rebuild the blocks from scratch.
- Sundays and missing time-in/out are surfaced as *anomalies*; consider showing
  these to the user in the UI before download.
```
