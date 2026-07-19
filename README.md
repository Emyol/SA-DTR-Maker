# DTR Helper

Automatically fills the **Student Assistant Daily Time Record (DTR)** Word form
from a simple spreadsheet of time logs.

You paste your time-in / time-out entries into a spreadsheet, upload it (web
app) or run one command (CLI), and the tool produces a ready-to-print Word
document with one weekly (Mon–Sat) DTR form per week, two forms per page —
with the daily hours, weekly totals, and "Assigned Work" all filled in for you.

---

## Table of contents

- [What it does](#what-it-does)
- [Repository contents](#repository-contents)
- [Requirements](#requirements)
- [Installation](#installation)
- [Preparing your spreadsheet](#preparing-your-spreadsheet) ← **most important**
- [Using the web app](#using-the-web-app)
- [Using the command line](#using-the-command-line)
- [Output files](#output-files)
- [How the DTR gets filled (the rules)](#how-the-dtr-gets-filled-the-rules)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [Architecture / for developers](#architecture--for-developers)

---

## What it does

- Reads a spreadsheet with two time logs: **ONLINE** (virtual duty) and **ONSITE** (onsite duty).
- Figures out the **exact weekday** of every date automatically.
- Groups all entries into **weekly forms (Monday–Saturday)** — one form per week that has records.
- Computes **hours rendered per session**, and the **weekly total**.
- Handles **multiple sessions on the same day** (they stack inside the same cell).
- Fills the **"Total Hours Rendered"** box at the top-right of each weekly form.
- Outputs a Word file that matches the official template exactly (two forms per page).
- Ships two ways to use it: a **web app** (upload a file in the browser, download
  the finished .docx — nothing is written to disk) and a **command-line script**.
- The spreadsheet layout is **auto-detected** — it doesn't have to match the
  sample's exact cell positions, only the ONLINE/ONSITE labels need to be present.

---

## Repository contents

| Path | Purpose |
| --- | --- |
| `dtr/` | The reusable engine: parsing, week-grouping, and Word generation. Used by both the web app and the CLI. |
| `webapp/` | The Flask web app (routes, in-memory download store, HTML templates, CSS). |
| `run_web.py` | Starts the web app (`python run_web.py`). |
| `dtr_generator.py` | The command-line entry point. |
| `EMPTY_DTR.docx` | The official blank DTR template that gets filled. **Required.** |
| `requirements.txt` | The Python libraries needed. |
| `SAMPLE DTR.xlsx` | An example time-logs spreadsheet you can copy and edit. |
| `inspect_files.py` | Optional helper for inspecting the template/spreadsheet structure. |
| `README.md` | This file. |

> ⚠️ **`EMPTY_DTR.docx` must stay in the project's root folder.** It is the
> template both the web app and the CLI fill in.

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
matches the included `SAMPLE DTR.xlsx` — the easiest approach is to **copy that
file and just replace the data**.

### The layout

There are **two side-by-side logs**. The left log is **ONLINE (virtual)** and the
right log is **ONSITE**. Each log has a Date, Time In, and Time Out column.

| Cell / Column | What goes there |
| --- | --- |
| A cell containing your full name | e.g. `ACUÑA, AMIEL JOSIAH C.` — one row above the `ONLINE` label, same column |
| A row with the section labels | `ONLINE` and, to its right, `ONSITE` |
| The next row | Column headers (`DATE (mm/dd/yy)`, `TIME IN`, `TIME OUT`, `DURATION`) |
| Every row below that | Your actual time entries, 3 columns per log: Date, Time In, Time Out |

> **The exact cell positions don't matter.** The tool searches the sheet for
> the `ONLINE`/`ONSITE` labels and works out the columns and starting row from
> there, so the sample's B2/row‑5 layout and a shifted layout (e.g. name in
> C3, data starting row 6) both work. Only the *relative* shape has to hold:
> name directly above the label, headers directly below it, data below that,
> and Date/Time In/Time Out as the next three columns after the label.

**Each log, from the label's column rightward:**

| Column offset | Meaning |
| --- | --- |
| **+0** (the label's own column) | Date of the duty (a real date, e.g. `04/21/26`) |
| **+1** | Time In (a real time, e.g. `9:00 AM`) |
| **+2** | Time Out (a real time, e.g. `1:00 PM`) |
| **+3** | Duration — *optional, ignored* (the tool computes hours itself) |

### Rules for entering data

1. **Put virtual duties under the ONLINE columns** and **onsite duties under the
   ONSITE columns.** The tool decides "Assigned Work" based on **which side
   you put the entry on**, not on the weekday:
   - Anything in the ONLINE columns → **"iCARE Virtual Office Duty"**
   - Anything in the ONSITE columns → **"iCARE Onsite Duty"**
2. **Dates and times should be real Excel dates/times** where possible. Common
   text formats (`04/21/2026`, `9:00 AM`, `21:00`) are also tolerated, but real
   Excel date/time cells are the most reliable — Excel auto-converts `9:00 AM`
   as you type it.
3. **The two sides are independent.** A row does not need matching ONLINE and
   ONSITE entries — leave a side blank if there was no duty there that row.
4. **Multiple sessions on one day are fine.** If you worked twice on the same
   date (e.g., 11:00 AM–3:00 PM and 4:00 PM–7:00 PM), just add two rows with the
   same date. They will stack in the same day's cell on the form.
5. **Sundays are not supported** — the DTR form only has Monday–Saturday rows.
   Any Sunday entry is flagged as an anomaly and left off the form.
6. **Empty rows are ignored**, so you can leave blank rows between entries.

---

## Using the web app

The web app runs entirely on your machine — your spreadsheet is processed in
memory and never written to disk (no upload folder, nothing left behind).

```powershell
pip install -r requirements.txt
python run_web.py
```

Then open **http://127.0.0.1:5000** in a browser:

1. Upload your `.xlsx` spreadsheet (max 5 MB).
2. Optionally override Name / Term / School Year / Office / Supervisor — leave
   blank to use the name from the spreadsheet and the defaults in
   [Configuration](#configuration).
3. Click **Generate DTR**. You'll land on a results page showing the weeks
   found, total hours, and any anomalies (Sunday entries, missing times, etc.)
   *before* you download — review those, then click **Download**.

The download link stays valid for 20 minutes; if it expires, just re-upload.

To stop the server, press `Ctrl+C` in its terminal.

---

## Using the command line

If you'd rather not run a server, the original CLI still works:

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

**Web app:** nothing is written to your project folder. The results page shows
the same information `preview.txt` used to (weeks, sessions, hours, anomalies),
and the finished file downloads as `DTR_<name>.docx` through your browser.

**Command line:** after running `dtr_generator.py`, you'll find these in the
project folder:

| File | What it is |
| --- | --- |
| **`DTR_output.docx`** | The finished DTR, ready to open in Word, review, and print. |
| **`preview.txt`** | A plain-text summary of everything the tool parsed — great for double-checking the data and spotting anomalies before printing. |
| `DTR_output_HHMMSS.docx` | A fallback copy, created automatically only if `DTR_output.docx` is currently open in Word (locked). |
| `docx_error.txt` | Only appears if Word generation failed — contains the error details. |

> Tip: open `preview.txt` first. It lists each week, each day's sessions, the
> hours, the assigned work, and any anomalies (missing times, Sunday entries, etc.).

To get a **PDF**, open the `.docx` in Word and use **File → Save As → PDF**.

---

## How the DTR gets filled (the rules)

- **Weekday detection:** computed from each date automatically.
- **Assigned Work:** ONLINE column → *iCARE Virtual Office Duty*; ONSITE column → *iCARE Onsite Duty*.
- **Weekly grouping:** entries are grouped by the week (Monday–Saturday) they fall in.
  Only weeks that have at least one entry get a form. Weeks are laid out **two forms per page**.
- **Hours per session:** `Time Out − Time In` (the spreadsheet's DURATION columns are ignored).
- **Multiple sessions/day:** stacked on separate lines within the same cell, with each session's hours on its own line.
- **Weekly total:** sum of all session hours that week, shown both in the totals row
  and in the top-right "Total Hours Rendered: _N_ Hours" box (value shown in 12pt bold).
- **Header fields:** Name (auto-detected from the spreadsheet, or overridden in the
  web form), plus Term, School Year, Office, and Supervisor
  (see [Configuration](#configuration)). The Month / from / to / Year line is filled per week.

---

## Configuration

Simple settings live in `dtr/defaults.py`:

```python
ENTRY_FONT_PT = 9    # font size for the filled-in time/date/hours/assigned cells

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
  printed on every form when the user doesn't override them.
- The **Name** is auto-detected from the spreadsheet (see
  [Preparing your spreadsheet](#preparing-your-spreadsheet)); the web form and
  the CLI both let it be overridden.
- Change **`ENTRY_FONT_PT`** to make the entry text bigger/smaller.
- Change **`ASSIGNED_LABELS`** if the duty wording ever changes.
- The web app's upload size cap (default 5 MB) is `MAX_UPLOAD_MB` in `webapp/__init__.py`.

---

## Troubleshooting

| Problem | Fix |
| --- | --- |
| `'python' is not recognized` | Python isn't installed or not on PATH. Reinstall from python.org and tick "Add Python to PATH". Try `py run_web.py` / `py dtr_generator.py`. |
| Browser shows "Unsupported file type" | Only `.xlsx` is accepted — older `.xls`, `.csv`, or Google Sheets exports need to be saved as `.xlsx` first. |
| "That file isn't a valid .xlsx workbook…" | The file is corrupted, actually a different format, or password-protected. Re-save it from Excel as a plain `.xlsx` with no password. |
| "No usable Monday–Saturday time entries were found" | The ONLINE/ONSITE labels are missing, or all entries were flagged as anomalies (check the list shown on the error page) — e.g. dates/times aren't recognizable, or every entry fell on a Sunday. |
| Download link says expired / 404 | Generated files are kept in memory for 20 minutes only. Re-upload the spreadsheet to generate a fresh one. |
| "That file is too large" | The web app caps uploads at 5 MB (configurable, see [Configuration](#configuration)); a normal DTR log is a few KB. |
| `PermissionError: ... DTR_output.docx` (CLI only) | The output file is open in Word. Close it and re-run (the tool will otherwise save a timestamped copy). |
| `ERROR: spreadsheet not found` (CLI only) | Make sure your file is named `Sample DTR.xlsx` (or pass the full path as an argument), and is in the project or Downloads folder. |
| `Template EMPTY_DTR.docx not found` | Keep `EMPTY_DTR.docx` in the project's root folder. |
| Times show as text / weekday looks wrong | Ensure dates and times are entered as real Excel dates/times where possible; common text formats are tolerated but real date/time cells are most reliable. |
| A duty is missing from the form | Check the anomalies list (results page in the web app, `preview.txt` on the CLI) — e.g. missing Time In/Out, or a Sunday entry, which the form can't hold. |
| `ModuleNotFoundError: openpyxl` / `docx` / `flask` | Run `pip install -r requirements.txt` again (make sure you're installing into the same Python/venv you're running with). |

---

## Architecture / for developers

```
dtr/            reusable engine — no I/O beyond what callers pass in
  engine.py       parse_workbook(), group_into_weeks(), generate_docx(), ...
  defaults.py     DEFAULTS, ASSIGNED_LABELS, and other constants
webapp/          Flask app
  __init__.py     app factory, upload-size limit, error handlers
  routes.py       "/", "/generate", "/download/<token>"
  storage.py      short-lived in-memory store for generated .docx bytes
  templates/      upload form, results page, error page
  static/         style.css
run_web.py       dev server entry point
dtr_generator.py CLI entry point (thin wrapper around dtr.engine)
```

**Key design points:**

- `dtr/engine.py` is the single source of truth for parsing/generation logic;
  both `dtr_generator.py` and `webapp/routes.py` call into it. `parse_workbook()`
  and `generate_docx()` accept either a filesystem path or a file-like object
  (`io.BytesIO`), so the web app never writes the uploaded spreadsheet or the
  generated `.docx` to disk.
- **Layout auto-detection:** `detect_layout()` locates the `ONLINE`/`ONSITE`
  labels instead of assuming fixed cell positions, so spreadsheets shifted by a
  row/column (as long as the labels and column order are intact) still parse.
  It falls back to the original fixed B2/row-5 layout if no labels are found.
  Date/time parsing also tolerates common text formats, not just native Excel
  date/time cells.
- **Errors are typed:** `WorkbookReadError` (unreadable file) and
  `NoEntriesError` (nothing usable to fill in) are raised as `dtr.engine.DTRError`
  subclasses with user-safe messages, so the web layer can show them directly
  without leaking stack traces; anything else is logged server-side and shown
  as a generic message.
- **Download flow:** `POST /generate` renders a results page (weeks, totals,
  anomalies) and stores the generated bytes in `webapp/storage.py` behind a
  random token; `GET /download/<token>` streams them back. This lets the user
  review anomalies before downloading, without ever touching a temp file on
  disk (sidesteps the file-locking issues a saved-to-disk approach can hit).
  Entries expire after 20 minutes.
- **Defensive caps:** uploads are capped at 5 MB (`MAX_UPLOAD_MB`), label
  search is bounded to the first 200 rows, and data parsing is bounded to 5000
  rows — a real DTR log never gets close, so hitting these limits always
  signals something's wrong with the input rather than truncating real data
  silently (a note is added to the anomalies list if truncation happens).

**PDF export (not yet implemented):**

- **Windows with Microsoft Word installed:** use [`docx2pdf`](https://pypi.org/project/docx2pdf/) (drives Word via COM).
- **Linux / cloud hosting (e.g., Render, Railway, Fly.io):** install **LibreOffice**
  and convert with `soffice --headless --convert-to pdf`. On serverless platforms
  that can't run LibreOffice (e.g., plain Vercel functions), use a Docker image
  that bundles LibreOffice, or a conversion service.

**Deploying:** `run_web.py` uses Flask's dev server, which is fine for local/LAN
use. For a real deployment, run the `webapp.create_app()` factory under a
production WSGI server (e.g. `waitress-serve --call webapp:create_app`) behind
a reverse proxy, and set `DTR_SECRET_KEY` to a real secret via environment
variable instead of the `dev-only-change-me` default.
