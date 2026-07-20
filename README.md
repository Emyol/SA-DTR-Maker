# DTR Helper

Automatically fills the **Student Assistant Daily Time Record (DTR)** Word form
from a simple spreadsheet of time logs.

You paste your time-in / time-out entries into a spreadsheet, upload it (web
app) or run one command (CLI), and the tool produces a ready-to-print Word
document with one weekly (Mon–Sat) DTR form per week, two forms per page —
with the daily hours, weekly totals, and "Assigned Work" all filled in for you.

---

## Table of contents

- [Design brief (for UI redesign)](#design-brief-for-ui-redesign)
- [What it does](#what-it-does)
- [Repository contents](#repository-contents)
- [Requirements](#requirements)
- [Installation](#installation)
- [Preparing your spreadsheet](#preparing-your-spreadsheet) ← **most important**
- [Using the web app](#using-the-web-app)
- [Deploying to Vercel](#deploying-to-vercel)
- [Using the command line](#using-the-command-line)
- [Output files](#output-files)
- [How the DTR gets filled (the rules)](#how-the-dtr-gets-filled-the-rules)
- [Configuration](#configuration)
- [Troubleshooting](#troubleshooting)
- [Architecture / for developers](#architecture--for-developers)

---

## Design brief (for UI redesign)

> This section is written to stand on its own — hand it to a fresh design
> session with no other context and it should be enough to redesign the UI.

**What this app is:** a small internal tool for student assistants (at what
the current copy calls the "iCARE" office) to turn a spreadsheet of daily
time-in/time-out logs into a filled, printable Word DTR (Daily Time Record)
form. One upload in, one Word document out. No accounts, no dashboard, no
history — a single-purpose utility used occasionally (weekly/monthly), not a
daily-driver app.

**Stack the UI is built in (constraints for a redesign):**
- Server-rendered **Flask + Jinja2** templates — not a JS framework/SPA.
  `webapp/templates/base.html` is the shared shell; `index.html`, `result.html`,
  `error.html` extend it via `{% block content %}`.
- Styling is one plain CSS file, `webapp/static/style.css`, using a small
  hand-rolled design-token system (CSS custom properties) — no Tailwind, no
  component library.
- A little vanilla JS in `index.html` (no build step, no framework): drag-and-drop
  styling for the file input, and disabling the submit button on submit.
- **Any redesign needs to stay implementable as HTML + CSS (+ optionally
  small vanilla JS) inside these same three Jinja2 templates** — the `{{ ... }}`
  variable interpolations and `{% for/if %}` loops in the templates below are
  load-bearing (real data), not placeholder content, and must be preserved
  even if every visual element around them changes.
- Light **and** dark mode both need to work — currently driven automatically by
  `prefers-color-scheme` (no manual toggle in the UI).
- Must be responsive — real users will open this on phones to check a form
  before printing it at a computer lab/office.

**Current visual system** (free to keep, evolve, or fully replace):
spacing scale 4/8/12/16/24/32/48/64px · type scale 12/13/14/16/20/24/32px ·
radius 6px (controls) / 10px (cards) / 999px (pills) · one accent hue (blue,
`hsl(217 88% 55%)`) used only for the primary button and focus rings · mostly
neutral grays otherwise · content constrained to `max-width: 720px`, centered.

### The three screens

**1. Upload page (`/`, `index.html`)** — the entry point, a single form:
- Heading: "DTR Helper" / subtext: "Fill the Student Assistant Daily Time
  Record from a time-log spreadsheet."
- Card 1, "1. Upload your time-log spreadsheet": a drag-and-drop file zone
  (click or drag a `.xlsx`), hint text "Click to choose a file, or drag it
  here" / ".xlsx only, up to 4 MB", plus an explainer paragraph about the
  ONLINE/ONSITE column format. Once a file is chosen, the zone should show
  "Selected: <filename>" instead of the empty-state prompt.
- Card 2, "2. Confirm the details on the form (optional)": a "Name override"
  text field (placeholder "Leave blank to use the name from the spreadsheet"),
  then a 2-column grid of Term / School Year / Office / Supervisor text
  fields, pre-filled with defaults.
- One full-width primary button: "Generate DTR" (becomes "Generating…" and
  disables itself on submit — this is a real network wait, generation of a
  multi-week Word doc isn't instant).

**2. Results page (`result.html`)** — shown after a successful upload:
- A green success banner: "Your DTR is ready. Parsed *N* time entries into
  *N* weeks for *Name*."
- *Conditionally*, an amber warning banner listing anomalies (things skipped
  or flagged — e.g. a Sunday entry the form has no row for, or a missing
  time-in/out) — this list can be empty (0 items) or fairly long (10+ items)
  and needs to degrade gracefully either way.
- A card with 3 stat tiles (Weeks filled / Time entries / Total hours), a
  primary "Download <filename>" button, a secondary "Start over" button, and
  a hint that the file only exists on this page (refreshing loses it — see
  *why this matters* below).
- A card listing every week included: date range, weekly total hours, and a
  one-line summary per day worked (e.g. "Wed 04/22/26 (2 sessions)"), with a
  small warning badge if that week had Sunday entries excluded from the form.
  This list can be 1 week or 20+ weeks long — needs a sane long-list treatment.

**3. Error page (`error.html`)** — shown when generation fails:
- A red banner with a short headline ("Couldn't generate the DTR.") and a
  specific, human-readable reason (e.g. "That file isn't a valid .xlsx
  workbook…", "No usable Monday–Saturday time entries were found…").
- *Conditionally*, a card listing raw diagnostic details (anomalies) for
  troubleshooting.
- A single "Try again" button back to the upload page.

### Why this matters (don't design these away)

- **The download button is not a link to a file on a server** — it's an
  `<a download>` with the entire `.docx` embedded as a base64 `data:` URI.
  This is a deliberate architecture choice (the app runs as a stateless
  serverless function on Vercel), so "download" is instant/local with no
  network request, but the file really does stop existing the moment you
  navigate away or refresh. Any redesign should keep this expectation clear
  to the user rather than implying a durable, revisitable download link.
- **Anomalies are the app's most important trust signal.** This tool silently
  drops data it can't confidently parse (a Sunday entry, a missing time-out);
  surfacing that list *before* the user relies on the generated document is
  the whole point — don't bury or minimize it.
- The two-card upload form is really one step conceptually (upload +
  optionally tweak fields, then one submit) — the "1. / 2." numbering is a
  weak stand-in for a clearer sense of progress/order, open to improvement.

### Known rough edges (fair game to fix)

- The upload page's two cards and the plain SVG upload-cloud icon read as
  generic/utilitarian — there's no real branding or personality to the
  file-drop icon or the page as a whole.
- The stat row / week-list on the results page is plain text-in-boxes; for
  20+ weeks it just becomes a long, undifferentiated scroll.
- No visual distinction between the "optional" identity fields and file
  upload in terms of visual weight — currently equal-sized cards even though
  step 1 is required and step 2 is all-optional overrides.
- No loading/progress feedback beyond the button label change — generating
  many weeks can take a couple of seconds.

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
| `webapp/` | The Flask web app (routes, HTML templates, CSS). |
| `run_web.py` | Starts the web app locally (`python run_web.py`). |
| `api/index.py` | Vercel serverless entrypoint — exposes the same Flask app for deployment. |
| `vercel.json`, `.vercelignore` | Vercel deployment config (see [Deploying to Vercel](#deploying-to-vercel)). |
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

1. Upload your `.xlsx` spreadsheet (max 4 MB).
2. Optionally override Name / Term / School Year / Office / Supervisor — leave
   blank to use the name from the spreadsheet and the defaults in
   [Configuration](#configuration).
3. Click **Generate DTR**. You'll land on a results page showing the weeks
   found, total hours, and any anomalies (Sunday entries, missing times, etc.)
   *before* you download — review those, then click **Download**.

The finished file is embedded directly in that results page (not stored on
the server), so there's nothing to expire — but it also means leaving or
refreshing the page means re-uploading to generate it again.

To stop the server, press `Ctrl+C` in its terminal.

---

## Deploying to Vercel

The web app deploys as-is to Vercel's Python serverless runtime — no code
changes needed beyond what's already in the repo (`api/index.py`,
`vercel.json`, `.vercelignore`).

```powershell
npm install -g vercel     # if you don't have the CLI yet
vercel login
vercel link                 # first time only, creates/links the project
vercel env add DTR_SECRET_KEY production   # paste a random secret when prompted
vercel env add DTR_SECRET_KEY preview      # same, for preview deployments
vercel deploy --prod
```

**Why the app is safe to run serverless:** each request may land on a
different function instance, so the app never relies on anything written to
memory or disk surviving between requests — the generated `.docx` is
returned as a `data:` URI embedded directly in the results page's download
button rather than via a separate download endpoint. See
[Architecture](#architecture--for-developers) for details.

**After deploying:**
- New Vercel projects have **Deployment Protection** (an SSO login wall) on
  by default. If you want the app usable by anyone with the link, go to the
  project on vercel.com → **Settings → Deployment Protection** and set it to
  **Disabled** (or "Only Preview Deployments" to keep Production public while
  protecting previews). This isn't something the CLI can toggle.
- The 4 MB upload cap is intentional — it's under Vercel's serverless
  request-body limit (4.5 MB), so you get the app's friendly error message
  instead of a raw platform 413.
- `DTR_SECRET_KEY` should be a real random value in production (e.g.
  `python -c "import secrets; print(secrets.token_hex(32))"`), not the
  `dev-only-change-me` fallback used locally.

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
- **Long sessions get split:** any session longer than `SPLIT_THRESHOLD_HOURS` (default 4 hours) is broken into same-day chunks of at most that many hours each, so no single line ever shows more hours than the threshold.
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
  printed on every form when the user doesn't override them.
- The **Name** is auto-detected from the spreadsheet (see
  [Preparing your spreadsheet](#preparing-your-spreadsheet)); the web form and
  the CLI both let it be overridden.
- Change **`ENTRY_FONT_PT`** to make the entry text bigger/smaller.
- Change **`SPLIT_THRESHOLD_HOURS`** to adjust how long a session can be before it gets split into same-day chunks.
- Change **`ASSIGNED_LABELS`** if the duty wording ever changes.
- The web app's upload size cap (default 4 MB) is `MAX_UPLOAD_MB` in `webapp/__init__.py`.

---

## Troubleshooting

| Problem | Fix |
| --- | --- |
| `'python' is not recognized` | Python isn't installed or not on PATH. Reinstall from python.org and tick "Add Python to PATH". Try `py run_web.py` / `py dtr_generator.py`. |
| Browser shows "Unsupported file type" | Only `.xlsx` is accepted — older `.xls`, `.csv`, or Google Sheets exports need to be saved as `.xlsx` first. |
| "That file isn't a valid .xlsx workbook…" | The file is corrupted, actually a different format, or password-protected. Re-save it from Excel as a plain `.xlsx` with no password. |
| "No usable Monday–Saturday time entries were found" | The ONLINE/ONSITE labels are missing, or all entries were flagged as anomalies (check the list shown on the error page) — e.g. dates/times aren't recognizable, or every entry fell on a Sunday. |
| Download button doesn't work / page feels huge | The file is embedded in the results page itself — if you left or refreshed that page, just re-upload to generate it again. |
| "That file is too large" | The web app caps uploads at 4 MB (configurable, see [Configuration](#configuration)) — kept under Vercel's platform limit; a normal DTR log is a few KB. |
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
  routes.py       "/", "/generate"
  templates/      upload form, results page, error page
  static/         style.css
run_web.py       local dev server entry point
api/index.py     Vercel serverless entrypoint (same Flask app)
vercel.json      Vercel routing config
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
  anomalies) with the generated `.docx` embedded directly as a base64 `data:`
  URI on the download button — no server-side download endpoint, no file
  written to disk, nothing kept in memory between requests. This is
  deliberate: as a serverless function, a given request can land on any
  instance, so anything that needs to survive from `POST /generate` to a
  later `GET /download/...` (e.g. an in-memory token store) is unreliable in
  that environment. Embedding the result in the same response sidesteps the
  problem entirely, and happens to also dodge the kind of file-locking issue
  a saved-to-disk approach hit earlier in this project's history.
- **Defensive caps:** uploads are capped at 4 MB (`MAX_UPLOAD_MB`, under
  Vercel's serverless request-body limit), label search is bounded to the
  first 200 rows, and data parsing is bounded to 5000 rows — a real DTR log
  never gets close, so hitting these limits always signals something's wrong
  with the input rather than truncating real data silently (a note is added
  to the anomalies list if truncation happens).

**PDF export (not yet implemented):**

- **Windows with Microsoft Word installed:** use [`docx2pdf`](https://pypi.org/project/docx2pdf/) (drives Word via COM).
- **Linux / cloud hosting (e.g., Render, Railway, Fly.io):** install **LibreOffice**
  and convert with `soffice --headless --convert-to pdf`. On serverless platforms
  that can't run LibreOffice (e.g., plain Vercel functions), use a Docker image
  that bundles LibreOffice, or a conversion service.

**Deploying:** `run_web.py` uses Flask's dev server, fine for local/LAN use.
For Vercel, see [Deploying to Vercel](#deploying-to-vercel) — no extra WSGI
server needed, `api/index.py` exposes the same `webapp.create_app()` app
directly to Vercel's Python runtime. For a traditional host instead, run the
`webapp.create_app()` factory under a production WSGI server (e.g.
`waitress-serve --call webapp:create_app`) behind a reverse proxy. Either way,
set `DTR_SECRET_KEY` to a real secret via environment variable instead of the
`dev-only-change-me` default.
