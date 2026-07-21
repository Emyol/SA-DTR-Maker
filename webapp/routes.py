"""HTTP routes for the DTR web app."""

import base64
import io
import os
import re

from flask import Blueprint, current_app, render_template, request

from dtr import engine
from dtr.defaults import DEFAULTS

bp = Blueprint("dtr", __name__)

# EMPTY_DTR.docx lives at the project root, one level above webapp/.
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE_PATH = os.path.join(PROJECT_ROOT, "EMPTY_DTR.docx")

ALLOWED_EXTENSION = ".xlsx"
MAX_FIELD_LEN = 200
_ILLEGAL_FILENAME_CHARS = re.compile(r'[\\/:*?"<>|]+')


def _safe_download_name(name: str) -> str:
    base = _ILLEGAL_FILENAME_CHARS.sub("_", name or "").strip(" ._") or "output"
    return f"DTR_{base[:80]}.docx"


def _clamp(value, max_len=MAX_FIELD_LEN):
    return (value or "").strip()[:max_len]


@bp.get("/")
def index():
    return render_template("index.html", defaults=DEFAULTS)


@bp.post("/generate")
def generate():
    if not os.path.exists(TEMPLATE_PATH):
        current_app.logger.error("Template EMPTY_DTR.docx missing at %s", TEMPLATE_PATH)
        return render_template(
            "error.html",
            message="The server is missing its DTR template file. Please contact the maintainer.",
        ), 500

    input_method = _clamp(request.form.get("input_method"), 20) or "upload"
    use_paste = input_method == "paste"

    try:
        if use_paste:
            parsed_name, entries, anomalies = engine.parse_pasted_table(
                request.form.get("pasted_table", "")
            )
        else:
            file = request.files.get("spreadsheet")
            if file is None or file.filename == "":
                return render_template(
                    "error.html",
                    message="Please choose an .xlsx file to upload, or switch to Paste values.",
                ), 400

            ext = os.path.splitext(file.filename)[1].lower()
            if ext != ALLOWED_EXTENSION:
                return render_template(
                    "error.html",
                    message=f'Unsupported file type "{ext or "(none)"}" — please upload a .xlsx spreadsheet.',
                ), 400

            raw = file.read()
            if not raw:
                return render_template("error.html", message="That file is empty."), 400

            parsed_name, entries, anomalies = engine.parse_workbook(io.BytesIO(raw))
    except engine.WorkbookReadError as exc:
        return render_template("error.html", message=str(exc)), 400
    except Exception:
        current_app.logger.exception("Unexpected error parsing DTR input")
        input_label = "pasted values" if use_paste else "spreadsheet"
        return render_template(
            "error.html",
            message=(
                f"Something went wrong reading the {input_label}. Double-check it has "
                "ONLINE/ONSITE columns with recognizable dates and times, matching "
                "the sample layout."
            ),
        ), 400

    weeks = engine.group_into_weeks(entries)

    ident = dict(DEFAULTS)
    name_override = _clamp(request.form.get("name"))
    ident["name"] = name_override or parsed_name
    for field in ("term", "school_year", "office", "supervisor"):
        override = _clamp(request.form.get(field))
        if override:
            ident[field] = override

    if not weeks:
        return render_template(
            "error.html",
            message=(
                f"No usable Monday-Saturday time entries were found (parsed "
                f"{len(entries)} raw entr{'y' if len(entries) == 1 else 'ies'} total). "
                "Make sure dates/times are recognizable and the input has ONLINE/ONSITE "
                "columns like the sample file."
            ),
            anomalies=anomalies,
        ), 400

    try:
        docx_bytes = engine.generate_docx(TEMPLATE_PATH, ident, weeks)
    except Exception:
        current_app.logger.exception("Unexpected error generating DTR docx")
        return render_template(
            "error.html",
            message="Something went wrong generating the Word document. Please try again.",
        ), 500

    filename = _safe_download_name(ident["name"])
    total_hours = sum(w["total"] for w in weeks)
    # Embed the file as a data: URI rather than a server-side download token:
    # this runs as a stateless serverless function (Vercel), so nothing
    # written to memory or disk in this request is guaranteed to still be
    # around by the time a second request (the download) arrives — it might
    # land on a different instance entirely. Putting the bytes directly in
    # the response the browser already has sidesteps that.
    docx_b64 = base64.b64encode(docx_bytes).decode("ascii")

    return render_template(
        "result.html",
        docx_b64=docx_b64,
        filename=filename,
        ident=ident,
        weeks=weeks,
        entries_count=len(entries),
        total_hours=total_hours,
        anomalies=anomalies,
        fmt_hours=engine.fmt_hours,
        fmt_date=engine.fmt_date,
    )
