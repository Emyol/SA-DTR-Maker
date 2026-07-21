"""Editable identity defaults and constants shared across the engine."""

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

WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday",
                 "Friday", "Saturday", "Sunday"]

TITLE_MARK = "STUDENT ASSISTANT DAILY TIME RECORD"

# Identity fields a caller may override (spreadsheet only ever supplies "name").
IDENTITY_FIELDS = ("name", "term", "school_year", "office", "supervisor")

# Defensive caps against pathological/adversarial spreadsheets (huge claimed
# dimensions, thousands of blank rows, etc.) — a real DTR log never gets close
# to these, so hitting them always means something is wrong with the input.
MAX_LABEL_SEARCH_ROWS = 200
MAX_DATA_ROWS = 5000
