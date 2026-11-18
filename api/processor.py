"""Reusable Excel processing logic extracted from the existing Streamlit app.

The Streamlit UI remains in app.py and is not imported by this API service.
"""
import re
from datetime import date, datetime
from io import BytesIO

import pandas as pd
from openpyxl.utils import get_column_letter

TECHNOLOGIES = ["SAG", "APIGEE", "ADF", "MX", "TALEND", "WMB", "KAFKA", "OAT"]

CRQ_INDEX = 0
DATE_INDEX = 1
SUMMARY_INDEX = 2
STATUS_INDEX = 3
MANAGER_INDEX = 5

OUTPUT_COLUMNS = [
    "Date", "Project", "Environment", "Component", "Implemented",
    "Category", "CRQ#", "Status", "Change Manager Name",
]

PRE_PROD_PATTERN = re.compile(r"\b(pre-?prod|prepord|stag(e|ing))\b", re.IGNORECASE)
PROD_PATTERN = re.compile(r"\bprod(uction)?\b", re.IGNORECASE)
IMPLEMENTED_STATUSES = {"closed", "completed"}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def format_date(d):
    return f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year % 100:02d}"


def build_date(y, m, d):
    try:
        return date(y, m, d)
    except ValueError:
        return None


def parse_date_string(text):
    s = text.strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$", s)
    if m:
        return build_date(int(m[1]), int(m[2]), int(m[3]))
    m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})(?:[\sT].*)?$", s)
    if m:
        a, b = int(m[1]), int(m[2])
        y = int(m[3])
        if y < 100:
            y += 2000
        month, day = (b, a) if a > 12 else (a, b)
        return build_date(y, month, day)
    m = re.match(r"^(\d{1,2})[\s-]([A-Za-z]{3,})[\s-](\d{2}|\d{4})(?:[\sT].*)?$", s)
    if m:
        name = m[2][:3].lower()
        idx = next((i for i, mn in enumerate(MONTHS) if mn.lower() == name), None)
        if idx is None:
            return None
        y = int(m[3])
        if y < 100:
            y += 2000
        return build_date(y, idx + 1, int(m[1]))
    return None


def normalize_cell(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if isinstance(value, (datetime, date)):
        return format_date(value)
    if isinstance(value, str):
        d = parse_date_string(value)
        return format_date(d) if d else value
    return value


def load_source(file_bytes, file_name):
    buffer = BytesIO(file_bytes)
    if file_name.lower().endswith(".csv"):
        df = pd.read_csv(buffer, header=None, dtype=object, keep_default_na=False)
    else:
        df = pd.read_excel(buffer, sheet_name=0, header=None, dtype=object, keep_default_na=False)
    rows = df.values.tolist()
    if len(rows) < 2:
        raise ValueError("The file has no data rows.")
    headers = [str(h).strip() for h in rows[0]]
    normalized = [[normalize_cell(c) for c in r] for r in rows[1:]]
    data = [r for r in normalized if any(c != "" for c in r)]
    if not data:
        raise ValueError("The file has no non-empty data rows.")
    return headers, data


def cell(row, i):
    return row[i] if i < len(row) else ""


def last_six_digits(value):
    if value == "" or value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    s = str(value).strip()
    return s[-6:] if len(s) > 6 else s


def get_environment(summary):
    if PRE_PROD_PATTERN.search(summary):
        return "Pre-Prod"
    if PROD_PATTERN.search(summary):
        return "Live"
    return ""


def get_implemented(status):
    return "Yes" if str(status).strip().lower() in IMPLEMENTED_STATUSES else "No"


def get_category(component):
    c = str(component).upper()
    return "Inf" if ("OAT" in c or "INFRA" in c) else "App"


def make_output_row(row, component, summary):
    status = cell(row, STATUS_INDEX)
    return {
        "Date": cell(row, DATE_INDEX),
        "Project": summary,
        "Environment": get_environment(summary),
        "Component": component,
        "Implemented": get_implemented(status),
        "Category": get_category(component),
        "CRQ#": last_six_digits(cell(row, CRQ_INDEX)),
        "Status": status,
        "Change Manager Name": cell(row, MANAGER_INDEX),
    }


def process_rows(data):
    output = []
    for row in data:
        summary = str(cell(row, SUMMARY_INDEX))
        upper = summary.upper()
        duplicable = [t for t in TECHNOLOGIES if t != "OAT" and t in upper]
        has_oat = "OAT" in upper
        if not duplicable:
            component = "OAT" if has_oat else "Infra"
            output.append(make_output_row(row, component, summary))
        else:
            for tech in duplicable:
                component = f"{tech}, OAT" if has_oat else tech
                output.append(make_output_row(row, component, summary))
    return output


def to_excel_bytes(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Processed Data", index=False)
        ws = writer.sheets["Processed Data"]
        for i, col in enumerate(df.columns, start=1):
            ws.column_dimensions[get_column_letter(i)].width = min(max(len(col), 12), 50)
    return buffer.getvalue()


def process_excel_file(file_bytes, file_name):
    _, data = load_source(file_bytes, file_name)
    output_rows = process_rows(data)
    df = pd.DataFrame(output_rows, columns=OUTPUT_COLUMNS)
    return to_excel_bytes(df), len(data), len(output_rows)
