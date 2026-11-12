import re
from datetime import date, datetime
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl.utils import get_column_letter

# ============================================================
# CONFIGURATION
# ============================================================
LOGIN_USERNAME = "BOOTS"
LOGIN_PASSWORD = "123213"

TECHNOLOGIES = ["SAG", "APIGEE", "ADF", "MX", "TALEND", "WMB", "KAFKA", "OAT"]

# Source column positions (A = 0, B = 1, C = 2, D = 3, E = 4, F = 5, G = 6)
CRQ_INDEX = 0        # A - Display Id
DATE_INDEX = 1       # B - Scheduled start date
SUMMARY_INDEX = 2    # C - Summary
STATUS_INDEX = 3     # D - Status
# E (Submitter) is intentionally not used
MANAGER_INDEX = 5    # F - Change Manager Name

OUTPUT_COLUMNS = [
    "Date", "Project", "Environment", "Component", "Implemented",
    "Category", "CRQ#", "Status", "Change Manager Name",
]

PRE_PROD_PATTERN = re.compile(r"\b(pre-?prod|prepord|stag(e|ing))\b", re.IGNORECASE)
PROD_PATTERN = re.compile(r"\bprod(uction)?\b", re.IGNORECASE)
IMPLEMENTED_STATUSES = {"closed", "completed"}
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


# ============================================================
# DATE HANDLING  (output format: dd-Mon-yy, e.g. 30-Oct-26)
# ============================================================
def format_date(d):
    return f"{d.day:02d}-{MONTHS[d.month - 1]}-{d.year % 100:02d}"


def build_date(y, m, d):
    try:
        return date(y, m, d)
    except ValueError:
        return None


def parse_date_string(text):
    s = text.strip()

    # ISO style: 2026-10-30 or 2026-10-30 09:30 PM
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s].*)?$", s)
    if m:
        return build_date(int(m[1]), int(m[2]), int(m[3]))

    # Numeric style: 10/13/2026 9:30 PM (month-first), 30/10/2026 (day-first), 30.10.26
    m = re.match(r"^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})(?:[\sT].*)?$", s)
    if m:
        a, b = int(m[1]), int(m[2])
        y = int(m[3])
        if y < 100:
            y += 2000
        month, day = (b, a) if a > 12 else (a, b)   # a > 12 can only be a day
        return build_date(y, month, day)

    # Text month style: 30-Oct-26, 30 Oct 2026, 30-October-2026
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


# ============================================================
# FILE LOADING
# ============================================================
@st.cache_data(show_spinner=False)
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
    return headers, data


# ============================================================
# PROCESSING
# ============================================================
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
        "Date": cell(row, DATE_INDEX),                       # A
        "Project": summary,                                  # B
        "Environment": get_environment(summary),             # C
        "Component": component,                              # D
        "Implemented": get_implemented(status),              # E
        "Category": get_category(component),                 # F
        "CRQ#": last_six_digits(cell(row, CRQ_INDEX)),       # G
        "Status": status,                                    # H
        "Change Manager Name": cell(row, MANAGER_INDEX),     # I
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


# ============================================================
# EXPORT
# ============================================================
def to_excel_bytes(df):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Processed Data", index=False)
        ws = writer.sheets["Processed Data"]
        for i, col in enumerate(df.columns, start=1):
            ws.column_dimensions[get_column_letter(i)].width = min(max(len(col), 12), 50)
    return buffer.getvalue()


# ============================================================
# STYLING & ANIMATIONS
# ============================================================
CUSTOM_CSS = """
<style>
/* ---------- Keyframes ---------- */
@keyframes gradientShift {
    0%   { background-position: 0% 50%; }
    50%  { background-position: 100% 50%; }
    100% { background-position: 0% 50%; }
}
@keyframes fadeUp {
    from { opacity: 0; transform: translateY(18px); }
    to   { opacity: 1; transform: translateY(0); }
}
@keyframes fadeIn {
    from { opacity: 0; }
    to   { opacity: 1; }
}
@keyframes scaleIn {
    from { opacity: 0; transform: scale(0.94); }
    to   { opacity: 1; transform: scale(1); }
}
@keyframes slideIn {
    from { opacity: 0; transform: translateX(-14px); }
    to   { opacity: 1; transform: translateX(0); }
}
@keyframes shimmer {
    0%   { background-position: -200% center; }
    100% { background-position: 200% center; }
}
@keyframes float {
    0%, 100% { transform: translateY(0); }
    50%      { transform: translateY(-6px); }
}
@keyframes pulseGlow {
    0%, 100% { box-shadow: 0 0 0 0 rgba(37, 99, 235, 0.35); }
    50%      { box-shadow: 0 0 0 12px rgba(37, 99, 235, 0); }
}

/* ---------- Background & page ---------- */
.stApp {
    background: linear-gradient(-45deg, #eef2ff, #f0f9ff, #faf5ff, #ecfeff);
    background-size: 400% 400%;
    animation: gradientShift 18s ease infinite;
}
.block-container {
    padding-top: 2rem;
    max-width: 1200px;
    animation: fadeIn 0.6s ease;
}
footer { visibility: hidden; }

/* ---------- Hero header ---------- */
.hero {
    text-align: center;
    padding: 8px 0 18px 0;
    animation: fadeUp 0.7s ease both;
}
.hero-logo {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 64px;
    height: 64px;
    border-radius: 18px;
    background: linear-gradient(135deg, #1e40af, #0ea5e9);
    color: #fff;
    font-size: 30px;
    box-shadow: 0 10px 28px rgba(30, 64, 175, 0.35);
    animation: float 4s ease-in-out infinite;
    margin-bottom: 12px;
}
.hero-title {
    font-size: 2.1rem;
    font-weight: 800;
    letter-spacing: -0.5px;
    background: linear-gradient(90deg, #0f172a 0%, #1e40af 35%, #0ea5e9 50%, #1e40af 65%, #0f172a 100%);
    background-size: 200% auto;
    -webkit-background-clip: text;
    background-clip: text;
    -webkit-text-fill-color: transparent;
    animation: shimmer 6s linear infinite;
}
.hero-sub {
    color: #64748b;
    font-size: 0.95rem;
    margin-top: 4px;
}

/* ---------- Section titles ---------- */
.section-title {
    display: flex;
    align-items: center;
    gap: 10px;
    font-size: 1.05rem;
    font-weight: 700;
    color: #0f172a;
    margin: 18px 0 10px 0;
    animation: slideIn 0.5s ease both;
}
.step-badge {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 26px;
    height: 26px;
    border-radius: 50%;
    background: #dbeafe;
    color: #1e40af;
    font-size: 0.8rem;
    font-weight: 700;
    transition: transform 0.3s ease, background 0.3s ease;
}
.section-title:hover .step-badge {
    transform: scale(1.15) rotate(-8deg);
    background: #bfdbfe;
}

/* ---------- Login card ---------- */
div[data-testid="stForm"] {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 18px;
    padding: 28px 28px 22px 28px;
    box-shadow: 0 20px 50px rgba(15, 23, 42, 0.12);
    animation: scaleIn 0.55s cubic-bezier(0.2, 0.8, 0.2, 1) both;
    transition: box-shadow 0.4s ease, transform 0.4s ease;
}
div[data-testid="stForm"]:hover {
    box-shadow: 0 26px 60px rgba(30, 64, 175, 0.18);
    transform: translateY(-2px);
}
.login-caption {
    text-align: center;
    color: #64748b;
    font-size: 0.9rem;
    margin-bottom: 14px;
    animation: fadeIn 0.9s ease both;
}
div[data-testid="stTextInput"] input {
    border-radius: 10px !important;
    transition: border-color 0.25s ease, box-shadow 0.25s ease;
}
div[data-testid="stTextInput"] input:focus {
    border-color: #2563eb !important;
    box-shadow: 0 0 0 4px rgba(37, 99, 235, 0.15) !important;
}

/* ---------- Buttons ---------- */
div.stButton > button,
div.stDownloadButton > button {
    border-radius: 10px;
    font-weight: 600;
    padding: 0.55rem 1.1rem;
    transition: transform 0.2s ease, box-shadow 0.25s ease, background 0.25s ease;
}
div.stButton > button:hover,
div.stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 10px 22px rgba(15, 23, 42, 0.14);
}
div.stButton > button:active,
div.stDownloadButton > button:active {
    transform: translateY(0) scale(0.98);
}
div.stButton > button[kind="primary"],
div.stFormSubmitButton > button[kind="primary"] {
    background: linear-gradient(135deg, #1e40af, #1e3a8a);
    border: none;
    color: #fff;
}
div.stButton > button[kind="primary"]:hover,
div.stFormSubmitButton > button[kind="primary"]:hover {
    box-shadow: 0 10px 24px rgba(30, 64, 175, 0.4);
    animation: pulseGlow 1.6s ease infinite;
}
div.stDownloadButton > button {
    background: linear-gradient(135deg, #059669, #047857);
    border: none;
    color: #fff;
}
div.stDownloadButton > button:hover {
    box-shadow: 0 10px 24px rgba(5, 150, 105, 0.35);
    color: #fff;
}
div.stDownloadButton > button p { color: #fff !important; }

/* ---------- File uploader ---------- */
div[data-testid="stFileUploader"] section {
    border: 2px dashed #93c5fd !important;
    border-radius: 16px !important;
    background: linear-gradient(180deg, #f8fbff, #eff6ff) !important;
    transition: all 0.3s ease;
    animation: fadeUp 0.6s ease both;
}
div[data-testid="stFileUploader"] section:hover {
    border-color: #1e40af !important;
    background: #eff6ff !important;
    box-shadow: 0 10px 28px rgba(30, 64, 175, 0.12);
    transform: translateY(-2px);
}

/* ---------- Stat cards ---------- */
.stat-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(190px, 1fr));
    gap: 16px;
    margin: 14px 0 22px 0;
}
.stat-card {
    position: relative;
    overflow: hidden;
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 14px;
    padding: 18px 20px;
    box-shadow: 0 1px 2px rgba(15, 23, 42, 0.05);
    animation: fadeUp 0.6s cubic-bezier(0.2, 0.8, 0.2, 1) both;
    transition: transform 0.3s ease, box-shadow 0.3s ease;
}
.stat-card::before {
    content: "";
    position: absolute;
    left: 0;
    top: 0;
    bottom: 0;
    width: 4px;
    background: linear-gradient(180deg, #1e40af, #0ea5e9);
}
.stat-card:hover {
    transform: translateY(-4px);
    box-shadow: 0 14px 30px rgba(30, 64, 175, 0.14);
}
.stat-value {
    font-size: 1.9rem;
    font-weight: 800;
    color: #0f172a;
    letter-spacing: -0.5px;
}
.stat-label {
    font-size: 0.72rem;
    font-weight: 700;
    color: #64748b;
    text-transform: uppercase;
    letter-spacing: 0.6px;
    margin-top: 2px;
}

/* ---------- Table & alerts ---------- */
div[data-testid="stDataFrame"] {
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0 4px 18px rgba(15, 23, 42, 0.07);
    animation: fadeUp 0.7s ease both;
}
div[data-testid="stAlert"] {
    border-radius: 12px;
    animation: slideIn 0.45s ease both;
}
div[data-testid="stExpander"] {
    border-radius: 12px;
    background: #ffffff;
    transition: box-shadow 0.3s ease;
}
div[data-testid="stExpander"]:hover {
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08);
}
div[data-testid="stProgress"] > div > div > div {
    background: linear-gradient(90deg, #1e40af, #0ea5e9, #1e40af) !important;
    background-size: 200% 100%;
    animation: shimmer 1.6s linear infinite;
}

/* ---------- Respect "reduce motion" settings ---------- */
@media (prefers-reduced-motion: reduce) {
    *, *::before, *::after {
        animation: none !important;
        transition: none !important;
    }
}
</style>
"""


def inject_css():
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def hero(title, subtitle, logo="📊"):
    st.markdown(
        f'<div class="hero"><div class="hero-logo">{logo}</div>'
        f'<div class="hero-title">{title}</div>'
        f'<div class="hero-sub">{subtitle}</div></div>',
        unsafe_allow_html=True,
    )


def section_title(number, text):
    st.markdown(
        f'<div class="section-title"><span class="step-badge">{number}</span>{text}</div>',
        unsafe_allow_html=True,
    )


def stat_cards(items):
    cards = "".join(
        f'<div class="stat-card" style="animation-delay:{i * 0.12:.2f}s">'
        f'<div class="stat-value">{value}</div>'
        f'<div class="stat-label">{label}</div></div>'
        for i, (label, value) in enumerate(items)
    )
    st.markdown(f'<div class="stat-grid">{cards}</div>', unsafe_allow_html=True)


# ============================================================
# UI
# ============================================================
st.set_page_config(page_title="Excel Data Processor Pro", page_icon="📊", layout="wide")
inject_css()

for key, default in {"logged_in": False, "processed": None, "uploader_key": 0}.items():
    st.session_state.setdefault(key, default)


def login_screen():
    _, col, _ = st.columns([1, 1.1, 1])
    with col:
        hero("Excel Data Processor Pro", "Secure sign-in")
        st.markdown('<div class="login-caption">Sign in to continue</div>', unsafe_allow_html=True)
        with st.form("login_form"):
            username = st.text_input("Username", placeholder="Enter username")
            password = st.text_input("Password", type="password", placeholder="Enter password")
            submitted = st.form_submit_button("Sign In", type="primary", use_container_width=True)
        if submitted:
            if username.strip() == LOGIN_USERNAME and password == LOGIN_PASSWORD:
                st.session_state.logged_in = True
                st.toast("Welcome back, BOOTS 👋", icon="✅")
                st.rerun()
            else:
                st.error("Invalid username or password.")


def app_screen():
    head_left, head_right = st.columns([5, 1])
    with head_left:
        hero("Excel Data Processor Pro",
             "Technology extraction & data transformation · Signed in as BOOTS")
    with head_right:
        st.write("")
        if st.button("Sign out"):
            st.session_state.logged_in = False
            st.session_state.processed = None
            st.rerun()

    section_title(1, "Upload source file")
    uploaded = st.file_uploader(
        "Drop your Excel or CSV file here",
        type=["xlsx", "xls", "csv"],
        key=f"uploader_{st.session_state.uploader_key}",
    )

    if uploaded is None:
        st.info("Upload an Excel or CSV file to get started.")
        return

    try:
        headers, data = load_source(uploaded.getvalue(), uploaded.name)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return

    st.success(f"Loaded {len(data)} rows. Click 'Process Data' to begin.")
    with st.expander("Preview source (first 5 rows)"):
        st.dataframe(pd.DataFrame(data[:5], columns=headers), hide_index=True)

    section_title(2, "Process and export")
    col1, col2, col3 = st.columns(3)

    if col1.button("⚙️ Process Data", type="primary", use_container_width=True):
        total = len(data)
        progress = st.progress(0, text="Starting...")
        output = []
        step = max(1, total // 25)
        for i, row in enumerate(data, start=1):
            output.extend(process_rows([row]))
            if i % step == 0 or i == total:
                progress.progress(i / total if total else 1.0,
                                  text=f"Processing row {i} of {total}")
        progress.empty()
        st.session_state.processed = pd.DataFrame(output, columns=OUTPUT_COLUMNS)
        st.toast("Processing complete", icon="🎉")

    processed = st.session_state.processed
    if processed is not None:
        col2.download_button(
            "⬇️ Download Excel",
            data=to_excel_bytes(processed),
            file_name=f"excel-processor_{date.today():%Y-%m-%d}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    if col3.button("🔄 Reset", use_container_width=True):
        st.session_state.processed = None
        st.session_state.uploader_key += 1
        st.toast("Ready for a new file", icon="🔄")
        st.rerun()

    if processed is not None:
        src_rows = len(data)
        out_rows = len(processed)
        components = processed["Component"].nunique()
        ratio = round(out_rows / src_rows, 1) if src_rows else 0

        stat_cards([
            ("Source Rows", src_rows),
            ("Output Rows", out_rows),
            ("Components", components),
            ("Expansion Ratio", f"{ratio}x"),
        ])

        section_title(3, "Processed Results")
        st.dataframe(processed, hide_index=True, use_container_width=True)


if st.session_state.logged_in:
    app_screen()
else:
    login_screen()
