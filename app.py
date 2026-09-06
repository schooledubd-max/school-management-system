# -*- coding: utf-8 -*-
"""
School Manager BD — Multi-Tenant School Management & Result System
=========================================================================
A single-file Streamlit application that connects to a Google Sheets
backend (the same structure as ResultManager.xlsx) and implements:

  1) Live subject-wise Mark -> Percentage/GPA/Grade/Pass-Fail engine
  2) Automated Result compilation (Total, %, Final GPA/Grade, Merit Position)
  3) Consolidated (weighted, multi-exam) Annual Results
  4) Automatic Promotion + Next-Class Roll generation
  5) Multi-tenant isolation (every table is scoped by SchoolID)
  6) Role-based access: SuperAdmin / Admin(Headteacher) / Teacher / Clerk
  7) Locked-until-published results
  8) Print-ready HTML/CSS Marksheets, Admit Cards & Routines

SETUP (read this first)
------------------------
1. Google Cloud Console -> create a Service Account -> enable
   "Google Sheets API" and "Google Drive API".
2. Create a JSON key for that service account.
3. Open your Google Sheet (the one mirroring ResultManager.xlsx) and
   "Share" it with the service account's email (Editor access).
4. In Streamlit, create `.streamlit/secrets.toml` with:

    SHEET_ID = "your-google-sheet-id-from-the-url"
    SUPERADMIN_PASSWORD = "choose-a-strong-password"

    [gcp_service_account]
    type = "service_account"
    project_id = "..."
    private_key_id = "..."
    private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
    client_email = "...@...iam.gserviceaccount.com"
    client_id = "..."
    auth_uri = "https://accounts.google.com/o/oauth2/auth"
    token_uri = "https://oauth2.googleapis.com/token"
    auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
    client_x509_cert_url = "..."

5. `pip install -r requirements.txt`
6. `streamlit run app.py`

Deploy on Streamlit Community Cloud (free) or any server, then wrap the
live HTTPS URL with PWABuilder.com or Median.co to publish on Play Store.
"""

import time
import uuid
import base64
from datetime import datetime, date, timedelta

import numpy as np
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# =============================================================================
# APP CONFIG
# =============================================================================
st.set_page_config(
    page_title="School Manager BD | School Result & Management System",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
)

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

DEFAULT_CLASS_ORDER = [
    "Play", "Nursery", "KG", "Class-1", "Class-2", "Class-3", "Class-4",
    "Class-5", "Class-6", "Class-7", "Class-8", "Class-9", "Class-10",
    "Class-11", "Class-12",
]

ROLE_SUPERADMIN = "SuperAdmin"
ROLE_ADMIN = "Admin"          # Headteacher / School Admin
ROLE_TEACHER = "Teacher"
ROLE_CLERK = "Clerk"
ROLE_STAFF = ""                # blank Role cell = ordinary staff, view-only, per spec
ADMIN_LIKE_ROLES = (ROLE_ADMIN, "Headmaster", "Headteacher")
ALL_STAFF_ROLES = [ROLE_ADMIN, ROLE_TEACHER, ROLE_CLERK, "Headmaster", "Headteacher"]

# Sheets that are scoped per-school (every one of these has a SchoolID column)
TENANT_SHEETS = [
    "Teachers", "JobApplications", "Students", "Applications", "Marks",
    "Results", "ConsolidatedResults", "Exams", "Routines",
    "SeatPlans_Attendance", "ExamDuties", "ScriptDistribution", "Subjects",
    "GradeScale", "ClassCategory", "Notices", "ContactMessages", "Settings",
    "UserSession", "Subscriptions", "PaymentLogs",
]

# Extra columns this app manages that may not exist yet in an older sheet.
# ensure_headers() will silently add them without touching existing data.
REQUIRED_EXTRA_HEADERS = {
    "Results": ["Published", "PublishedDate", "PublishedBy"],
    "ConsolidatedResults": ["Published", "PublishedDate", "PublishedBy"],
    "Marks": ["EnteredBy", "EnteredDate"],
    "Teachers": ["Password"],
    "ExamDuties": ["ExamDate"],
    "Students": ["Religion"],
}

# =============================================================================
# STYLING — "beautiful" themed UI
# =============================================================================
CUSTOM_CSS = """
<style>
    /* ============================================================
       CRITICAL FIX: many Android phones (and Chrome's own "Force Dark
       Mode for web contents" feature) auto-invert/auto-darken pages
       that don't explicitly declare their color scheme. That is what
       was making input boxes render "সাদা লেখা, সাদা পেজ" (white text
       on a white/blank page) even though our own CSS said otherwise —
       the browser was repainting OVER our colors. Declaring
       color-scheme: light tells the browser/OS this page is
       intentionally light, so it stops guessing and stops repainting
       form fields. This one rule fixes the vast majority of the
       "text is invisible" reports.
       ============================================================ */
    :root, html, body, .stApp { color-scheme: light !important; }
    input, textarea, select, button {
        color-scheme: light !important;
        -webkit-text-fill-color: #0f172a !important;
        caret-color: #0f172a !important;
    }

    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    /* IMPORTANT: do NOT hide the whole <header>, its collapse/expand arrow
       for the sidebar lives there — hiding/shrinking it made the sidebar
       toggle impossible to see or tap once the sidebar was closed on
       mobile. Give it normal height and keep it fully visible/on-top. */
    header[data-testid="stHeader"] {
        background: transparent !important; height: 3.2rem !important;
        overflow: visible !important; z-index: 999997 !important;
    }
    [data-testid="stToolbar"] { visibility: hidden !important; }

    /* The sidebar open/close toggle — cover every Streamlit version's
       testid for it, pin it on-screen, and make it big + colorful so
       it is unmistakably tappable (fixes "swipe বন্ধ করলে আর মেনু
       ফিরে আসে না"). */
    [data-testid="collapsedControl"], [data-testid="stSidebarCollapsedControl"],
    [data-testid="stSidebarCollapseButton"], [data-testid="baseButton-headerNoPadding"],
    button[kind="header"], [data-testid="stSidebarNavCollapseIcon"] {
        visibility: visible !important; display: flex !important; opacity: 1 !important;
        position: fixed !important; top: 10px !important; left: 10px !important;
        background: #2563eb !important; color: #ffffff !important;
        border-radius: 10px !important; box-shadow: 0 4px 12px rgba(0,0,0,0.25) !important;
        z-index: 999999 !important; padding: 6px !important;
    }
    [data-testid="collapsedControl"] svg, [data-testid="stSidebarCollapsedControl"] svg,
    button[kind="header"] svg { fill: #ffffff !important; color: #ffffff !important; }
    /* Push page content down so it never sits under the fixed toggle button */
    .block-container { padding-top: 3.2rem !important; }

    /* --------------------------------------------------------------
       Also force readable dark text everywhere by default (belt and
       braces alongside color-scheme above); specific colored elements
       below (hero banner, stat numbers, badges, table headers,
       sidebar) use more specific selectors so they still win.
       -------------------------------------------------------------- */
    .stApp, .stApp * { color: #0f172a; }
    .stApp { background: linear-gradient(180deg, #f4f7fb 0%, #eef2f9 100%); }

    [data-testid="stWidgetLabel"] p, [data-testid="stWidgetLabel"] label {
        color: #0f172a !important; font-weight: 700 !important; font-size: 0.95rem !important;
    }
    [data-testid="stCaptionContainer"] p { color: #475569 !important; }
    .stTextInput input, .stNumberInput input, .stTextArea textarea,
    .stDateInput input, .stTimeInput input, input, textarea {
        color: #0f172a !important; background: #ffffff !important; border: 1px solid #cbd5e1 !important;
    }
    .stSelectbox [data-baseweb="select"] > div, .stMultiSelect [data-baseweb="select"] > div {
        color: #0f172a !important; background: #ffffff !important; border: 1px solid #cbd5e1 !important;
    }
    .stTabs [data-baseweb="tab"] p { color: #475569 !important; font-weight: 600 !important; }
    .stTabs [aria-selected="true"] p { color: #2563eb !important; font-weight: 800 !important; }
    ::placeholder { color: #94a3b8 !important; opacity: 1 !important; }

    /* Dropdown / multiselect popovers render in a portal OUTSIDE .stApp,
       so they need their own (unscoped) override or their text can be
       invisible (white-on-white) too. */
    div[data-baseweb="popover"], div[data-baseweb="menu"], ul[role="listbox"],
    li[role="option"], div[data-baseweb="select"] {
        background: #ffffff !important;
    }
    div[data-baseweb="popover"] *, div[data-baseweb="menu"] *, ul[role="listbox"] *,
    li[role="option"] *, div[data-baseweb="select"] * {
        color: #0f172a !important;
    }
    li[role="option"][aria-selected="true"] { background: #dbeafe !important; }
    li[role="option"]:hover { background: #eff6ff !important; }

    /* Radio / checkbox option text */
    .stRadio label p, .stCheckbox label p, .stRadio [data-testid="stMarkdownContainer"] p {
        color: #0f172a !important;
    }
    /* Streamlit dataframes / tables */
    [data-testid="stDataFrame"], [data-testid="stDataFrame"] * { color: #0f172a !important; }
    [data-testid="stTable"] table, [data-testid="stTable"] * { color: #0f172a !important; }

    .em-hero {
        background: linear-gradient(120deg, #1e3a8a 0%, #2563eb 55%, #0ea5e9 100%);
        padding: 28px 32px; border-radius: 18px; color: white;
        margin-bottom: 22px; box-shadow: 0 10px 30px rgba(30,58,138,0.25);
    }
    .em-hero h1, .em-hero * { margin: 0; font-size: 1.7rem; font-weight: 800; color: white !important; }
    .em-hero p { margin: 4px 0 0 0; opacity: 0.92; font-size: 0.95rem; }

    .em-card {
        background: white; border-radius: 16px; padding: 20px 22px;
        box-shadow: 0 4px 18px rgba(15, 23, 42, 0.06);
        border: 1px solid rgba(15,23,42,0.05); margin-bottom: 16px;
    }
    .em-hint {
        background: #eff6ff; border-left: 4px solid #2563eb; border-radius: 8px;
        padding: 8px 12px; font-size: 0.85rem; color: #1e3a8a !important; margin: 4px 0 14px 0;
    }
    .em-hint * { color: #1e3a8a !important; }
    .em-stat {
        background: white; border-radius: 14px; padding: 16px 18px;
        border-left: 5px solid #2563eb; box-shadow: 0 4px 14px rgba(15,23,42,0.05);
    }
    .em-stat .val, .em-stat .val * { font-size: 1.6rem; font-weight: 800; color: #1e3a8a !important; }
    .em-stat .lbl, .em-stat .lbl * { font-size: 0.82rem; color: #64748b !important; font-weight: 600; text-transform: uppercase; letter-spacing: .04em;}

    .em-badge, .em-badge * { display:inline-block; padding: 3px 12px; border-radius: 999px; font-size: 0.78rem; font-weight: 700; }
    .em-pass, .em-pass * { background:#dcfce7; color:#15803d !important; }
    .em-fail, .em-fail * { background:#fee2e2; color:#b91c1c !important; }
    .em-pending, .em-pending * { background:#fef9c3; color:#a16207 !important; }

    div[data-testid="stSidebar"] { background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%); }
    div[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
    div[data-testid="stSidebar"] .stButton>button {
        background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.12);
        width: 100%; text-align: left; border-radius: 10px; margin-bottom: 4px;
    }
    div[data-testid="stSidebar"] .stButton>button:hover { background: #2563eb; border-color:#2563eb; }
    /* Whole dashboard stat cards are now real buttons (see stat_card(),
       which renders them inside st.container(key="statcard-...")) — style
       them to clearly look like a card AND look clickable. */
    div[class*="st-key-statcard"] button {
        background: white !important; border-radius: 14px !important; padding: 16px 12px !important;
        border: 1px solid rgba(15,23,42,0.08) !important; border-left: 5px solid #2563eb !important;
        box-shadow: 0 4px 14px rgba(15,23,42,0.05) !important; color: #1e3a8a !important;
        font-size: 1.05rem !important; font-weight: 800 !important; text-align: left !important;
        white-space: pre-line !important; line-height: 1.4 !important;
    }
    div[class*="st-key-statcard"] button:hover {
        border-left: 5px solid #0ea5e9 !important; box-shadow: 0 6px 18px rgba(15,23,42,0.12) !important;
    }
    div[class*="st-key-statcard"] button p { color: #1e3a8a !important; font-weight: 800 !important; white-space: pre-line !important; }

    .stButton>button {
        border-radius: 10px; font-weight: 600; padding: 0.5rem 1.1rem; color: #0f172a;
    }
    .stButton>button[kind="primary"] { background: #2563eb; color: white !important; }
    .stButton>button[kind="primary"] * { color: white !important; }

    .marksheet, .marksheet * { color: #0f172a; }
    .marksheet {
        background: white; padding: 26px 34px; border-radius: 10px;
        border: 2px solid #1e3a8a; font-family: 'Georgia', serif;
    }
    .marksheet h2 { text-align:center; color:#1e3a8a !important; margin: 2px 0; }
    .marksheet .sub, .marksheet .sub * { text-align:center; color:#475569 !important; font-size: 0.9rem; margin-bottom: 10px;}
    .ms-table { width:100%; border-collapse: collapse; margin-top: 14px; }
    .ms-table th, .ms-table td { border: 1px solid #94a3b8; padding: 7px 10px; font-size: 0.92rem; }
    .ms-table th, .ms-table th * { background: #1e3a8a; color: white !important; }
    .ms-table tr:nth-child(even) { background: #f1f5f9; }

    @media print {
        div[data-testid="stSidebar"], .stButton, header, footer { display:none !important; }
        .em-hero { display:none !important; }
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def hint(text: str):
    """Small blue instruction box — used under form headings across every
    section to tell the user exactly what to enter, as required."""
    st.markdown(f'<div class="em-hint">ℹ️ {text}</div>', unsafe_allow_html=True)


def hero(title: str, subtitle: str = ""):
    st.markdown(
        f"""<div class="em-hero"><h1>🎓 {title}</h1><p>{subtitle}</p></div>""",
        unsafe_allow_html=True,
    )


def admob_placeholder(slot: str = "banner"):
    """Google AdMob monetization PLACEHOLDER. This is a Streamlit web app, so
    a real AdMob SDK banner can't render here directly — but when this app
    is wrapped with Capacitor/Median/PWABuilder into an Android APK for the
    Play Store, replace this placeholder <div> with the native AdMob banner
    view (its id is kept stable as 'admob-banner-slot' for that purpose)."""
    if str(st.secrets.get("ADMOB_ENABLED", "")).lower() not in ("1", "true", "yes"):
        return
    st.markdown(
        f"""<div id="admob-{slot}-slot" style="margin:14px 0;padding:10px;text-align:center;
             border:1px dashed #94a3b8;border-radius:8px;color:#94a3b8;font-size:0.78rem;
             background:#f8fafc;">📢 Ad Space ({slot}) — Play Store APK-এ এখানে আসল AdMob ব্যানার বসবে</div>""",
        unsafe_allow_html=True,
    )


def stat_card(col, label, value, nav_key=None):
    """One big, unmistakably-clickable stat card. Uses a keyed container so
    our CSS (div[class*="st-key-statcard"]) can style the *actual* <button>
    — tapping anywhere on the number or label jumps straight to that
    section with full detail (fixes 'ক্লিক করলে কিছু আসে না')."""
    safe_key = f"statcard-{(nav_key or 'x')}-{label}".lower().replace(" ", "-")
    if nav_key:
        try:
            box = col.container(key=safe_key)
        except TypeError:
            # Older Streamlit without container(key=...) support — still
            # fully clickable, just without the extra CSS card styling.
            box = col
        clicked = box.button(f"{value}\n{label}", key=f"btn_{safe_key}", use_container_width=True)
        if clicked:
            st.session_state["nav"] = nav_key
            st.rerun()
    else:
        col.markdown(
            f"""<div class="em-stat"><div class="val">{value}</div><div class="lbl">{label}</div></div>""",
            unsafe_allow_html=True,
        )


def badge(text, kind="pending"):
    cls = {"pass": "em-pass", "fail": "em-fail", "pending": "em-pending"}.get(kind, "em-pending")
    return f'<span class="em-badge {cls}">{text}</span>'


def logo_tag(school_row, height=64):
    """Render the school logo (works with a plain image URL OR a base64 data
    string saved in the 'Logo' column) for use on every printed document.
    Returns '' safely if no logo is set, so templates never break."""
    logo = str(school_row.get("Logo", "") or "").strip()
    if not logo:
        return ""
    src = logo if logo.startswith("http") or logo.startswith("data:") else f"data:image/png;base64,{logo}"
    return f'<img src="{src}" style="height:{height}px;max-width:160px;object-fit:contain;" />'


def doc_header(school_row, title, subtitle=""):
    """Shared header block (logo + school name + address/EIIN + document
    title) used across Marksheet / Admit Card / Certificates / Guard List /
    Seat Plan / Script Register so every printed paper looks consistent and
    always carries the school logo, as required."""
    logo = logo_tag(school_row)
    return f"""
    <div style="display:flex;align-items:center;justify-content:center;gap:14px;margin-bottom:6px;">
        {logo}
        <div>
            <h2 style="margin:0;">{school_row.get('SchoolName','')}</h2>
            <div class="sub">{school_row.get('Address','')} &nbsp;|&nbsp; EIIN: {school_row.get('SchoolEIIN','N/A')}</div>
        </div>
    </div>
    <h3 style="text-align:center;text-decoration:underline;margin-top:2px;">{title}</h3>
    {f'<div class="sub" style="text-align:center;">{subtitle}</div>' if subtitle else ''}
    """


def student_eligible_for_subject(student_row: pd.Series, subject_row: pd.Series) -> bool:
    """Group + Religion aware subject eligibility check — used everywhere a
    subject's student-list is built, so Marks Entry / Marksheets automatically
    show only 'যার যার' (each student's own) Group and Religion subjects,
    never every subject configured for the class."""
    subj_group = str(subject_row.get("Group", "") or "").strip()
    if subj_group and subj_group.lower() not in ("core", "all", "general", "common"):
        if subj_group != str(student_row.get("Group", "") or "").strip():
            return False
    is_religion = str(subject_row.get("IsReligion", "") or "").strip().lower() == "yes"
    if is_religion:
        appl = str(subject_row.get("ApplicableReligion", "") or "").strip()
        stu_religion = str(student_row.get("Religion", "") or "").strip()
        if appl and stu_religion and appl != stu_religion:
            return False
    return True


# =============================================================================
# GOOGLE SHEETS CONNECTION LAYER
# =============================================================================
@st.cache_resource(show_spinner=False)
def get_client():
    try:
        info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(info, scopes=SCOPES)
        return gspread.authorize(creds)
    except Exception as e:
        st.error(
            "⚠️ Google Sheets credentials not configured. Please add "
            "`gcp_service_account` and `SHEET_ID` to `.streamlit/secrets.toml`. "
            f"(Detail: {e})"
        )
        st.stop()


@st.cache_resource(show_spinner=False)
def get_spreadsheet():
    client = get_client()
    try:
        return client.open_by_key(st.secrets["SHEET_ID"])
    except Exception as e:
        st.error(f"⚠️ Could not open the Google Sheet. Check SHEET_ID and sharing permissions. ({e})")
        st.stop()


def get_ws(sheet_name: str):
    sh = get_spreadsheet()
    try:
        return sh.worksheet(sheet_name)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=sheet_name, rows=200, cols=26)
        return ws


def ensure_headers(sheet_name: str):
    """Make sure any extra columns this app needs exist, without deleting data."""
    extra = REQUIRED_EXTRA_HEADERS.get(sheet_name)
    if not extra:
        return
    ws = get_ws(sheet_name)
    header = ws.row_values(1)
    missing = [h for h in extra if h not in header]
    if missing:
        new_header = header + missing
        ws.update("A1", [new_header])


if "data_version" not in st.session_state:
    st.session_state["data_version"] = 0


def bump_version():
    st.session_state["data_version"] += 1
    _fetch_values.clear()


@st.cache_data(show_spinner=False)
def _fetch_values(sheet_name: str, version: int):
    ws = get_ws(sheet_name)
    return ws.get_all_values()


def read_df(sheet_name: str, school_id: str = None) -> pd.DataFrame:
    """Read a whole sheet as a DataFrame (all-string cells preserved).
    Pass school_id to auto-filter tenant-scoped sheets."""
    ensure_headers(sheet_name)
    values = _fetch_values(sheet_name, st.session_state["data_version"])
    if not values:
        return pd.DataFrame()
    header, rows = values[0], values[1:]
    # pad short rows
    rows = [r + [""] * (len(header) - len(r)) for r in rows]
    df = pd.DataFrame(rows, columns=header)
    if school_id and "SchoolID" in df.columns:
        df = df[df["SchoolID"] == school_id].reset_index(drop=True)
    return df


def num(series_or_val, default=0.0):
    """Safe numeric coercion (sheet cells come back as text)."""
    if isinstance(series_or_val, pd.Series):
        return pd.to_numeric(series_or_val, errors="coerce").fillna(default)
    try:
        if series_or_val in (None, ""):
            return default
        return float(series_or_val)
    except (TypeError, ValueError):
        return default


def _clean_cell(v):
    """Format a cell for writing back to Sheets — avoids '5.0' / 'nan' noise."""
    if v is None:
        return ""
    if isinstance(v, float):
        if pd.isna(v):
            return ""
        if v == int(v):
            return str(int(v))
        return str(round(v, 2))
    if isinstance(v, bool):
        return "Yes" if v else "No"
    return str(v)


def write_full_sheet(sheet_name: str, full_df: pd.DataFrame, headers: list):
    """Overwrite an entire sheet's contents with full_df (already contains
    ALL schools' rows — used after merging in changes for one tenant)."""
    ws = get_ws(sheet_name)
    full_df = full_df.reindex(columns=headers)
    full_df = full_df.map(_clean_cell) if hasattr(full_df, "map") else full_df.applymap(_clean_cell)
    values = [headers] + full_df.values.tolist()
    ws.clear()
    ws.update("A1", values)
    bump_version()


def upsert_rows(sheet_name: str, new_rows_df: pd.DataFrame, key_cols: list):
    """Merge new_rows_df into the sheet: any existing row whose key_cols match
    is replaced, everything else (incl. other schools) is preserved, and rows
    with brand-new keys are appended. Single API round trip."""
    ensure_headers(sheet_name)
    ws = get_ws(sheet_name)
    header = ws.row_values(1)
    if not header:
        header = list(new_rows_df.columns)
        ws.update("A1", [header])
    existing = read_df(sheet_name)
    new_rows_df = new_rows_df.reindex(columns=header).fillna("")
    if existing.empty:
        combined = new_rows_df
    else:
        existing_keyed = existing.set_index(key_cols, drop=False)
        new_keyed = new_rows_df.set_index(key_cols, drop=False)
        existing_keyed = existing_keyed[~existing_keyed.index.isin(new_keyed.index)]
        combined = pd.concat([existing_keyed, new_keyed], ignore_index=True)
    write_full_sheet(sheet_name, combined, header)


def append_row(sheet_name: str, row: dict):
    ensure_headers(sheet_name)
    ws = get_ws(sheet_name)
    header = ws.row_values(1)
    if not header:
        header = list(row.keys())
        ws.update("A1", [header])
    ws.append_row([str(row.get(h, "")) for h in header], value_input_option="USER_ENTERED")
    bump_version()


def next_seq_id(school_id: str, sheet_name: str, id_col: str, prefix: str, width: int = 4,
                 include_year: bool = True) -> str:
    """Generate the next sequential ID like '000001-26-S0007' style, scoped to school."""
    df = read_df(sheet_name, school_id=school_id)
    school_num = school_id.replace("SCH-", "")
    if include_year:
        yy = datetime.now().strftime("%y")
        base = f"{school_num}-{yy}-{prefix}"
    else:
        base = f"{school_num}-{prefix}"
    n = 1
    if not df.empty and id_col in df.columns:
        existing = df[id_col].astype(str)
        matching = existing[existing.str.startswith(base)]
        if not matching.empty:
            nums = matching.str.extract(r"(\d+)$")[0].astype(float)
            n = int(nums.max()) + 1 if not nums.isna().all() else 1
    return f"{base}{n:0{width}d}"


# =============================================================================
# SUBSCRIPTION LOCK & PAYMENT APPROVAL
#   স্কুলের সাবস্ক্রিপশন মেয়াদ শেষ হলে Admin/Teacher ফিচার লক হয়ে শুধু
#   Billing পেজ (রিনিউ + বিকাশ/নগদ পেমেন্ট) দেখাবে, SuperAdmin অনুমোদন
#   করলেই আবার আনলক হবে।
# =============================================================================
def get_subscription_status(school_id: str):
    """Returns (is_locked: bool, latest_sub_row_or_None, human_message: str).
    A school with NO subscription rows at all is treated as unlocked (grace
    period) so existing/older schools never get accidentally locked out."""
    subs = read_df("Subscriptions", school_id=school_id)
    if subs.empty:
        return False, None, "কোনো সাবস্ক্রিপশন রেকর্ড এখনো তৈরি হয়নি (Grace period — আনলকড)।"
    subs = subs.copy()
    subs["_exp"] = pd.to_datetime(subs.get("ExpiryDate", ""), errors="coerce")
    subs = subs.sort_values("_exp", ascending=False)
    latest = subs.iloc[0]
    if pd.isna(latest["_exp"]):
        return False, latest, "মেয়াদের তারিখ পাওয়া যায়নি — আনলকড ধরা হচ্ছে।"
    today = pd.Timestamp(datetime.now().date())
    if str(latest.get("Status", "")) == "Active" and latest["_exp"].date() >= today.date():
        return False, latest, f"সক্রিয় সাবস্ক্রিপশন — মেয়াদ শেষ হবে {latest.get('ExpiryDate','')} তারিখে।"
    if latest["_exp"].date() < today.date():
        return True, latest, f"সাবস্ক্রিপশনের মেয়াদ {latest.get('ExpiryDate','')} তারিখে শেষ হয়ে গেছে।"
    return False, latest, f"সক্রিয় — মেয়াদ শেষ হবে {latest.get('ExpiryDate','')} তারিখে।"


def approve_payment(payment_row: pd.Series):
    """SuperAdmin approves a Pending bKash/Nagad/Rocket claim: marks the
    payment SUCCESS and extends (or creates) that school's subscription."""
    school_id = payment_row.get("SchoolID", "")
    plan_days = {"Monthly": 30, "Yearly": 365, "Trial": 15}
    pays = read_df("PaymentLogs")
    pays.loc[pays["PaymentID"] == payment_row.get("PaymentID"), "Status"] = "SUCCESS"
    upsert_rows("PaymentLogs", pays[pays["PaymentID"] == payment_row.get("PaymentID")], key_cols=["PaymentID"])

    subs = read_df("Subscriptions", school_id=school_id)
    today = datetime.now().date()
    if not subs.empty:
        subs = subs.copy()
        subs["_exp"] = pd.to_datetime(subs.get("ExpiryDate", ""), errors="coerce")
        latest = subs.sort_values("_exp", ascending=False).iloc[0]
        base_date = latest["_exp"].date() if pd.notna(latest["_exp"]) and latest["_exp"].date() > today else today
        plan = str(latest.get("PlanType", "Monthly")) or "Monthly"
        new_expiry = base_date + timedelta(days=plan_days.get(plan, 30))
        subs_all = read_df("Subscriptions")
        subs_all.loc[subs_all["SubscriptionID"] == latest["SubscriptionID"], ["ExpiryDate", "Status", "LastPaymentDate"]] = [
            new_expiry.strftime("%Y-%m-%d"), "Active", today.strftime("%Y-%m-%d"),
        ]
        upsert_rows("Subscriptions", subs_all[subs_all["SubscriptionID"] == latest["SubscriptionID"]], key_cols=["SubscriptionID"])
    else:
        new_expiry = today + timedelta(days=30)
        subid = f"{school_id.replace('SCH-','')}-{datetime.now().strftime('%y')}-SUBS001"
        append_row("Subscriptions", {
            "SchoolID": school_id, "SubscriptionID": subid, "PlanType": "Monthly",
            "Amount": payment_row.get("Amount", ""), "StartDate": today.strftime("%Y-%m-%d"),
            "ExpiryDate": new_expiry.strftime("%Y-%m-%d"), "Status": "Active",
            "LastPaymentDate": today.strftime("%Y-%m-%d"),
        })
    return new_expiry



def get_class_order(school_id: str) -> list:
    cc = read_df("ClassCategory", school_id=school_id)
    if not cc.empty and "Class" in cc.columns:
        seen, order = set(), []
        for c in cc["Class"].tolist():
            if c and c not in seen:
                seen.add(c)
                order.append(c)
        if order:
            return order
    return DEFAULT_CLASS_ORDER


def next_class(school_id: str, current_class: str) -> str:
    order = get_class_order(school_id)
    if current_class in order:
        idx = order.index(current_class)
        if idx + 1 < len(order):
            return order[idx + 1]
        return "Graduated"  # last class -> passing out
    return "N/A"


def get_gradescale(school_id: str) -> pd.DataFrame:
    gs = read_df("GradeScale", school_id=school_id).copy()
    if gs.empty:
        # sensible fallback matching Bangladesh national grading
        gs = pd.DataFrame([
            {"Grade": "A+", "GPA": 5.0, "MinPercentage": 80, "MaxPercentage": 100},
            {"Grade": "A", "GPA": 4.0, "MinPercentage": 70, "MaxPercentage": 79.99},
            {"Grade": "A-", "GPA": 3.5, "MinPercentage": 60, "MaxPercentage": 69.99},
            {"Grade": "B", "GPA": 3.0, "MinPercentage": 50, "MaxPercentage": 59.99},
            {"Grade": "C", "GPA": 2.0, "MinPercentage": 40, "MaxPercentage": 49.99},
            {"Grade": "D", "GPA": 1.0, "MinPercentage": 33, "MaxPercentage": 39.99},
            {"Grade": "F", "GPA": 0.0, "MinPercentage": 0, "MaxPercentage": 32.99},
        ])
    gs["GPA"] = num(gs["GPA"])
    gs["MinPercentage"] = num(gs["MinPercentage"])
    gs["MaxPercentage"] = num(gs["MaxPercentage"])
    return gs.sort_values("MinPercentage", ascending=False).reset_index(drop=True)


def grade_from_percentage(gs_df: pd.DataFrame, pct: float):
    """Return (grade, gpa) for a percentage using the school's grade scale."""
    if pct is None or pd.isna(pct):
        return "F", 0.0
    pct = max(0.0, min(100.0, float(pct)))
    for _, row in gs_df.iterrows():
        if row["MinPercentage"] <= pct <= row["MaxPercentage"] + 0.001:
            return row["Grade"], float(row["GPA"])
    # fallback: below lowest band
    lowest = gs_df.iloc[-1]
    if pct < lowest["MinPercentage"]:
        return "F", 0.0
    return lowest["Grade"], float(lowest["GPA"])


def compute_subject_result(subject_row: pd.Series, written: float, mcq: float, practical: float,
                            is_present: bool, gs_df: pd.DataFrame) -> dict:
    """THE LIVE GRADING ENGINE — runs the instant a teacher enters marks.
    Supports both 'Individual' (separate CQ/MCQ/Practical pass thresholds)
    and 'Combined' (single total pass-mark) policies, exactly as defined
    per-subject in the Subjects sheet."""
    full_mark = num(subject_row.get("FullMark"), 0)
    policy = str(subject_row.get("PassPolicy", "Combined")).strip() or "Combined"

    if not is_present:
        return {"ObtainedMarks": 0, "FullMarks": full_mark, "Percentage": 0.0,
                "SubjectGPA": 0.0, "Grade": "F", "IsPass": "No", "IsPresent": "No"}

    written = num(written, 0)
    mcq = num(mcq, 0)
    practical = num(practical, 0)
    obtained = written + mcq + practical

    if policy == "Individual":
        written_pass = num(subject_row.get("WrittenPass"), 0)
        mcq_pass = num(subject_row.get("MCQPass"), 0)
        practical_full = num(subject_row.get("PracticalFull"), 0)
        practical_pass = num(subject_row.get("PracticalPass"), 0)
        ok = (written >= written_pass) and (mcq >= mcq_pass)
        if practical_full > 0:
            ok = ok and (practical >= practical_pass)
    else:  # Combined
        combined_pass = num(subject_row.get("CombinedPassMark"), 0)
        ok = obtained >= combined_pass

    pct = (obtained / full_mark * 100) if full_mark > 0 else 0.0

    if not ok:
        grade, gpa = "F", 0.0
    else:
        grade, gpa = grade_from_percentage(gs_df, pct)
        if gpa <= 0:  # percentage band itself is failing even though threshold passed
            ok = False

    return {
        "ObtainedMarks": round(obtained, 2), "FullMarks": full_mark,
        "Percentage": round(pct, 2), "SubjectGPA": gpa, "Grade": grade,
        "IsPass": "Yes" if ok else "No", "IsPresent": "Yes",
    }


# =============================================================================
# AUTH / SESSION
# =============================================================================
def init_session():
    defaults = {
        "logged_in": False, "role": None, "school_id": None, "school_name": None,
        "teacher_id": None, "user_name": None, "guardian_mode": False,
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


def do_logout():
    for k in ["logged_in", "role", "school_id", "school_name", "teacher_id",
              "user_name", "guardian_mode", "guardian_student_id", "guardian_roll"]:
        st.session_state[k] = False if k in ("logged_in", "guardian_mode") else None
    st.rerun()


def login_page():
    hero("School Manager BD", "Multi-Tenant School Management & Result System")
    # Order requested: Student Login -> Teacher/Staff Login -> SuperAdmin Login
    tab_student, tab_school, tab_super = st.tabs(
        ["🎒 Student Login (Check Result)", "🏫 Teacher / Staff Login", "🛡️ SuperAdmin Login"]
    )

    with tab_student:
        st.markdown('<div class="em-card">', unsafe_allow_html=True)
        hint("লগইন ছাড়াই আপনার (বা আপনার সন্তানের) প্রকাশিত রেজাল্ট দেখতে — স্কুল বেছে নিয়ে "
             "Student ID এবং Roll নম্বর দুটোই দিন (দুটো মিলে গেলে তবেই রেজাল্ট দেখাবে, নিরাপত্তার জন্য)।")
        schools = read_df("Schools")
        if not schools.empty:
            options = schools["SchoolName"] + "  (" + schools["SchoolID"] + ")"
            g_choice = st.selectbox("School", options.tolist(), key="g_school",
                                     help="আপনার সন্তান যে স্কুলে পড়ে সেটি বেছে নিন।")
            c1, c2 = st.columns(2)
            g_student_id = c1.text_input(
                "Student ID", placeholder="000001-26-S0001",
                help="মার্কশীট/এডমিট কার্ড/ভর্তি ফর্মে ছাপানো Student ID এখানে হুবহু লিখুন।",
            )
            g_roll = c2.text_input(
                "Roll No.", placeholder="যেমন: 05",
                help="বর্তমান ক্লাসের Roll নম্বর — Student ID-এর সাথে মিলিয়ে যাচাই করা হবে।",
            )
            if st.button("View Result", type="primary"):
                if not g_student_id.strip() or not g_roll.strip():
                    st.error("Student ID এবং Roll — দুটোই লিখুন।")
                else:
                    g_school_id = g_choice.split("(")[-1].rstrip(")")
                    st.session_state.update({
                        "guardian_mode": True, "school_id": g_school_id,
                        "school_name": g_choice.split("  (")[0], "role": "Guardian",
                    })
                    st.session_state["guardian_student_id"] = g_student_id.strip()
                    st.session_state["guardian_roll"] = g_roll.strip()
                    st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_school:
        st.markdown('<div class="em-card">', unsafe_allow_html=True)
        hint("আপনার স্কুল বেছে নিন, তারপর আপনার Teacher/Staff ID ও PIN দিন (এই দুটো আপনার স্কুলের "
             "Headteacher/Admin আপনাকে দিয়েছেন)। ভুলে গেলে Admin-কে জিজ্ঞেস করুন।")
        schools = read_df("Schools")
        if schools.empty:
            st.warning("No schools found in the sheet yet.")
        else:
            active = schools[schools.get("IsActive", "Yes") == "Yes"] if "IsActive" in schools.columns else schools
            options = active["SchoolName"] + "  (" + active["SchoolID"] + ")"
            choice = st.selectbox(
                "Select your School", options.tolist() if not active.empty else [],
                help="তালিকা থেকে আপনার স্কুলের নাম বেছে নিন। (SchoolID বন্ধনীতে দেখানো আছে)",
            )
            teacher_id = st.text_input(
                "Teacher / Staff ID", placeholder="যেমন: 000001-T001",
                help="আপনার স্কুলের Admin আপনাকে যে Teacher ID দিয়েছেন সেটা লিখুন।",
            )
            pin = st.text_input(
                "PIN", type="password", placeholder="৪ সংখ্যার গোপন PIN",
                help="আপনার গোপন PIN নম্বর — কাউকে শেয়ার করবেন না।",
            )
            if st.button("Login", type="primary", key="staff_login"):
                if not choice:
                    st.error("Please select a school.")
                else:
                    school_id = choice.split("(")[-1].rstrip(")")
                    teachers = read_df("Teachers", school_id=school_id)
                    match = teachers[
                        (teachers["TeacherID"].astype(str) == teacher_id.strip())
                        & (teachers["PIN"].astype(str) == pin.strip())
                    ]
                    if match.empty:
                        st.error("Invalid Teacher ID or PIN.")
                    elif str(match.iloc[0].get("IsActive", "Yes")) == "No":
                        st.error("This account has been deactivated. Contact your Admin.")
                    else:
                        row = match.iloc[0]
                        srow = active[active["SchoolID"] == school_id].iloc[0]
                        st.session_state.update({
                            "logged_in": True,
                            "role": row.get("Role", ROLE_TEACHER),
                            "school_id": school_id,
                            "school_name": srow.get("SchoolName", ""),
                            "teacher_id": row.get("TeacherID"),
                            "user_name": row.get("Name"),
                        })
                        st.rerun()
        st.markdown("</div>", unsafe_allow_html=True)

    with tab_super:
        st.markdown('<div class="em-card">', unsafe_allow_html=True)
        hint("এই ট্যাবটি শুধু মূল Super Administrator-এর জন্য (যিনি সব স্কুল পরিচালনা করেন)। "
             "স্কুল স্টাফ/টিচার হলে 'Teacher / Staff Login' ট্যাব ব্যবহার করুন।")
        pw = st.text_input(
            "SuperAdmin Password", type="password", key="sa_pw",
            placeholder="আপনার secrets.toml-এ রাখা SUPERADMIN_PASSWORD",
            help="এটি .streamlit/secrets.toml ফাইলে SUPERADMIN_PASSWORD হিসেবে সেট করা আছে।",
        )
        if st.button("Login as SuperAdmin", type="primary"):
            expected = st.secrets.get("SUPERADMIN_PASSWORD", "")
            if expected and pw == expected:
                st.session_state.update({
                    "logged_in": True, "role": ROLE_SUPERADMIN, "school_id": None,
                    "school_name": "All Schools", "teacher_id": "SUPERADMIN",
                    "user_name": "Super Administrator",
                })
                st.rerun()
            else:
                st.error("Incorrect password.")
        st.markdown("</div>", unsafe_allow_html=True)


# =============================================================================
# RESULT COMPILATION ENGINE
# =============================================================================
def generate_results_for_group(school_id, exam_id, klass, section, session_year, actor_name=""):
    """Aggregate every subject Mark row for (exam, class, section) into one
    Result row per student: totals, percentage, final GPA/Grade, pass/fail,
    and Class Merit Position. Returns the DataFrame written."""
    marks = read_df("Marks", school_id=school_id)
    if marks.empty:
        return pd.DataFrame(), "No marks found for this school yet."

    scope = marks[
        (marks["ExamID"] == exam_id) & (marks["Class"] == klass) & (marks["Section"] == section)
    ].copy()
    if scope.empty:
        return pd.DataFrame(), "No marks entered yet for this Exam / Class / Section."

    for c in ["FullMarks", "ObtainedMarks", "SubjectGPA"]:
        scope[c] = num(scope[c])

    gs_df = get_gradescale(school_id)
    rows = []
    for student_id, g in scope.groupby("StudentID"):
        total_full = g["FullMarks"].sum()
        total_obtained = g["ObtainedMarks"].sum()
        pct = round((total_obtained / total_full * 100), 2) if total_full > 0 else 0.0
        failed_subjects = int((g["IsPass"].astype(str) == "No").sum())
        absent_subjects = int((g["IsPresent"].astype(str) == "No").sum())
        attended = int(len(g) - absent_subjects)

        if failed_subjects > 0:
            final_grade, final_gpa, status = "F", 0.0, "Failed"
        else:
            final_grade, final_gpa = grade_from_percentage(gs_df, pct)
            status = "Passed" if final_gpa > 0 else "Failed"

        rows.append({
            "SchoolID": school_id, "StudentID": student_id, "ExamID": exam_id,
            "Class": klass, "Section": section, "Session": session_year,
            "TotalFullMarks": total_full, "TotalMarks": total_obtained,
            "Percentage": pct, "FinalGPA": final_gpa, "FinalGrade": final_grade,
            "Status": status, "TotalFailedSubjects": failed_subjects,
            "AttendedSubjects": attended, "AbsentSubjects": absent_subjects,
        })

    result_df = pd.DataFrame(rows)
    # ---- Class Merit Position: passed students ranked by GPA desc, marks desc ----
    passed = result_df[result_df["Status"] == "Passed"].sort_values(
        ["FinalGPA", "TotalMarks"], ascending=[False, False]
    ).reset_index(drop=True)
    passed["ClassMeritPosition"] = (passed.index + 1).astype(str)
    failed = result_df[result_df["Status"] != "Passed"].copy()
    failed["ClassMeritPosition"] = "—"
    result_df = pd.concat([passed, failed], ignore_index=True)

    # keep existing ResultID if a row already exists, else mint a new one
    existing_results = read_df("Results", school_id=school_id)
    result_df["ResultID"] = ""
    for i, r in result_df.iterrows():
        if not existing_results.empty:
            m = existing_results[
                (existing_results["StudentID"] == r["StudentID"]) & (existing_results["ExamID"] == exam_id)
            ]
        else:
            m = pd.DataFrame()
        if not m.empty:
            result_df.at[i, "ResultID"] = m.iloc[0]["ResultID"]
            result_df.at[i, "Published"] = m.iloc[0].get("Published", "No")
            result_df.at[i, "PublishedDate"] = m.iloc[0].get("PublishedDate", "")
        else:
            result_df.at[i, "ResultID"] = next_seq_id(school_id, "Results", "ResultID", "R")
            result_df.at[i, "Published"] = "No"
            result_df.at[i, "PublishedDate"] = ""
    result_df["PublishedBy"] = actor_name

    upsert_rows("Results", result_df, key_cols=["StudentID", "ExamID"])
    return result_df, f"✅ Generated results for {len(result_df)} students."


def publish_group(sheet_name, school_id, key_filter: dict, publish: bool, actor_name=""):
    df = read_df(sheet_name, school_id=school_id)
    if df.empty:
        return 0
    mask = pd.Series(True, index=df.index)
    for k, v in key_filter.items():
        mask &= (df[k] == v)
    df.loc[mask, "Published"] = "Yes" if publish else "No"
    df.loc[mask, "PublishedDate"] = datetime.now().strftime("%Y-%m-%d %H:%M") if publish else ""
    df.loc[mask, "PublishedBy"] = actor_name if publish else ""
    key_cols = ["StudentID", "ExamID"] if sheet_name == "Results" else ["StudentID", "Session", "CurrentClass"]
    upsert_rows(sheet_name, df, key_cols=key_cols)
    return int(mask.sum())


# =============================================================================
# CONSOLIDATED (WEIGHTED, MULTI-EXAM) ANNUAL RESULTS + PROMOTION ENGINE
# =============================================================================
def generate_consolidated(school_id, session_year, klass, section, actor_name=""):
    exams = read_df("Exams", school_id=school_id)
    exams = exams[exams["Session"] == session_year].copy()
    if exams.empty:
        return pd.DataFrame(), "No exams configured for this session."
    exams["Weight"] = num(exams["Weight"])
    exams = exams.sort_values("OrderSequence")

    results = read_df("Results", school_id=school_id)
    results = results[(results["Class"] == klass) & (results["Section"] == section) & (results["Session"] == session_year)].copy()
    if results.empty:
        return pd.DataFrame(), "No exam results found yet for this Class / Section / Session."
    for c in ["TotalFullMarks", "TotalMarks"]:
        results[c] = num(results[c])

    students = sorted(results["StudentID"].unique().tolist())
    gs_df = get_gradescale(school_id)
    rows = []
    for sid in students:
        srows = results[results["StudentID"] == sid]
        total_full = srows["TotalFullMarks"].max() if not srows.empty else 0
        exam_cols, weighted_sum, any_fail, exams_counted = {}, 0.0, False, 0
        for i, erow in enumerate(exams.itertuples(), start=1):
            m = srows[srows["ExamID"] == erow.ExamID]
            if m.empty:
                exam_cols[f"Exam{i}_Marks"] = ""
                exam_cols[f"Exam{i}_Weighted"] = ""
                continue
            exams_counted += 1
            marks_val = float(m.iloc[0]["TotalMarks"])
            weighted = round(marks_val * (erow.Weight / 100.0), 2)
            exam_cols[f"Exam{i}_Marks"] = marks_val
            exam_cols[f"Exam{i}_Weighted"] = weighted
            weighted_sum += weighted
            if str(m.iloc[0]["Status"]) != "Passed":
                any_fail = True

        consolidated_total = round(weighted_sum, 2)
        pct = round((consolidated_total / total_full * 100), 2) if total_full > 0 else 0.0
        if any_fail or exams_counted == 0:
            grade, gpa, status = "F", 0.0, "Failed"
        else:
            grade, gpa = grade_from_percentage(gs_df, pct)
            status = "Passed" if gpa > 0 else "Failed"

        promoted = next_class(school_id, klass) if status == "Passed" else klass + " (Retained)"

        row = {
            "SchoolID": school_id, "StudentID": sid, "CurrentClass": klass,
            "Section": section, "Session": session_year, "TotalFullMarks": total_full,
            **exam_cols, "ConsolidatedTotal": consolidated_total,
            "ConsolidatedPercentage": pct, "ConsolidatedGPA": gpa,
            "ConsolidatedGrade": grade, "Status": status, "PromotedClass": promoted,
        }
        rows.append(row)

    cdf = pd.DataFrame(rows)
    # ---- Merit ranking + Next Roll, grouped per PromotedClass + Section ----
    # Roll is generated per-section ("শাখা অনুযায়ী শাখাসহ রোল") so each
    # section restarts its own 1..N roll numbering after promotion.
    cdf["NextRoll"] = ""
    for (pclass, psec), grp in cdf.groupby(["PromotedClass", "Section"]):
        ranked = grp[grp["Status"] == "Passed"].sort_values(
            ["ConsolidatedGPA", "ConsolidatedTotal"], ascending=[False, False]
        )
        for i, idx in enumerate(ranked.index, start=1):
            cdf.at[idx, "NextRoll"] = f"{psec}-{i:02d}"

    existing = read_df("ConsolidatedResults", school_id=school_id)
    cdf["ConsolidatedID"] = ""
    for i, r in cdf.iterrows():
        if not existing.empty:
            m = existing[(existing["StudentID"] == r["StudentID"]) & (existing["Session"] == session_year)
                         & (existing["CurrentClass"] == klass)]
        else:
            m = pd.DataFrame()
        if not m.empty:
            cdf.at[i, "ConsolidatedID"] = m.iloc[0]["ConsolidatedID"]
            cdf.at[i, "Published"] = m.iloc[0].get("Published", "No")
            cdf.at[i, "PublishedDate"] = m.iloc[0].get("PublishedDate", "")
        else:
            cdf.at[i, "ConsolidatedID"] = next_seq_id(school_id, "ConsolidatedResults", "ConsolidatedID", "CR")
            cdf.at[i, "Published"] = "No"
            cdf.at[i, "PublishedDate"] = ""
    cdf["PublishedBy"] = actor_name

    upsert_rows("ConsolidatedResults", cdf, key_cols=["StudentID", "Session", "CurrentClass"])
    return cdf, f"✅ Consolidated results generated for {len(cdf)} students."


def apply_promotion(school_id, session_year, klass, section):
    """Write PromotedClass / NextRoll back onto the Students sheet for the
    next academic year (bumps Class + Roll for students who passed)."""
    cdf = read_df("ConsolidatedResults", school_id=school_id)
    cdf = cdf[(cdf["Session"] == session_year) & (cdf["CurrentClass"] == klass) & (cdf["Section"] == section)]
    if cdf.empty:
        return 0
    students = read_df("Students", school_id=school_id)
    updated = 0
    for _, r in cdf.iterrows():
        mask = students["StudentID"] == r["StudentID"]
        if not mask.any():
            continue
        if r["Status"] == "Passed" and r["NextRoll"]:
            students.loc[mask, "Class"] = r["PromotedClass"]
            students.loc[mask, "Roll"] = r["NextRoll"]
            updated += 1
    upsert_rows("Students", students, key_cols=["StudentID"])
    return updated


# =============================================================================
# PRINT TEMPLATES — Marksheet & Admit Card (HTML/CSS)
# =============================================================================
def render_marksheet(school_row, student_row, result_row, marks_df, subjects_df, exam_name):
    subj_lookup = subjects_df.set_index("SubjectID")["SubjectName"].to_dict() if not subjects_df.empty else {}
    body_rows = ""
    for _, m in marks_df.iterrows():
        sname = subj_lookup.get(m["SubjectID"], m["SubjectID"])
        pass_txt = "Pass" if str(m.get("IsPass")) == "Yes" else "Fail"
        body_rows += (
            f"<tr><td>{sname}</td><td>{m.get('FullMarks','')}</td>"
            f"<td>{m.get('ObtainedMarks','')}</td><td>{m.get('Grade','')}</td>"
            f"<td>{m.get('SubjectGPA','')}</td><td>{pass_txt}</td></tr>"
        )
    status = result_row.get("Status", "")
    status_color = "#15803d" if status == "Passed" else "#b91c1c"
    html = f"""
    <div class="marksheet">
        {doc_header(school_row, f"ACADEMIC MARKSHEET — {exam_name}")}
        <table style="width:100%;margin-top:10px;">
            <tr>
                <td><b>Student:</b> {student_row.get('StudentName','')}</td>
                <td><b>Student ID:</b> {student_row.get('StudentID','')}</td>
            </tr>
            <tr>
                <td><b>Class:</b> {result_row.get('Class','')} &nbsp; <b>Section:</b> {result_row.get('Section','')}</td>
                <td><b>Roll:</b> {student_row.get('Roll','')} &nbsp; <b>Session:</b> {result_row.get('Session','')}</td>
            </tr>
            <tr><td><b>Father:</b> {student_row.get('FatherName','')}</td><td><b>Mother:</b> {student_row.get('MotherName','')}</td></tr>
        </table>
        <table class="ms-table">
            <tr><th>Subject</th><th>Full Marks</th><th>Obtained</th><th>Grade</th><th>GPA</th><th>Result</th></tr>
            {body_rows}
        </table>
        <table style="width:100%;margin-top:14px;">
            <tr>
                <td><b>Total Marks:</b> {result_row.get('TotalMarks','')} / {result_row.get('TotalFullMarks','')}</td>
                <td><b>Percentage:</b> {result_row.get('Percentage','')}%</td>
            </tr>
            <tr>
                <td><b>Final GPA:</b> {result_row.get('FinalGPA','')}</td>
                <td><b>Final Grade:</b> {result_row.get('FinalGrade','')}</td>
            </tr>
            <tr>
                <td><b>Class Merit Position:</b> {result_row.get('ClassMeritPosition','')}</td>
                <td><b>Result:</b> <span style="color:{status_color};font-weight:800;">{status}</span></td>
            </tr>
        </table>
        <div style="display:flex;justify-content:space-between;margin-top:60px;">
            <div>_____________________<br/>Class Teacher</div>
            <div>_____________________<br/>Headteacher</div>
        </div>
    </div>
    """
    return html


def render_admit_card(school_row, student_row, exam_name, routine_df):
    """One admit card, WITH its exam routine printed on the same card (so the
    admit card and routine are always together on paper). Styled like a
    real admit card: photo box, colored ADMIT CARD ribbon, subject-wise
    schedule table and three signature lines."""
    sched_rows = ""
    for _, r in routine_df.sort_values("ExamDate").iterrows() if not routine_df.empty else []:
        sched_rows += (
            f"<tr><td>{r.get('ExamDate','')}</td><td>{r.get('DayName','')}</td>"
            f"<td>{r.get('StartTime','')} - {r.get('EndTime','')}</td>"
            f"<td>{r.get('SubjectID','')}</td><td>{r.get('Classes', r.get('RoomID',''))}</td></tr>"
        )
    photo = str(student_row.get("Photo", "") or "").strip()
    photo_html = (
        f'<img src="{photo}" style="width:80px;height:96px;object-fit:cover;border:1px solid #94a3b8;" />'
        if photo else
        '<div style="width:80px;height:96px;border:1px dashed #94a3b8;display:flex;align-items:center;'
        'justify-content:center;font-size:0.65rem;color:#94a3b8;text-align:center;">Photo</div>'
    )
    html = f"""
    <div class="marksheet" style="max-width:100%;padding:16px 20px;position:relative;">
        <div style="background:linear-gradient(90deg,#1e3a8a,#2563eb);color:white;text-align:center;
                    padding:4px;border-radius:6px;font-weight:800;letter-spacing:2px;margin-bottom:8px;">
            ADMIT CARD &nbsp;•&nbsp; প্রবেশপত্র
        </div>
        {doc_header(school_row, exam_name)}
        <div style="display:flex;justify-content:space-between;gap:14px;margin-top:8px;">
            <table style="width:100%;">
                <tr><td><b>Name:</b> {student_row.get('StudentName','')}</td><td><b>ID:</b> {student_row.get('StudentID','')}</td></tr>
                <tr><td><b>Class:</b> {student_row.get('Class','')}</td><td><b>Section:</b> {student_row.get('Section','')}</td></tr>
                <tr><td><b>Roll:</b> {student_row.get('Roll','')}</td><td><b>Session:</b> {student_row.get('Session','')}</td></tr>
            </table>
            {photo_html}
        </div>
        <table class="ms-table" style="margin-top:10px;"><tr><th>Date</th><th>Day</th><th>Time</th><th>Subject</th><th>Room</th></tr>{sched_rows}</table>
        <div style="display:flex;justify-content:space-between;margin-top:30px;font-size:0.85rem;">
            <div>_____________________<br/>Student Signature</div>
            <div>_____________________<br/>Class Teacher</div>
            <div>_____________________<br/>Headteacher</div>
        </div>
    </div>
    """
    return html


def render_admit_cards_2up(school_row, cards: list):
    """Lay out admit cards two-per-A4-page for printing (each card already
    contains its routine). 'cards' is a list of pre-rendered card HTML strings."""
    pages = ""
    for i in range(0, len(cards), 2):
        pair = cards[i:i + 2]
        cells = "".join(f'<div style="flex:1;border:1px dashed #94a3b8;padding:10px;">{c}</div>' for c in pair)
        pages += (
            '<div style="display:flex;gap:10px;width:100%;page-break-after:always;'
            f'min-height:48vh;">{cells}</div>'
        )
    return pages


def render_two_copies_stacked(single_html: str) -> str:
    """Generic 'একই পাতায় ২ কপি' helper: stacks the same printed document
    twice on one A4 sheet with a dashed cut-line between them — used for the
    exam routine (as required) and reusable for any other single-page doc."""
    return f"""
    <div style="border-bottom:2px dashed #94a3b8;padding-bottom:16px;margin-bottom:16px;">{single_html}</div>
    <div>{single_html}</div>
    """


def render_routine_grid(school_row, title: str, subtitle: str, routine_df: pd.DataFrame) -> str:
    """Board-style exam routine grid — Date/Day rows down the side, Time-Shift
    groups across the top, Class-Group sub-columns under each shift, and the
    subject/label in each cell — matching the reference routine design."""
    if routine_df.empty:
        return f"""<div class="marksheet">{doc_header(school_row, title, subtitle)}
                    <p style='text-align:center;color:#64748b;'>এখনো কোনো রুটিন এন্ট্রি নেই।</p></div>"""

    shifts = (
        routine_df[["StartTime", "EndTime"]].drop_duplicates()
        .sort_values("StartTime").itertuples(index=False)
    )
    shift_list = list(shifts)
    class_groups = list(dict.fromkeys([c for c in routine_df["ClassGroup"].tolist() if str(c).strip()]))
    if not class_groups:
        class_groups = ["সকল ক্লাস"]
    dates = (
        routine_df[["ExamDate", "DayName"]].drop_duplicates()
        .sort_values("ExamDate").itertuples(index=False)
    )

    header_top = "<tr><th rowspan='2'>তারিখ</th><th rowspan='2'>বার</th>"
    for st_, et_ in shift_list:
        header_top += f"<th colspan='{len(class_groups)}'>সময়: {st_} - {et_}</th>"
    header_top += "</tr>"
    header_bottom = "<tr>" + "".join(f"<th>{cg}</th>" for _ in shift_list for cg in class_groups) + "</tr>"

    body = ""
    for date_, day_ in dates:
        body += f"<tr><td><b>{date_}</b></td><td>{day_}</td>"
        for st_, et_ in shift_list:
            for cg in class_groups:
                match = routine_df[
                    (routine_df["ExamDate"] == date_) & (routine_df["StartTime"] == st_)
                    & (routine_df["EndTime"] == et_) & (routine_df["ClassGroup"] == cg)
                ]
                label = match.iloc[0].get("SubjectID", "") if not match.empty else ""
                body += f"<td style='text-align:center;'>{label or '&nbsp;'}</td>"
        body += "</tr>"

    table = f"<table class='ms-table' style='margin-top:12px;'>{header_top}{header_bottom}{body}</table>"
    return f"""
    <div class="marksheet" style="max-width:100%;">
        {doc_header(school_row, title, subtitle)}
        {table}
    </div>
    """


def render_certificate(school_row, student_row, cert_type: str, extra: dict = None):
    """Auto-generates one of the three certificate types:
      - 'testimonial'  -> প্রশংসাপত্র (character/conduct testimonial)
      - 'transfer'     -> ছাড়পত্র (Transfer / School-Leaving Certificate)
      - 'certification'-> প্রত্যয়নপত্র (bonafide/attendance certification)
    """
    extra = extra or {}
    titles = {
        "testimonial": "TESTIMONIAL (প্রশংসাপত্র)",
        "transfer": "TRANSFER CERTIFICATE (ছাড়পত্র)",
        "certification": "CERTIFICATION (প্রত্যয়নপত্র)",
    }
    title = titles.get(cert_type, "Certificate")
    today = datetime.now().strftime("%d-%m-%Y")

    if cert_type == "transfer":
        body = f"""
        <p>This is to certify that <b>{student_row.get('StudentName','')}</b>, son/daughter of
        <b>{student_row.get('FatherName','')}</b> and <b>{student_row.get('MotherName','')}</b>,
        Student ID <b>{student_row.get('StudentID','')}</b>, was a bona fide student of this
        institution in Class <b>{student_row.get('Class','')}</b>, Section <b>{student_row.get('Section','')}</b>,
        Roll <b>{student_row.get('Roll','')}</b>, Session <b>{student_row.get('Session','')}</b>.</p>
        <p>Reason for leaving: <b>{student_row.get('TCReason','') or extra.get('reason','N/A')}</b>.</p>
        <p>Character during his/her stay in this institution was
        <b>{student_row.get('CharacterStatus','') or extra.get('character','Good')}</b>, and all dues have been
        <b>{student_row.get('DuesStatus','') or extra.get('dues','cleared')}</b>.</p>
        <p>He/She is hereby permitted to be admitted to another institution.</p>
        """
    elif cert_type == "certification":
        body = f"""
        <p>This is to certify that <b>{student_row.get('StudentName','')}</b>,
        Student ID <b>{student_row.get('StudentID','')}</b>, is/was a bona fide student of this
        institution, studying in Class <b>{student_row.get('Class','')}</b>, Section
        <b>{student_row.get('Section','')}</b>, Roll <b>{student_row.get('Roll','')}</b>,
        Session <b>{student_row.get('Session','')}</b>.</p>
        <p>{extra.get('purpose_line', 'This certificate is issued on the student/guardian\'s request for necessary purposes.')}</p>
        """
    else:  # testimonial
        body = f"""
        <p>This is to certify that <b>{student_row.get('StudentName','')}</b>, son/daughter of
        <b>{student_row.get('FatherName','')}</b>, Student ID <b>{student_row.get('StudentID','')}</b>,
        was a student of Class <b>{student_row.get('Class','')}</b>, Section
        <b>{student_row.get('Section','')}</b>, Roll <b>{student_row.get('Roll','')}</b> of this
        institution during the session <b>{student_row.get('Session','')}</b>.</p>
        <p>During this period his/her conduct and character were found to be
        <b>{extra.get('character', 'good and satisfactory')}</b>. We wish him/her every success in life.</p>
        """

    html = f"""
    <div class="marksheet" style="max-width:760px;margin:auto;min-height:480px;">
        {doc_header(school_row, title)}
        <div style="font-size:1rem;line-height:1.9;margin-top:18px;text-align:justify;">
            {body}
        </div>
        <div style="display:flex;justify-content:space-between;margin-top:70px;">
            <div>Date: {extra.get('issue_date', today)}</div>
            <div>_____________________<br/>Headteacher &amp; Signature/Seal</div>
        </div>
    </div>
    """
    return html


def render_guard_list(school_row, exam_date: str, class_scope: str, duty_rows: pd.DataFrame, teacher_lookup: dict):
    """Matches the reference 'DAILY INVIGILATION DUTY & ANSWER SCRIPT
    DISTRIBUTION CHART' design: grouped by Shift, with SL / Room / Assigned
    Classes & Students / Total Scripts Needed / Invigilator Name / Signature."""
    shifts = list(dict.fromkeys(duty_rows["Shift"].tolist())) if not duty_rows.empty else []
    shift_blocks = ""
    for shift_label in shifts:
        rows_in_shift = duty_rows[duty_rows["Shift"] == shift_label].reset_index(drop=True)
        body = ""
        for i, d in rows_in_shift.iterrows():
            names = []
            for c in ["Invigilator1_ID", "Invigilator2_ID", "Invigilator3_ID"]:
                tid = d.get(c, "")
                if tid:
                    names.append(teacher_lookup.get(tid, tid))
            inv_html = "<br>".join(f"{n_i+1}. {nm}" for n_i, nm in enumerate(names)) or "—"
            classes_html = str(d.get("AssignedClasses", "")).replace("\n", "<br>")
            body += (
                f"<tr><td style='text-align:center;'>{i+1}</td>"
                f"<td>{d.get('RoomNo','')}</td>"
                f"<td>{classes_html}</td>"
                f"<td style='text-align:center;'>{d.get('TotalScriptsNeeded','')}</td>"
                f"<td>{inv_html}</td>"
                f"<td style='min-width:90px;'>&nbsp;</td></tr>"
            )
        shift_blocks += f"""
        <div style="background:#1e3a8a;color:white;font-weight:800;padding:7px 12px;
                    margin-top:16px;border-radius:4px 4px 0 0;">{shift_label}</div>
        <table class="ms-table" style="margin-top:0;">
            <tr><th>SL</th><th>Room No</th><th>Assigned Classes &amp; Students</th>
                <th>Total Scripts Needed</th><th>Invigilator Name</th><th>Signature</th></tr>
            {body}
        </table>
        """
    if not shifts:
        shift_blocks = "<p style='text-align:center;color:#64748b;'>No duties assigned yet.</p>"

    html = f"""
    <div class="marksheet" style="max-width:100%;">
        {doc_header(school_row, "DAILY INVIGILATION DUTY &amp; ANSWER SCRIPT DISTRIBUTION CHART",
                    f"Exam Date: {exam_date} &nbsp;|&nbsp; {class_scope}")}
        {shift_blocks}
        <p style="font-size:0.75rem;color:#64748b;margin-top:10px;">
            *Total Scripts Needed = Total Students in that room. Teachers must collect exact script count before entering.
        </p>
        <div style="display:flex;justify-content:flex-end;margin-top:40px;">
            <div>_____________________<br/>Headteacher Signature</div>
        </div>
    </div>
    """
    return html


def render_seat_plan(school_row, exam_name, room_no, plan_df: pd.DataFrame):
    body = ""
    for _, s in plan_df.iterrows():
        body += (
            f"<tr><td>{s.get('BenchNo','')}</td><td>{s.get('SeatPosition','')}</td>"
            f"<td>{s.get('StudentID','')}</td><td>{s.get('Roll','')}</td>"
            f"<td>{s.get('Class','')}-{s.get('Section','')}</td><td>&nbsp;</td></tr>"
        )
    html = f"""
    <div class="marksheet">
        {doc_header(school_row, f"Seat Plan — {exam_name}", f"Room: {room_no}")}
        <table class="ms-table">
            <tr><th>Bench</th><th>Position</th><th>Student ID</th><th>Roll</th><th>Class-Section</th><th>Signature</th></tr>
            {body}
        </table>
    </div>
    """
    return html


def render_script_sheet(school_row, exam_name, dist_rows: pd.DataFrame, subject_lookup: dict):
    body = ""
    for _, d in dist_rows.iterrows():
        body += (
            f"<tr><td>{subject_lookup.get(d.get('SubjectID',''), d.get('SubjectID',''))}</td>"
            f"<td>{d.get('Class','')}-{d.get('Section','')}</td><td>{d.get('TeacherID','')}</td>"
            f"<td>{d.get('TotalScriptsHandedOver','')}</td><td>{d.get('HandoverDate','')}</td>"
            f"<td>{d.get('ReturnedScriptsCount','')}</td><td>{d.get('ReturnDate','')}</td>"
            f"<td>{d.get('ReturnStatus','')}</td></tr>"
        )
    html = f"""
    <div class="marksheet">
        {doc_header(school_row, f"Answer-Script Handover / Return Register — {exam_name}")}
        <table class="ms-table">
            <tr><th>Subject</th><th>Class-Section</th><th>Teacher</th><th>Handed Over</th>
            <th>Handover Date</th><th>Returned</th><th>Return Date</th><th>Status</th></tr>
            {body}
        </table>
    </div>
    """
    return html


def print_button(label="🖨️ Print"):
    st.markdown(
        f"""<button onclick="window.print()" style="background:#2563eb;color:white;border:none;
        padding:8px 18px;border-radius:8px;font-weight:600;cursor:pointer;">{label}</button>""",
        unsafe_allow_html=True,
    )


# =============================================================================
# PAGES
# =============================================================================
def page_dashboard():
    role, school_id = st.session_state["role"], st.session_state["school_id"]
    hero(f"Welcome, {st.session_state['user_name']}", f"{st.session_state['school_name']} · Role: {role}")

    if role == ROLE_SUPERADMIN:
        schools = read_df("Schools")
        subs = read_df("Subscriptions")
        pays = read_df("PaymentLogs")
        c1, c2, c3, c4 = st.columns(4)
        stat_card(c1, "Total Schools", len(schools), nav_key="superadmin")
        stat_card(c2, "Active Schools", (schools.get("IsActive", pd.Series(dtype=str)) == "Yes").sum() if not schools.empty else 0, nav_key="superadmin")
        stat_card(c3, "Active Subscriptions", (subs.get("Status", pd.Series(dtype=str)) == "Active").sum() if not subs.empty else 0, nav_key="superadmin")
        total_paid = num(pays["Amount"]).sum() if not pays.empty and "Amount" in pays.columns else 0
        stat_card(c4, "Total Collected (৳)", f"{total_paid:,.0f}", nav_key="superadmin")
        st.caption("👆 উপরের যেকোনো কার্ডে ট্যাপ করুন — সরাসরি সেই বিভাগের বিস্তারিত তথ্যে চলে যাবেন।")
        with st.expander("🔍 বিস্তারিত দেখুন (Quick Detail Preview)"):
            st.markdown("**Schools**")
            st.dataframe(schools, use_container_width=True, hide_index=True)
            st.markdown("**Subscriptions**")
            st.dataframe(subs if not subs.empty else pd.DataFrame(), use_container_width=True, hide_index=True)
            st.markdown("**Recent Payments**")
            st.dataframe(pays.tail(10) if not pays.empty else pd.DataFrame(), use_container_width=True, hide_index=True)
        st.markdown("### Recent Schools")
        st.dataframe(schools, use_container_width=True, hide_index=True)
        admob_placeholder("dashboard-superadmin")
        return

    students = read_df("Students", school_id=school_id)
    teachers = read_df("Teachers", school_id=school_id)
    results = read_df("Results", school_id=school_id)
    c1, c2, c3, c4 = st.columns(4)
    can_manage = role in ADMIN_LIKE_ROLES or role == ROLE_TEACHER
    stat_card(c1, "Total Students", len(students) if not students.empty else 0, nav_key="students" if can_manage else None)
    stat_card(c2, "Total Teachers", len(teachers) if not teachers.empty else 0,
              nav_key="teachers" if role in ADMIN_LIKE_ROLES else None)
    stat_card(c3, "Results Generated", len(results) if not results.empty else 0,
              nav_key="results" if role in ADMIN_LIKE_ROLES else "print")
    published = (results.get("Published", pd.Series(dtype=str)) == "Yes").sum() if not results.empty else 0
    stat_card(c4, "Published Results", published, nav_key="print")
    st.caption("👆 উপরের যেকোনো কার্ডে ট্যাপ করুন — সরাসরি সেই বিভাগের বিস্তারিত তালিকায় চলে যাবেন।")

    with st.expander("🔍 বিস্তারিত দেখুন (Quick Detail Preview)"):
        dcol1, dcol2 = st.columns(2)
        with dcol1:
            st.markdown("**সাম্প্রতিক ছাত্র-ছাত্রী (Recent Students)**")
            if not students.empty:
                cols = [c for c in ["StudentID", "StudentName", "Class", "Section", "Roll"] if c in students.columns]
                st.dataframe(students[cols].tail(10), use_container_width=True, hide_index=True)
            else:
                st.caption("কোনো ছাত্র-ছাত্রী যোগ করা হয়নি।")
        with dcol2:
            st.markdown("**টিচার তালিকা (Teachers)**")
            if not teachers.empty:
                cols = [c for c in ["TeacherID", "Name", "Role", "Subject"] if c in teachers.columns]
                st.dataframe(teachers[cols], use_container_width=True, hide_index=True)
            else:
                st.caption("কোনো টিচার যোগ করা হয়নি।")
        st.markdown("**সাম্প্রতিক রেজাল্ট (Recent Results)**")
        if not results.empty:
            cols = [c for c in ["StudentID", "Class", "Section", "Percentage", "FinalGrade", "Status", "Published"] if c in results.columns]
            st.dataframe(results[cols].tail(10), use_container_width=True, hide_index=True)
        else:
            st.caption("এখনো কোনো রেজাল্ট তৈরি হয়নি।")

    st.markdown("### 📢 Latest Notices")
    notices = read_df("Notices", school_id=school_id)
    if not notices.empty:
        for _, n in notices.sort_values("Date", ascending=False).head(5).iterrows():
            st.markdown(
                f"""<div class="em-card"><b>{n.get('Title','')}</b>
                <span style="color:#64748b;float:right;">{n.get('Date','')}</span>
                <p style="margin-top:6px;">{n.get('Description','')}</p></div>""",
                unsafe_allow_html=True,
            )
    else:
        st.info("No notices yet.")
    admob_placeholder("dashboard")


def _class_section_options(school_id):
    students = read_df("Students", school_id=school_id)
    classes = sorted(students["Class"].dropna().unique().tolist()) if not students.empty else []
    return students, classes


def page_marks_entry():
    role, school_id, actor = st.session_state["role"], st.session_state["school_id"], st.session_state["user_name"]
    hero("📝 Marks Entry", "Enter marks — grade, GPA and pass/fail are calculated live.")
    hint("প্রথমে Exam, Class, Section ও Subject বেছে নিন — নিচে শুধু সেই Subject-এর যোগ্য "
         "ছাত্র-ছাত্রীদের (তাদের Group ও Religion অনুযায়ী) তালিকা আসবে। নম্বর দেওয়ার পর 'Save Marks' চাপুন — "
         "Grade/GPA/Pass-Fail নিজে থেকেই হিসাব হয়ে যাবে।")

    exams = read_df("Exams", school_id=school_id)
    subjects = read_df("Subjects", school_id=school_id)
    students, classes = _class_section_options(school_id)
    if exams.empty or students.empty or subjects.empty:
        st.warning("Please make sure Exams, Students and Subjects sheets have data for this school.")
        return

    c1, c2, c3 = st.columns(3)
    exam_label = c1.selectbox("Exam", (exams["ExamName"] + " (" + exams["ExamID"] + ")").tolist())
    exam_id = exam_label.split("(")[-1].rstrip(")")
    exam_row = exams[exams["ExamID"] == exam_id].iloc[0]
    klass = c2.selectbox("Class", classes)
    sections = sorted(students[students["Class"] == klass]["Section"].dropna().unique().tolist())
    section = c3.selectbox("Section", sections if sections else ["A"])

    subj_scope = subjects[subjects["Class"] == klass]
    if subj_scope.empty:
        st.warning("No subjects configured for this class.")
        return
    subj_label = st.selectbox("Subject", (subj_scope["SubjectName"] + " (" + subj_scope["SubjectID"] + ")").tolist())
    subject_id = subj_label.split("(")[-1].rstrip(")")
    subject_row = subj_scope[subj_scope["SubjectID"] == subject_id].iloc[0]

    show_written = num(subject_row.get("WrittenFull")) > 0 or subject_row.get("PassPolicy") == "Combined"
    show_mcq = num(subject_row.get("MCQFull")) > 0
    show_practical = num(subject_row.get("PracticalFull")) > 0

    cls_students = students[(students["Class"] == klass) & (students["Section"] == section)
                             & (students.get("Status", "Active") == "Active")].copy()
    # Group + Religion aware filtering: a student only appears for a subject
    # that actually applies to them (their own Group / their own Religion).
    if not cls_students.empty:
        eligible_mask = cls_students.apply(lambda r: student_eligible_for_subject(r, subject_row), axis=1)
        cls_students = cls_students[eligible_mask]
    if cls_students.empty:
        st.info("No active students eligible for this subject in this Class / Section "
                "(check the subject's Group / Religion settings).")
        return

    existing_marks = read_df("Marks", school_id=school_id)
    existing_marks = existing_marks[
        (existing_marks["ExamID"] == exam_id) & (existing_marks["SubjectID"] == subject_id)
    ] if not existing_marks.empty else existing_marks

    grid = cls_students[["StudentID", "Roll", "StudentName"]].sort_values(
        "Roll", key=lambda s: pd.to_numeric(s, errors="coerce")
    ).reset_index(drop=True)
    grid["Present"] = True
    grid["Written"] = 0
    grid["MCQ"] = 0
    grid["Practical"] = 0
    if existing_marks is not None and not existing_marks.empty:
        em = existing_marks.set_index("StudentID")
        for i, r in grid.iterrows():
            if r["StudentID"] in em.index:
                e = em.loc[r["StudentID"]]
                grid.at[i, "Present"] = str(e.get("IsPresent", "Yes")) != "No"
                grid.at[i, "Written"] = num(e.get("Written"))
                grid.at[i, "MCQ"] = num(e.get("MCQ"))
                grid.at[i, "Practical"] = num(e.get("Practical"))

    st.caption(
        f"Full Marks — Written: {subject_row.get('WrittenFull','-')}"
        f"{' | MCQ: ' + str(subject_row.get('MCQFull')) if show_mcq else ''}"
        f"{' | Practical: ' + str(subject_row.get('PracticalFull')) if show_practical else ''}"
        f" | Policy: {subject_row.get('PassPolicy')}"
    )

    cols_to_show = ["StudentID", "Roll", "StudentName", "Present", "Written"]
    if show_mcq:
        cols_to_show.append("MCQ")
    if show_practical:
        cols_to_show.append("Practical")

    edited = st.data_editor(
        grid[cols_to_show], use_container_width=True, hide_index=True, key="marks_grid",
        column_config={
            "StudentID": st.column_config.TextColumn(disabled=True),
            "Roll": st.column_config.TextColumn(disabled=True),
            "StudentName": st.column_config.TextColumn(disabled=True),
            "Present": st.column_config.CheckboxColumn(),
        },
    )

    if st.button("💾 Save & Auto-Calculate", type="primary"):
        gs_df = get_gradescale(school_id)
        new_rows = []
        for _, r in edited.iterrows():
            calc = compute_subject_result(
                subject_row, r.get("Written", 0), r.get("MCQ", 0) if show_mcq else 0,
                r.get("Practical", 0) if show_practical else 0, bool(r["Present"]), gs_df,
            )
            mark_id = None
            if existing_marks is not None and not existing_marks.empty:
                m = existing_marks[existing_marks["StudentID"] == r["StudentID"]]
                if not m.empty:
                    mark_id = m.iloc[0]["MarkID"]
            if not mark_id:
                mark_id = next_seq_id(school_id, "Marks", "MarkID", "M")
            new_rows.append({
                "SchoolID": school_id, "MarkID": mark_id, "StudentID": r["StudentID"],
                "ExamID": exam_id, "Class": klass, "Section": section,
                "Session": exam_row.get("Session", ""), "SubjectID": subject_id,
                "Written": r.get("Written", 0), "MCQ": r.get("MCQ", 0) if show_mcq else 0,
                "Practical": r.get("Practical", 0) if show_practical else 0,
                "EnteredBy": actor, "EnteredDate": datetime.now().strftime("%Y-%m-%d %H:%M"),
                **calc,
            })
        new_df = pd.DataFrame(new_rows)
        upsert_rows("Marks", new_df, key_cols=["StudentID", "ExamID", "SubjectID"])
        st.success(f"Saved & auto-calculated marks for {len(new_df)} students.")
        st.dataframe(
            new_df[["StudentID", "ObtainedMarks", "Percentage", "SubjectGPA", "Grade", "IsPass"]],
            use_container_width=True, hide_index=True,
        )


def page_generate_results():
    school_id, actor = st.session_state["school_id"], st.session_state["user_name"]
    hero("📊 Generate & Publish Results", "Compile subject marks into final results with class merit ranking.")
    hint("Exam, Class, Section বেছে 'Generate / Refresh Results' চাপুন — সব সাবজেক্টের নম্বর যোগ হয়ে "
         "GPA/Grade/মেধাক্রম হিসাব হবে। ফলাফল ঠিক আছে দেখে নিয়ে 'Publish' চাপলেই ছাত্র/অভিভাবক দেখতে পাবে।")

    exams = read_df("Exams", school_id=school_id)
    students, classes = _class_section_options(school_id)
    if exams.empty or not classes:
        st.warning("Need Exams and Students data first.")
        return

    c1, c2, c3 = st.columns(3)
    exam_label = c1.selectbox("Exam", (exams["ExamName"] + " (" + exams["ExamID"] + ")").tolist())
    exam_id = exam_label.split("(")[-1].rstrip(")")
    exam_row = exams[exams["ExamID"] == exam_id].iloc[0]
    klass = c2.selectbox("Class", classes)
    sections = sorted(students[students["Class"] == klass]["Section"].dropna().unique().tolist())
    section = c3.selectbox("Section", sections if sections else ["A"])

    colA, colB = st.columns(2)
    if colA.button("⚙️ Generate / Refresh Results", type="primary"):
        df, msg = generate_results_for_group(school_id, exam_id, klass, section, exam_row.get("Session", ""), actor)
        st.session_state["_last_results"] = df
        st.info(msg)

    df = st.session_state.get("_last_results")
    results_all = read_df("Results", school_id=school_id)
    view = results_all[(results_all["ExamID"] == exam_id) & (results_all["Class"] == klass) & (results_all["Section"] == section)] \
        if not results_all.empty else pd.DataFrame()
    if not view.empty:
        show = view.copy()
        show["Result"] = show["Status"].apply(lambda s: badge("PASS", "pass") if s == "Passed" else badge("FAIL", "fail"))
        show["Published?"] = show["Published"].apply(lambda p: badge("Published", "pass") if p == "Yes" else badge("Draft", "pending"))
        st.markdown("#### Results Preview")
        st.write(
            show[["StudentID", "TotalMarks", "TotalFullMarks", "Percentage", "FinalGPA", "FinalGrade",
                  "ClassMeritPosition", "Result", "Published?"]].to_html(escape=False, index=False),
            unsafe_allow_html=True,
        )
        if colB.button("🔓 Publish These Results"):
            n = publish_group("Results", school_id, {"ExamID": exam_id, "Class": klass, "Section": section}, True, actor)
            st.success(f"Published {n} results. Students/Guardians can now view them.")
            st.rerun()
        if st.button("🔒 Unpublish"):
            publish_group("Results", school_id, {"ExamID": exam_id, "Class": klass, "Section": section}, False, actor)
            st.rerun()
    else:
        st.info("No results generated yet for this selection — click Generate above.")


def page_consolidated():
    school_id, actor = st.session_state["school_id"], st.session_state["user_name"]
    hero("🏆 Consolidated Annual Results & Promotion", "Weighted across all exams · auto-generates next class roll numbers.")
    hint("বার্ষিক পরীক্ষার পর এই পেজ থেকে সব সেমিস্টারের ফলাফল একসাথে করে চূড়ান্ত GPA/Grade বের করুন এবং "
         "'Apply Promotion' চাপলে পাশ করা ছাত্র-ছাত্রীরা পরের ক্লাসে নতুন (শাখাসহ) রোল নিয়ে উঠে যাবে।")

    students, classes = _class_section_options(school_id)
    exams = read_df("Exams", school_id=school_id)
    if not classes or exams.empty:
        st.warning("Need Students and Exams data first.")
        return
    sessions = sorted(exams["Session"].dropna().unique().tolist())

    c1, c2, c3 = st.columns(3)
    session_year = c1.selectbox("Session", sessions)
    klass = c2.selectbox("Class", classes)
    sections = sorted(students[students["Class"] == klass]["Section"].dropna().unique().tolist())
    section = c3.selectbox("Section", sections if sections else ["A"])

    weights = exams[exams["Session"] == session_year][["ExamName", "Weight", "OrderSequence"]].sort_values("OrderSequence")
    st.caption("Exam weights used: " + ", ".join(f"{r.ExamName} = {r.Weight}%" for r in weights.itertuples()))

    if st.button("⚙️ Generate Consolidated Results", type="primary"):
        df, msg = generate_consolidated(school_id, session_year, klass, section, actor)
        st.session_state["_last_consolidated"] = df
        st.info(msg)

    cons_all = read_df("ConsolidatedResults", school_id=school_id)
    view = cons_all[(cons_all["Session"] == session_year) & (cons_all["CurrentClass"] == klass) & (cons_all["Section"] == section)] \
        if not cons_all.empty else pd.DataFrame()
    if not view.empty:
        show = view.copy()
        show["Result"] = show["Status"].apply(lambda s: badge("PASS", "pass") if s == "Passed" else badge("FAIL", "fail"))
        st.markdown("#### Consolidated Preview")
        st.write(
            show[["StudentID", "ConsolidatedTotal", "ConsolidatedPercentage", "ConsolidatedGPA",
                  "ConsolidatedGrade", "Result", "PromotedClass", "NextRoll"]].to_html(escape=False, index=False),
            unsafe_allow_html=True,
        )
        b1, b2, b3 = st.columns(3)
        if b1.button("🔓 Publish Consolidated Results"):
            n = publish_group("ConsolidatedResults", school_id,
                               {"Session": session_year, "CurrentClass": klass, "Section": section}, True, actor)
            st.success(f"Published {n} consolidated results.")
            st.rerun()
        if b2.button("🔒 Unpublish"):
            publish_group("ConsolidatedResults", school_id,
                           {"Session": session_year, "CurrentClass": klass, "Section": section}, False, actor)
            st.rerun()
        if b3.button("🚀 Apply Promotion to Students Sheet"):
            n = apply_promotion(school_id, session_year, klass, section)
            st.success(f"Promoted {n} students — Class & Roll updated on the Students sheet for next year.")
    else:
        st.info("No consolidated results yet — click Generate above (make sure each exam's Results were generated first).")


def page_print_center():
    school_id = st.session_state["school_id"]
    role = st.session_state["role"]
    hero("🖨️ Print Center", "Marksheets & Admit Cards — print-ready, publish-locked.")
    hint("Marksheet ট্যাব থেকে Exam/Class/Section/Student বেছে সুন্দর ডিজাইনের মার্কশীট বানান; "
         "Admit Card ট্যাব থেকে রুটিনসহ এডমিট কার্ড বানান। যেকোনোটাতে '🖨️ Print / Save as PDF' চাপলেই ব্রাউজারের প্রিন্ট ডায়ালগ খুলবে — সেখান থেকে PDF সেভ বা প্রিন্ট করতে পারবেন।")

    schools = read_df("Schools")
    school_row = schools[schools["SchoolID"] == school_id].iloc[0] if not schools.empty else pd.Series()

    tab1, tab2 = st.tabs(["📄 Marksheet", "🎫 Admit Card"])

    with tab1:
        exams = read_df("Exams", school_id=school_id)
        students, classes = _class_section_options(school_id)
        if exams.empty or not classes:
            st.info("Need Exams and Students data.")
        else:
            c1, c2 = st.columns(2)
            exam_label = c1.selectbox("Exam", (exams["ExamName"] + " (" + exams["ExamID"] + ")").tolist(), key="ms_exam")
            exam_id = exam_label.split("(")[-1].rstrip(")")
            klass = c2.selectbox("Class", classes, key="ms_class")
            cs = students[students["Class"] == klass]
            student_label = st.selectbox(
                "Student", (cs["StudentName"] + " — Roll " + cs["Roll"] + " (" + cs["StudentID"] + ")").tolist()
            ) if not cs.empty else None
            if student_label:
                student_id = student_label.split("(")[-1].rstrip(")")
                student_row = cs[cs["StudentID"] == student_id].iloc[0]
                results = read_df("Results", school_id=school_id)
                match = results[(results["StudentID"] == student_id) & (results["ExamID"] == exam_id)]
                if match.empty:
                    st.warning("Result not generated yet for this student/exam.")
                elif str(match.iloc[0].get("Published")) != "Yes" and role not in (ROLE_ADMIN, "Headmaster", "Headteacher", ROLE_SUPERADMIN):
                    st.warning("🔒 Result is not published yet.")
                else:
                    marks = read_df("Marks", school_id=school_id)
                    marks = marks[(marks["StudentID"] == student_id) & (marks["ExamID"] == exam_id)]
                    subjects = read_df("Subjects", school_id=school_id)
                    exam_name = exams[exams["ExamID"] == exam_id].iloc[0]["ExamName"]
                    html = render_marksheet(school_row, student_row, match.iloc[0], marks, subjects, exam_name)
                    st.markdown(html, unsafe_allow_html=True)
                    print_button()

    with tab2:
        exams2 = read_df("Exams", school_id=school_id)
        students2, classes2 = _class_section_options(school_id)
        if exams2.empty or not classes2:
            st.info("Need Exams and Students data.")
        else:
            hint("একজনের এডমিট কার্ড দেখতে নিচে Student বেছে নিন। পুরো ক্লাসের সবার এডমিট কার্ড একসাথে "
                 "(প্রতি A4 পাতায় ২ জন করে, কাটার জন্য মাঝে ডট-লাইনসহ) ছাপাতে নিচের 'Bulk Print' অপশন ব্যবহার করুন।")
            mode = st.radio("মোড", ["👤 একজন ছাত্রের Admit Card", "👥 পুরো ক্লাসের Admit Card (২ জন/পাতা)"], horizontal=True)
            c1, c2 = st.columns(2)
            exam_label2 = c1.selectbox("Exam", (exams2["ExamName"] + " (" + exams2["ExamID"] + ")").tolist(), key="ac_exam")
            exam_id2 = exam_label2.split("(")[-1].rstrip(")")
            klass2 = c2.selectbox("Class", classes2, key="ac_class")
            cs2 = students2[students2["Class"] == klass2]
            exam_name2 = exams2[exams2["ExamID"] == exam_id2].iloc[0]["ExamName"]
            routines_all = read_df("Routines", school_id=school_id)
            routines_all = routines_all[routines_all["ExamID"] == exam_id2] if not routines_all.empty else routines_all

            if mode.startswith("👤"):
                student_label2 = st.selectbox(
                    "Student", (cs2["StudentName"] + " — Roll " + cs2["Roll"] + " (" + cs2["StudentID"] + ")").tolist(), key="ac_student"
                ) if not cs2.empty else None
                if student_label2:
                    sid2 = student_label2.split("(")[-1].rstrip(")")
                    srow2 = cs2[cs2["StudentID"] == sid2].iloc[0]
                    html2 = render_admit_card(school_row, srow2, exam_name2, routines_all)
                    st.markdown(html2, unsafe_allow_html=True)
                    print_button()
            else:
                sections2 = sorted(cs2["Section"].dropna().unique().tolist())
                section2 = st.selectbox("Section", sections2 if sections2 else ["A"], key="ac_bulk_section")
                bulk_students = cs2[cs2["Section"] == section2].sort_values(
                    "Roll", key=lambda s: pd.to_numeric(s, errors="coerce")
                )
                if bulk_students.empty:
                    st.info("এই ক্লাস/শাখায় কোনো ছাত্র নেই।")
                elif st.button("🖨️ Generate All Admit Cards (২ জন/পাতা)", type="primary"):
                    cards = [
                        render_admit_card(school_row, srow, exam_name2, routines_all)
                        for _, srow in bulk_students.iterrows()
                    ]
                    st.success(f"{len(cards)} জনের Admit Card তৈরি হয়েছে — মোট {(len(cards) + 1)//2} পাতা (২ জন/পাতা)।")
                    st.markdown(render_admit_cards_2up(school_row, cards), unsafe_allow_html=True)
                    print_button()


def page_teachers():
    school_id = st.session_state["school_id"]
    hero("👩‍🏫 Teachers & Staff", "Manage staff accounts. PINs are masked for security.")
    hint("নতুন টিচার/স্টাফ যোগ করতে নিচের ফর্ম পূরণ করুন। Role বেছে নিন — Headteacher/Admin সব দেখবে, "
         "Teacher নিজের ক্লাসের কাজ করবে, Clerk শুধু নোটিশ দেবে, আর '(No Role)' দিলে সে শুধু সীমিত তথ্য দেখবে। "
         "তৈরি হওয়া PIN-টি গোপনে ওই টিচারকে জানিয়ে দিন।")
    teachers = read_df("Teachers", school_id=school_id)

    with st.expander("➕ Add New Teacher / Staff"):
        with st.form("add_teacher"):
            c1, c2, c3 = st.columns(3)
            name = c1.text_input("Name", placeholder="যেমন: Md. Karim Uddin")
            role_choice = c2.selectbox(
                "Role", [ROLE_ADMIN, ROLE_TEACHER, ROLE_CLERK, "Headmaster", "(No Role — limited view-only Staff)"],
                help="Pick '(No Role)' for staff who should only see limited school info with no editing rights."
            )
            role_ = "" if role_choice.startswith("(No Role") else role_choice
            designation = c3.text_input("Designation", placeholder="যেমন: Assistant Teacher")
            c4, c5, c6 = st.columns(3)
            klass = c4.text_input("Class (if class teacher)", placeholder="যেমন: Class-9")
            section = c5.text_input("Section", placeholder="যেমন: A")
            subject = c6.text_input("Subject", placeholder="যেমন: Bangla")
            c7, c8, c9 = st.columns(3)
            phone = c7.text_input("Phone", placeholder="01XXXXXXXXX")
            email = c8.text_input("Email", placeholder="name@example.com")
            pin = c9.text_input("Set PIN", value=str(np.random.randint(1000, 9999)),
                                 help="৪ সংখ্যার একটি PIN অটো বসানো আছে — চাইলে বদলে দিন। এটি লগইনের সময় লাগবে।")
            if st.form_submit_button("Add Teacher", type="primary"):
                tid = next_seq_id(school_id, "Teachers", "TeacherID", "T", width=3, include_year=False)
                new = pd.DataFrame([{
                    "SchoolID": school_id, "TeacherID": tid, "Name": name, "Email": email,
                    "Class": klass, "Section": section, "Role": role_, "Designation": designation,
                    "Subject": subject, "Phone": phone, "IsActive": "Yes",
                    "JoiningDate": datetime.now().strftime("%Y-%m-%d"), "PIN": pin,
                }])
                upsert_rows("Teachers", new, key_cols=["TeacherID"])
                st.success(f"Added {name} — ID: {tid} — PIN: {pin} (share this with them securely).")
                st.rerun()

    if not teachers.empty:
        show = teachers.copy()
        show["PIN"] = show["PIN"].astype(str).apply(lambda p: "••" + p[-2:] if len(p) >= 2 else "••")
        cols = [c for c in ["TeacherID", "Name", "Role", "Designation", "Class", "Section", "Subject", "Phone", "IsActive", "PIN"] if c in show.columns]
        st.dataframe(show[cols], use_container_width=True, hide_index=True)

        with st.expander("🔑 Reset a Teacher's PIN"):
            pick = st.selectbox("Teacher", (teachers["Name"] + " (" + teachers["TeacherID"] + ")").tolist())
            new_pin = st.text_input("New PIN", value=str(np.random.randint(1000, 9999)))
            if st.button("Reset PIN"):
                tid = pick.split("(")[-1].rstrip(")")
                teachers.loc[teachers["TeacherID"] == tid, "PIN"] = new_pin
                upsert_rows("Teachers", teachers[teachers["TeacherID"] == tid], key_cols=["TeacherID"])
                st.success("PIN updated.")

        with st.expander("🚫 Activate / Deactivate"):
            pick2 = st.selectbox("Teacher ", (teachers["Name"] + " (" + teachers["TeacherID"] + ")").tolist(), key="deact")
            new_status = st.radio("Status", ["Yes", "No"], horizontal=True)
            if st.button("Update Status"):
                tid2 = pick2.split("(")[-1].rstrip(")")
                teachers.loc[teachers["TeacherID"] == tid2, "IsActive"] = new_status
                upsert_rows("Teachers", teachers[teachers["TeacherID"] == tid2], key_cols=["TeacherID"])
                st.success("Status updated.")
    else:
        st.info("No teachers yet — add one above.")


def page_students():
    school_id = st.session_state["school_id"]
    hero("🎒 Students", "Manage student records.")
    hint("নতুন ছাত্র-ছাত্রী যোগ করতে ফর্ম পূরণ করুন। Group (Science/Commerce/Arts/Core) ও Religion অনুযায়ী "
         "Marks Entry ও Marksheet-এ শুধু তার নিজের প্রযোজ্য সাবজেক্টই দেখাবে/আসবে — তাই এই দুটো ঠিকভাবে দিন।")
    students = read_df("Students", school_id=school_id)

    with st.expander("➕ Add New Student"):
        with st.form("add_student"):
            c1, c2, c3 = st.columns(3)
            name = c1.text_input("Student Name", placeholder="যেমন: Karim Ahmed")
            klass = c2.text_input("Class", placeholder="যেমন: Class-9")
            section = c3.text_input("Section", placeholder="যেমন: A")
            c4, c5, c6 = st.columns(3)
            roll = c4.text_input("Roll", placeholder="যেমন: 05")
            session_year = c5.text_input("Session", value=datetime.now().strftime("%Y"))
            gender = c6.selectbox("Gender", ["Male", "Female", "Other"])
            c7, c8, c9 = st.columns(3)
            group = c7.selectbox("Group (বিভাগ)", ["Core", "Science", "Commerce", "Arts"],
                                  help="ক্লাস ৯-১০ হলে Science/Commerce/Arts বেছে নিন; না হলে 'Core' রাখুন।")
            religion = c8.selectbox("Religion (ধর্ম)", ["Islam", "Hindu", "Christian", "Buddhist", "Other"],
                                     help="ধর্ম বিষয়ের মার্কশীটে শুধু নিজের ধর্মের সাবজেক্টই দেখানোর জন্য দরকার।")
            father = c9.text_input("Father's Name")
            c10, c11 = st.columns(2)
            mother = c10.text_input("Mother's Name")
            phone = c11.text_input("Guardian Phone", placeholder="01XXXXXXXXX")
            if st.form_submit_button("Add Student", type="primary"):
                sid = next_seq_id(school_id, "Students", "StudentID", "S")
                new = pd.DataFrame([{
                    "SchoolID": school_id, "StudentID": sid, "Roll": roll, "Class": klass,
                    "Section": section, "Group": group, "Religion": religion, "StudentName": name,
                    "FatherName": father, "MotherName": mother, "Phone": phone, "Status": "Active",
                    "Gender": gender, "Session": session_year,
                }])
                upsert_rows("Students", new, key_cols=["StudentID"])
                st.success(f"Added {name} — ID: {sid}")
                st.rerun()

    if not students.empty:
        f1, f2 = st.columns(2)
        classes = ["All"] + sorted(students["Class"].dropna().unique().tolist())
        fclass = f1.selectbox("Filter by Class", classes)
        fstatus = f2.selectbox("Filter by Status", ["All", "Active", "Inactive"])
        view = students.copy()
        if fclass != "All":
            view = view[view["Class"] == fclass]
        if fstatus != "All":
            view = view[view.get("Status", "Active") == fstatus]
        cols = [c for c in ["StudentID", "Roll", "StudentName", "Class", "Section", "FatherName", "Phone", "Status"] if c in view.columns]
        st.dataframe(view[cols], use_container_width=True, hide_index=True)
    else:
        st.info("No students yet — add one above.")


def page_notices():
    school_id, role, actor = st.session_state["school_id"], st.session_state["role"], st.session_state["user_name"]
    hero("📢 Notices", "School-wide announcements.")
    if role in (ROLE_ADMIN, ROLE_CLERK, "Headmaster", "Headteacher"):
        hint("একটি শিরোনাম ও বিস্তারিত লিখে 'Publish Notice' চাপুন — সাথে সাথে স্কুলের সবাই "
             "(টিচার/স্টাফ) তাদের Dashboard ও Notices পেজে এটি দেখতে পাবে।")
        with st.form("add_notice"):
            title = st.text_input("Title", placeholder="যেমন: বার্ষিক পরীক্ষার রুটিন প্রকাশ")
            desc = st.text_area("Description", placeholder="নোটিশের সম্পূর্ণ বিবরণ এখানে লিখুন...")
            if st.form_submit_button("Publish Notice", type="primary"):
                nid = next_seq_id(school_id, "Notices", "NoticeID", "N")
                append_row("Notices", {
                    "SchoolID": school_id, "NoticeID": nid, "Title": title,
                    "Date": datetime.now().strftime("%Y-%m-%d"), "Description": desc,
                })
                st.success("Notice published.")
                st.rerun()
    notices = read_df("Notices", school_id=school_id)
    if not notices.empty:
        for _, n in notices.sort_values("Date", ascending=False).iterrows():
            st.markdown(
                f"""<div class="em-card"><b>{n.get('Title','')}</b>
                <span style="color:#64748b;float:right;">{n.get('Date','')}</span>
                <p style="margin-top:6px;">{n.get('Description','')}</p></div>""",
                unsafe_allow_html=True,
            )
    else:
        st.info("No notices posted yet.")


def page_routines():
    school_id, role = st.session_state["school_id"], st.session_state["role"]
    hero("🗓️ Routines", "Exam & class schedules.")
    exams = read_df("Exams", school_id=school_id)

    if role in (ROLE_ADMIN, "Headmaster", "Headteacher") and not exams.empty:
        with st.expander("➕ Add Routine Entry"):
            hint("প্রতিটি (তারিখ + সময় + ক্লাস-গ্রুপ) কম্বিনেশনের জন্য একটা করে এন্ট্রি যোগ করুন। "
                 "'Class Group' ঘরে কলামের হেডিং লিখুন (যেমন: '৯ম ও ১০ম'), আর 'Subject/বিষয়' ঘরে ঠিক যা "
                 "প্রিন্টে সেই ঘরে দেখতে চান তা লিখুন (যেমন: 'ইংরেজি-১ম')। একই তারিখ+সময়ে সব ক্লাস-গ্রুপের "
                 "এন্ট্রি যোগ করলে ছবির মতো একটা সম্পূর্ণ রুটিন গ্রিড তৈরি হবে।")
            with st.form("add_routine"):
                c1, c2 = st.columns(2)
                exam_label = c1.selectbox("Exam", (exams["ExamName"] + " (" + exams["ExamID"] + ")").tolist())
                exam_id = exam_label.split("(")[-1].rstrip(")")
                shift = c2.text_input("Shift (internal label)", value="SHIFT - 1: MORNING")
                c3, c4, c5 = st.columns(3)
                classes_txt = c3.text_input("Classes (label)", help="স্বাধীন নোট — গার্ড লিস্ট/অন্য জায়গায় ব্যবহার হয়।")
                class_group = c4.text_input(
                    "Class Group (গ্রিডের কলাম হেডিং)", placeholder="যেমন: ৯ম ও ১০ম",
                    help="রুটিন গ্রিড প্রিন্টে এই লেখাটাই কলাম হেডিং হিসেবে যাবে।",
                )
                exam_date = c5.date_input("Exam Date")
                c6, c7, c8 = st.columns(3)
                start_t = c6.time_input("Start Time")
                end_t = c7.time_input("End Time")
                subject_label = c8.text_input(
                    "Subject / বিষয় (এই ঘরে যা লেখা থাকবে)", placeholder="যেমন: ইংরেজি-১ম",
                    help="গ্রিডের এই তারিখ+সময়+ক্লাস-গ্রুপের ঘরে ঠিক এই লেখাটাই দেখাবে।",
                )
                if st.form_submit_button("Add", type="primary"):
                    rid = next_seq_id(school_id, "Routines", "RoutineID", "ROU")
                    append_row("Routines", {
                        "SchoolID": school_id, "RoutineID": rid, "ExamID": exam_id, "Shift": shift,
                        "Classes": classes_txt, "ClassGroup": class_group,
                        "ExamDate": exam_date.strftime("%Y-%m-%d"), "DayName": exam_date.strftime("%A"),
                        "StartTime": start_t.strftime("%H:%M"), "EndTime": end_t.strftime("%H:%M"),
                        "SubjectID": subject_label,
                    })
                    st.success("Routine added.")
                    st.rerun()

    routines = read_df("Routines", school_id=school_id)
    if not routines.empty:
        cols = [c for c in ["ExamID", "Shift", "Classes", "ClassGroup", "ExamDate", "DayName",
                             "StartTime", "EndTime", "SubjectID"] if c in routines.columns]
        st.dataframe(routines[cols], use_container_width=True, hide_index=True)

        st.markdown("#### 🖨️ Print Routine — Board Format (A4, ২ কপি এক পাতায়)")
        exam_label2 = st.selectbox(
            "Exam", (exams["ExamName"] + " (" + exams["ExamID"] + ")").tolist(), key="routine_print_exam"
        ) if not exams.empty else None
        if exam_label2:
            exam_id2 = exam_label2.split("(")[-1].rstrip(")")
            exam_row2 = exams[exams["ExamID"] == exam_id2].iloc[0]
            scoped = routines[routines["ExamID"] == exam_id2]
            title = st.text_input(
                "Print Title", value=f"{exam_row2.get('ExamName','')} পরীক্ষার রুটিন - {exam_row2.get('Session','')} ইং",
            )
            schools = read_df("Schools")
            school_row = schools[schools["SchoolID"] == school_id].iloc[0] if not schools.empty else pd.Series()
            single = render_routine_grid(school_row, title, "", scoped)
            st.markdown(render_two_copies_stacked(single), unsafe_allow_html=True)
            print_button()
    else:
        st.info("No routines added yet.")


def page_superadmin():
    hero("🛡️ SuperAdmin Console", "Schools, subscriptions, payment approval, support chat & role preview.")
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        ["🏫 Schools", "👑 School Admins", "💳 Subscriptions", "🧾 Payment Approval",
         "💬 Support Chat", "🎭 Role Preview"]
    )

    with tab1:
        schools = read_df("Schools")
        st.dataframe(schools, use_container_width=True, hide_index=True)
        with st.expander("➕ Add School"):
            with st.form("add_school"):
                c1, c2 = st.columns(2)
                name = c1.text_input("School Name")
                phone = c2.text_input("Phone")
                address = st.text_input("Address")
                c3, c4 = st.columns(2)
                eiin = c3.text_input("School EIIN")
                exam_system = c4.selectbox("Exam System", ["3-Term", "2-Term", "Semester"])
                logo_url = st.text_input("Logo URL (or paste a base64 string — shown on every printed document)")
                if st.form_submit_button("Create School", type="primary"):
                    n = len(schools) + 1
                    sid = f"SCH-{n:06d}"
                    append_row("Schools", {
                        "SchoolID": sid, "SchoolName": name, "Address": address, "Phone": phone,
                        "SchoolEIIN": eiin, "Logo": logo_url, "IsActive": "Yes", "ExamSystem": exam_system,
                    })
                    st.success(f"School created: {sid}")
                    st.rerun()

        if not schools.empty:
            st.markdown("#### ✏️ Edit / Activate / Deactivate a School")
            with st.form("edit_school"):
                pick = st.selectbox("School", (schools["SchoolName"] + " (" + schools["SchoolID"] + ")").tolist())
                sid = pick.split("(")[-1].rstrip(")")
                srow = schools[schools["SchoolID"] == sid].iloc[0]
                c1, c2 = st.columns(2)
                name = c1.text_input("School Name", value=str(srow.get("SchoolName", "")))
                phone = c2.text_input("Phone", value=str(srow.get("Phone", "")))
                address = st.text_input("Address", value=str(srow.get("Address", "")))
                c3, c4 = st.columns(2)
                eiin = c3.text_input("EIIN", value=str(srow.get("SchoolEIIN", "")))
                status = c4.radio("Status", ["Yes", "No"], index=0 if str(srow.get("IsActive", "Yes")) == "Yes" else 1, horizontal=True)
                logo_url = st.text_input("Logo URL / base64", value=str(srow.get("Logo", "")))
                if st.form_submit_button("💾 Save School", type="primary"):
                    schools.loc[schools["SchoolID"] == sid, ["SchoolName", "Phone", "Address", "SchoolEIIN", "IsActive", "Logo"]] = \
                        [name, phone, address, eiin, status, logo_url]
                    upsert_rows("Schools", schools[schools["SchoolID"] == sid], key_cols=["SchoolID"])
                    st.success("School updated. Deactivated schools can no longer log in.")
                    st.rerun()

    with tab2:
        st.caption("SuperAdmin can view, reset PIN, and activate/deactivate ANY school's Admin (Headteacher) account.")
        all_teachers = read_df("Teachers")
        schools_lu = read_df("Schools").set_index("SchoolID")["SchoolName"].to_dict()
        admins = all_teachers[all_teachers["Role"].isin([ROLE_ADMIN, "Headmaster", "Headteacher"])] if not all_teachers.empty else pd.DataFrame()
        if admins.empty:
            st.info("No school admins found yet.")
        else:
            show = admins.copy()
            show["School"] = show["SchoolID"].map(lambda s: schools_lu.get(s, s))
            show["PIN"] = show["PIN"].astype(str).apply(lambda p: "••" + p[-2:] if len(p) >= 2 else "••")
            cols = [c for c in ["School", "TeacherID", "Name", "Role", "Phone", "Email", "IsActive", "PIN"] if c in show.columns]
            st.dataframe(show[cols], use_container_width=True, hide_index=True)

            with st.expander("🔑 Reset an Admin's PIN"):
                pick = st.selectbox("Admin", (admins["Name"] + " — " + admins["SchoolID"].map(lambda s: schools_lu.get(s, s)) + " (" + admins["TeacherID"] + ")").tolist())
                new_pin = st.text_input("New PIN", value=str(np.random.randint(1000, 9999)), key="sa_reset_pin")
                if st.button("Reset Admin PIN"):
                    tid = pick.split("(")[-1].rstrip(")")
                    all_teachers.loc[all_teachers["TeacherID"] == tid, "PIN"] = new_pin
                    upsert_rows("Teachers", all_teachers[all_teachers["TeacherID"] == tid], key_cols=["TeacherID"])
                    st.success("Admin PIN reset.")

            with st.expander("🚫 Activate / Deactivate an Admin"):
                pick2 = st.selectbox("Admin ", (admins["Name"] + " — " + admins["SchoolID"].map(lambda s: schools_lu.get(s, s)) + " (" + admins["TeacherID"] + ")").tolist(), key="sa_deact")
                new_status = st.radio("Status ", ["Yes", "No"], horizontal=True, key="sa_deact_status")
                if st.button("Update Admin Status"):
                    tid2 = pick2.split("(")[-1].rstrip(")")
                    all_teachers.loc[all_teachers["TeacherID"] == tid2, "IsActive"] = new_status
                    upsert_rows("Teachers", all_teachers[all_teachers["TeacherID"] == tid2], key_cols=["TeacherID"])
                    st.success("Admin status updated.")

    with tab3:
        subs = read_df("Subscriptions")
        st.dataframe(subs, use_container_width=True, hide_index=True)
        with st.expander("➕ Add Subscription"):
            with st.form("add_sub"):
                schools = read_df("Schools")
                sch_label = st.selectbox("School", (schools["SchoolName"] + " (" + schools["SchoolID"] + ")").tolist()) if not schools.empty else None
                plan = st.selectbox("Plan", ["Monthly", "Yearly", "Trial"])
                amount = st.number_input("Amount (৳)", min_value=0)
                start = st.date_input("Start Date")
                expiry = st.date_input("Expiry Date")
                if st.form_submit_button("Add Subscription", type="primary") and sch_label:
                    sid = sch_label.split("(")[-1].rstrip(")")
                    subid = f"{sid.replace('SCH-','')}-{datetime.now().strftime('%y')}-SUBS{len(subs)+1:03d}"
                    append_row("Subscriptions", {
                        "SchoolID": sid, "SubscriptionID": subid, "PlanType": plan, "Amount": amount,
                        "StartDate": start.strftime("%Y-%m-%d"), "ExpiryDate": expiry.strftime("%Y-%m-%d"),
                        "Status": "Active", "LastPaymentDate": start.strftime("%Y-%m-%d"),
                    })
                    st.success("Subscription added.")
                    st.rerun()

    with tab4:
        hint("নিচে যেসব পেমেন্ট 'Pending' অবস্থায় আছে সেগুলো স্কুল নিজেরাই Billing পেজ থেকে জমা দিয়েছে "
             "(bKash/Nagad TrxID সহ)। TrxID যাচাই করে 'Approve' চাপলেই সেই স্কুলের সাবস্ক্রিপশন অটো নবায়ন "
             "হয়ে যাবে এবং লক খুলে যাবে।")
        pays = read_df("PaymentLogs")
        schools_lu4 = read_df("Schools").set_index("SchoolID")["SchoolName"].to_dict()
        pending = pays[pays.get("Status", "") == "Pending"] if not pays.empty else pd.DataFrame()
        if not pending.empty:
            st.markdown("#### ⏳ Pending Approval")
            for _, p in pending.iterrows():
                with st.container(border=True) if hasattr(st, "container") else st.container():
                    c1, c2 = st.columns([3, 1])
                    c1.markdown(
                        f"**{schools_lu4.get(p.get('SchoolID',''), p.get('SchoolID',''))}** — "
                        f"{p.get('Gateway','')} · TrxID: `{p.get('TrxID','')}` · ৳{p.get('Amount','')} · {p.get('PaymentDate','')}"
                    )
                    if c2.button("✅ Approve", key=f"approve_{p.get('PaymentID')}", type="primary"):
                        new_exp = approve_payment(p)
                        st.success(f"অনুমোদিত! নতুন মেয়াদ: {new_exp}")
                        st.rerun()
        else:
            st.caption("এখন কোনো Pending পেমেন্ট নেই।")

        st.markdown("#### 🧾 সম্পূর্ণ পেমেন্ট হিস্ট্রি")
        st.dataframe(pays, use_container_width=True, hide_index=True)
        with st.expander("➕ ম্যানুয়ালি একটা পেমেন্ট সরাসরি লগ করুন (যেমন: হাতে-হাতে ক্যাশ)"):
            with st.form("add_pay"):
                schools = read_df("Schools")
                sch_label = st.selectbox("School", (schools["SchoolName"] + " (" + schools["SchoolID"] + ")").tolist(), key="pay_school") if not schools.empty else None
                gateway = st.selectbox("Gateway", ["bKash", "Nagad", "Rocket", "Bank", "Cash"])
                trx = st.text_input("Transaction ID")
                amount = st.number_input("Amount (৳) ", min_value=0)
                if st.form_submit_button("Log &amp; Approve Payment", type="primary") and sch_label:
                    sid = sch_label.split("(")[-1].rstrip(")")
                    payid = f"{sid.replace('SCH-','')}-{datetime.now().strftime('%y%m%d%H%M%S')}"
                    row = {
                        "SchoolID": sid, "PaymentID": payid, "TrxID": trx, "Gateway": gateway,
                        "Amount": amount, "PaymentDate": datetime.now().strftime("%Y-%m-%d"), "Status": "Pending",
                    }
                    append_row("PaymentLogs", row)
                    new_exp = approve_payment(pd.Series(row))
                    st.success(f"পেমেন্ট লগ ও অনুমোদিত। নতুন মেয়াদ: {new_exp}")
                    st.rerun()

    with tab5:
        hint("এখান থেকে যেকোনো স্কুলের Admin-এর সাথে সরাসরি চ্যাট করতে পারবেন — সাপোর্ট বা নোটিফিকেশনের জন্য।")
        schools5 = read_df("Schools")
        if schools5.empty:
            st.info("কোনো স্কুল নেই।")
        else:
            pick5 = st.selectbox("School", (schools5["SchoolName"] + " (" + schools5["SchoolID"] + ")").tolist(), key="chat_school_pick")
            sid5 = pick5.split("(")[-1].rstrip(")")
            sname5 = pick5.split("  (")[0]
            render_chat_thread(sid5, sname5, sender_default="SuperAdmin", sender_role_default=ROLE_SUPERADMIN)

    with tab6:
        hint("এটা শুধু টেস্টিং/ডেমোর জন্য — একটা স্কুল ও রোল বেছে সেই ভূমিকায় অ্যাপটা কেমন দেখাবে তা "
             "প্রিভিউ করতে পারবেন। আসল লগইন পরিবর্তন হয় না — 'Exit Preview' চাপলেই আবার SuperAdmin-এ ফিরে আসবেন।")
        schools6 = read_df("Schools")
        if schools6.empty:
            st.info("কোনো স্কুল নেই।")
        else:
            c1, c2 = st.columns(2)
            pick6 = c1.selectbox("Preview School", (schools6["SchoolName"] + " (" + schools6["SchoolID"] + ")").tolist(), key="preview_school_pick")
            preview_role = c2.selectbox("Preview Role", [ROLE_ADMIN, ROLE_TEACHER, ROLE_CLERK, ROLE_STAFF],
                                         format_func=lambda r: r if r else "(No Role — Staff)")
            if st.button("🎭 Start Preview", type="primary"):
                sid6 = pick6.split("(")[-1].rstrip(")")
                sname6 = pick6.split("  (")[0]
                st.session_state["real_role"] = st.session_state["role"]
                st.session_state["real_school_id"] = st.session_state["school_id"]
                st.session_state["real_school_name"] = st.session_state["school_name"]
                st.session_state["real_user_name"] = st.session_state["user_name"]
                st.session_state["impersonating"] = True
                st.session_state["role"] = preview_role
                st.session_state["school_id"] = sid6
                st.session_state["school_name"] = sname6
                st.session_state["user_name"] = f"(Preview) {preview_role or 'Staff'}"
                st.session_state["nav"] = "dashboard"
                st.rerun()


# =============================================================================
# MY PROFILE — self-service for Teacher/Admin/Clerk/Staff
#   (নিজের PIN পরিবর্তন করা, নিজের প্রোফাইল edit করা)
# =============================================================================
def page_my_profile():
    school_id = st.session_state["school_id"]
    teacher_id = st.session_state["teacher_id"]
    hero("👤 My Profile", "Edit your own details and change your login PIN.")

    teachers = read_df("Teachers", school_id=school_id)
    mine = teachers[teachers["TeacherID"] == teacher_id]
    if mine.empty:
        st.info("Profile record not found.")
        return
    me = mine.iloc[0]

    st.markdown("#### ✏️ Edit My Profile")
    with st.form("edit_my_profile"):
        c1, c2 = st.columns(2)
        phone = c1.text_input("Phone", value=str(me.get("Phone", "")))
        email = c2.text_input("Email", value=str(me.get("Email", "")))
        address = st.text_input("Address", value=str(me.get("Address", "")))
        if st.form_submit_button("💾 Save Profile", type="primary"):
            teachers.loc[teachers["TeacherID"] == teacher_id, "Phone"] = phone
            teachers.loc[teachers["TeacherID"] == teacher_id, "Email"] = email
            teachers.loc[teachers["TeacherID"] == teacher_id, "Address"] = address
            upsert_rows("Teachers", teachers[teachers["TeacherID"] == teacher_id], key_cols=["TeacherID"])
            st.success("Profile updated.")
            st.rerun()

    st.markdown("#### 🔑 Change My PIN")
    with st.form("change_my_pin"):
        old_pin = st.text_input("Current PIN", type="password")
        new_pin = st.text_input("New PIN", type="password")
        confirm_pin = st.text_input("Confirm New PIN", type="password")
        if st.form_submit_button("Update PIN", type="primary"):
            if str(me.get("PIN", "")) != old_pin.strip():
                st.error("Current PIN is incorrect.")
            elif not new_pin.strip() or new_pin != confirm_pin:
                st.error("New PIN and Confirm PIN must match and cannot be empty.")
            else:
                teachers.loc[teachers["TeacherID"] == teacher_id, "PIN"] = new_pin.strip()
                upsert_rows("Teachers", teachers[teachers["TeacherID"] == teacher_id], key_cols=["TeacherID"])
                st.success("PIN updated — use it next time you log in.")


# =============================================================================
# EXAM GUARD LIST / INVIGILATION DUTIES
#   এক হলে একাধিক টিচার, একেক হলে একেকদিন, নির্দিষ্ট টিচার নির্দিষ্ট ক্লাসে না
# =============================================================================
# =============================================================================
# SUBSCRIPTION & BILLING — School Admin side (renewal + bKash/Nagad claim)
# =============================================================================
def page_billing():
    school_id = st.session_state["school_id"]
    hero("💳 Subscription & Billing", "আপনার স্কুলের সাবস্ক্রিপশন স্ট্যাটাস ও পেমেন্ট।")
    locked, sub, msg = get_subscription_status(school_id)
    if locked:
        st.error(f"🔒 {msg} নবায়ন না করা পর্যন্ত অন্যান্য সব ফিচার লক থাকবে।")
    else:
        st.success(f"✅ {msg}")
    if sub is not None:
        c1, c2, c3 = st.columns(3)
        c1.metric("Plan", str(sub.get("PlanType", "")))
        c2.metric("Amount", f"৳{sub.get('Amount','')}")
        c3.metric("Expiry", str(sub.get("ExpiryDate", "")))

    st.markdown("#### 💰 নবায়ন করতে পেমেন্ট পাঠান")
    bkash = st.secrets.get("BKASH_NUMBER", "সেট করা হয়নি — secrets.toml-এ BKASH_NUMBER দিন")
    nagad = st.secrets.get("NAGAD_NUMBER", "সেট করা হয়নি — secrets.toml-এ NAGAD_NUMBER দিন")
    st.info(f"📱 **bKash** (Send Money): **{bkash}**  \n📱 **Nagad** (Send Money): **{nagad}**")
    hint("উপরের নম্বরে টাকা পাঠিয়ে নিচে Transaction ID (TrxID)-সহ ফর্মটি জমা দিন। SuperAdmin যাচাই করে "
         "অনুমোদন করলেই আপনার সাবস্ক্রিপশন সাথে সাথে সচল হয়ে যাবে — এখানে টাকা কাটার কিছু নেই, শুধু "
         "প্রমাণ জমা দেওয়ার ফর্ম।")
    with st.form("submit_payment_claim"):
        c1, c2 = st.columns(2)
        gateway = c1.selectbox("Payment Method", ["bKash", "Nagad", "Rocket", "Bank Transfer"])
        amount = c2.number_input("Amount Sent (৳)", min_value=0)
        trx = st.text_input("Transaction ID (TrxID)", placeholder="যেমন: 8N7A6QZK2X — SMS-এ পাবেন")
        if st.form_submit_button("✅ Submit for Approval", type="primary"):
            if not trx.strip():
                st.error("Transaction ID দিন।")
            else:
                payid = f"{school_id.replace('SCH-','')}-{datetime.now().strftime('%y%m%d%H%M%S')}"
                append_row("PaymentLogs", {
                    "SchoolID": school_id, "PaymentID": payid,
                    "SubscriptionID": sub.get("SubscriptionID", "") if sub is not None else "",
                    "TrxID": trx.strip(), "Gateway": gateway, "Amount": amount,
                    "PaymentDate": datetime.now().strftime("%Y-%m-%d"), "Status": "Pending",
                })
                st.success("জমা হয়েছে — SuperAdmin অনুমোদন করলে সাবস্ক্রিপশন সচল হয়ে যাবে। নিচে স্ট্যাটাস দেখতে পাবেন।")
                st.rerun()

    st.markdown("#### 🧾 আমার পেমেন্ট হিস্ট্রি")
    pays = read_df("PaymentLogs", school_id=school_id)
    if not pays.empty:
        st.dataframe(pays.sort_values("PaymentDate", ascending=False), use_container_width=True, hide_index=True)
    else:
        st.caption("এখনো কোনো পেমেন্ট জমা দেওয়া হয়নি।")


# =============================================================================
# SUPPORT CHAT — School Admin <-> SuperAdmin in-app messaging
# =============================================================================
def page_support_chat():
    school_id = st.session_state["school_id"]
    school_name = st.session_state.get("school_name", school_id)
    hero("💬 Support Chat", "SuperAdmin-এর সাথে সরাসরি মেসেজ আদান-প্রদান করুন।")
    render_chat_thread(school_id, school_name, sender_default=st.session_state.get("user_name", "Admin"),
                        sender_role_default=st.session_state.get("role", ROLE_ADMIN))


def render_chat_thread(school_id: str, school_name: str, sender_default: str, sender_role_default: str):
    msgs = read_df("SupportMessages", school_id=school_id)
    try:
        box = st.container(height=360, border=True)
    except TypeError:
        box = st.container()
    with box:
        if msgs.empty:
            st.caption("এখনো কোনো মেসেজ নেই — নিচে থেকে প্রথম মেসেজ পাঠান।")
        else:
            for _, m in msgs.sort_values("Timestamp").iterrows():
                who = "🛡️ SuperAdmin" if m.get("SenderRole") == ROLE_SUPERADMIN else f"🏫 {m.get('Sender','')}"
                align = "right" if m.get("SenderRole") == ROLE_SUPERADMIN else "left"
                bg = "#dbeafe" if m.get("SenderRole") == ROLE_SUPERADMIN else "#f1f5f9"
                st.markdown(
                    f"<div style='text-align:{align};margin:6px 0;'>"
                    f"<span style='background:{bg};padding:6px 12px;border-radius:10px;display:inline-block;"
                    f"max-width:80%;color:#0f172a;'><b>{who}</b><br/>{m.get('Message','')}"
                    f"<br/><span style='font-size:0.7rem;color:#64748b;'>{m.get('Timestamp','')}</span></span></div>",
                    unsafe_allow_html=True,
                )
    with st.form(f"send_msg_{school_id}", clear_on_submit=True):
        text = st.text_area("মেসেজ লিখুন", height=70, placeholder="এখানে লিখুন...")
        if st.form_submit_button("📤 Send", type="primary") and text.strip():
            append_row("SupportMessages", {
                "SchoolID": school_id, "MessageID": f"MSG-{datetime.now().strftime('%Y%m%d%H%M%S%f')}",
                "Sender": sender_default, "SenderRole": sender_role_default,
                "Message": text.strip(), "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            st.rerun()


def find_junior_proxy_teacher(teachers_df: pd.DataFrame, exclude_ids: set):
    """Auto-Proxy Scheduler: when a teacher is absent/on leave, pick the most
    JUNIOR available active teacher (latest JoiningDate = most junior) who
    isn't already excluded (absent teacher + everyone already on duty that
    day), so exam duty coverage never has a gap."""
    if teachers_df.empty:
        return None
    pool = teachers_df[
        (teachers_df.get("IsActive", "Yes") != "No") & (~teachers_df["TeacherID"].isin(exclude_ids))
    ].copy()
    if pool.empty:
        return None
    pool["_join"] = pd.to_datetime(pool.get("JoiningDate", ""), errors="coerce")
    pool = pool.sort_values("_join", ascending=False, na_position="last")  # latest joining = most junior
    return pool.iloc[0]


def page_exam_duties():
    school_id, role, actor = st.session_state["school_id"], st.session_state["role"], st.session_state["user_name"]
    hero("🛡️ Guard List / Exam Duties", "Assign invigilators per room per exam-day — prints in the official Invigilation Duty & Script Distribution Chart format.")

    exams = read_df("Exams", school_id=school_id)
    routines = read_df("Routines", school_id=school_id)
    teachers = read_df("Teachers", school_id=school_id)
    if exams.empty or teachers.empty:
        st.info("Add Exams and Teachers first.")
        return
    teacher_lookup = teachers.set_index("TeacherID")["Name"].to_dict()

    exam_label = st.selectbox("Exam", (exams["ExamName"] + " (" + exams["ExamID"] + ")").tolist())
    exam_id = exam_label.split("(")[-1].rstrip(")")
    exam_name = exams[exams["ExamID"] == exam_id].iloc[0]["ExamName"]
    day_routines = routines[routines["ExamID"] == exam_id] if not routines.empty else pd.DataFrame()
    day_options = sorted(day_routines["ExamDate"].dropna().unique().tolist()) if not day_routines.empty else []

    hint("প্রতিটি রুমের জন্য একটি করে Duty সেভ করুন — Shift ঠিক ছবির মতো লিখুন, যেমন: "
         "'SHIFT - 1: MORNING (09:00 AM - 11:30 AM) | Classes: Play to 5th'। একই Shift-এর সব রুম "
         "প্রিন্টে একসাথে একটা টেবিলে দেখাবে।")

    if role in ADMIN_LIKE_ROLES:
        with st.expander("➕ Assign / Update a Room's Duty", expanded=True):
            with st.form("add_duty"):
                c1, c2 = st.columns(2)
                exam_date = c1.selectbox("Exam Date (from Routine)", day_options) if day_options else c1.text_input(
                    "Exam Date (YYYY-MM-DD)", placeholder="18/09/2026")
                room_no = c2.text_input("Room No.", placeholder="যেমন: Room - 101")
                shift = st.text_input(
                    "Shift (ঠিক এভাবে লিখুন — প্রিন্টে এই লেখাটাই শিরোনাম হিসেবে যাবে)",
                    placeholder="SHIFT - 1: MORNING (09:00 AM - 11:30 AM) | Classes: Play to 5th",
                    help="একই Shift টেক্সট যেসব রুমে বসাবেন, প্রিন্টের সময় তারা একসাথে একই ব্লকে দেখাবে।",
                )
                assigned_classes = st.text_area(
                    "Assigned Classes & Students",
                    placeholder="Class 1(D), Class 2(D)\n40 Students",
                    help="এই রুমে কোন কোন ক্লাসের ছাত্র বসবে ও মোট কতজন — একাধিক লাইনে লিখতে পারেন, প্রিন্টে যেমন লিখবেন তেমনই দেখাবে।",
                    height=70,
                )
                total_scripts = st.number_input(
                    "Total Scripts Needed", min_value=0, step=1,
                    help="এই রুমে মোট কতজন পরীক্ষা দেবে — সাধারণত Assigned Students-এর সমান।",
                )
                teacher_names = (teachers["Name"] + " (" + teachers["TeacherID"] + ")").tolist()
                invigilators = st.multiselect("Invigilator(s) — one room can have more than one", teacher_names)
                excluded = st.multiselect(
                    "Teachers who must NOT be assigned here (their own private students sit in this room)",
                    teacher_names,
                )
                notes = st.text_input("Notes (optional)")
                if st.form_submit_button("Save Duty", type="primary"):
                    inv_ids = [t.split("(")[-1].rstrip(")") for t in invigilators][:3]
                    excl_ids = [t.split("(")[-1].rstrip(")") for t in excluded]
                    conflict = set(inv_ids) & set(excl_ids)
                    if conflict:
                        st.error("A teacher can't be both an invigilator and excluded for the same room. Fix the overlap.")
                    else:
                        # Conflict check: same teacher, same exam date, different room already assigned.
                        existing_duties = read_df("ExamDuties", school_id=school_id)
                        clash = []
                        if not existing_duties.empty:
                            same_day = existing_duties[
                                (existing_duties["ExamID"] == exam_id) & (existing_duties.get("ExamDate", "") == exam_date)
                                if "ExamDate" in existing_duties.columns else existing_duties["ExamID"] == exam_id
                            ]
                            for _, row_ in same_day.iterrows():
                                if row_.get("RoomNo") == room_no:
                                    continue
                                for c in ["Invigilator1_ID", "Invigilator2_ID", "Invigilator3_ID"]:
                                    if row_.get(c) in inv_ids:
                                        clash.append(teacher_lookup.get(row_.get(c), row_.get(c)))
                        if clash:
                            st.warning(f"⚠️ Already assigned to another room the same day: {', '.join(set(clash))}. Saved anyway — please review.")
                        did = next_seq_id(school_id, "ExamDuties", "DutyID", "DUTY")
                        row = {
                            "SchoolID": school_id, "DutyID": did, "ExamID": exam_id, "ExamDate": exam_date,
                            "Shift": shift, "RoomNo": room_no, "AssignedClasses": assigned_classes,
                            "TotalScriptsNeeded": total_scripts, "Invigilator1_ID": inv_ids[0] if len(inv_ids) > 0 else "",
                            "Invigilator2_ID": inv_ids[1] if len(inv_ids) > 1 else "",
                            "Invigilator3_ID": inv_ids[2] if len(inv_ids) > 2 else "",
                            "PrivateTutorTeacherIDs": ",".join(excl_ids), "Status": "Assigned", "Notes": notes,
                        }
                        append_row("ExamDuties", row)
                        st.success("Duty saved.")
                        st.rerun()

    duties = read_df("ExamDuties", school_id=school_id)
    duties = duties[duties["ExamID"] == exam_id] if not duties.empty else duties
    if duties is not None and not duties.empty:
        print_date_options = sorted(duties["ExamDate"].dropna().unique().tolist())
        print_date = st.selectbox("🖨️ Print chart for date", print_date_options, key="guardlist_print_date")
        day_duties = duties[duties["ExamDate"] == print_date]

        if role in ADMIN_LIKE_ROLES:
            with st.expander("🔄 Auto-Proxy Scheduler — অনুপস্থিত/ছুটিতে থাকা শিক্ষকের বদলে অটো-এসাইন"):
                hint("যে শিক্ষক আজ ডিউটিতে অনুপস্থিত/ছুটিতে আছেন তাকে বেছে নিন — অ্যাপ স্বয়ংক্রিয়ভাবে "
                     "সবচেয়ে জুনিয়র (সর্বশেষ যোগদানকারী), সেদিন অন্য কোথাও ডিউটিতে নেই এমন সক্রিয় "
                     "শিক্ষককে সেই ডিউটিতে বসিয়ে দেবে — কভারেজে কোনো ফাঁক থাকবে না।")
                inv_map = []
                for _, d in day_duties.iterrows():
                    for col in ["Invigilator1_ID", "Invigilator2_ID", "Invigilator3_ID"]:
                        tid = d.get(col, "")
                        if tid:
                            inv_map.append((d.get("DutyID"), col, tid, d.get("RoomNo", "")))
                if not inv_map:
                    st.caption("এই দিনে কোনো ইনভিজিলেটর এসাইন করা নেই।")
                else:
                    labels = [f"{teacher_lookup.get(tid, tid)} — Room {room}" for _, _, tid, room in inv_map]
                    pick_idx = st.selectbox("আজ কে অনুপস্থিত/ছুটিতে?", range(len(labels)), format_func=lambda i: labels[i])
                    if st.button("🔄 Auto-Assign Junior Proxy", type="primary"):
                        did, col, absent_tid, _ = inv_map[pick_idx]
                        already_on_duty_today = {t for _, _, t, _ in inv_map}
                        proxy = find_junior_proxy_teacher(teachers, already_on_duty_today)
                        if proxy is None:
                            st.error("কোনো খালি (আজ অন্য কোথাও ডিউটিতে নেই এমন) সক্রিয় শিক্ষক পাওয়া যায়নি।")
                        else:
                            duties_all = read_df("ExamDuties", school_id=school_id)
                            mask = duties_all["DutyID"] == did
                            duties_all.loc[mask, col] = proxy["TeacherID"]
                            note_add = (f" | Auto-proxy {datetime.now().strftime('%Y-%m-%d %H:%M')}: "
                                        f"{teacher_lookup.get(absent_tid, absent_tid)} অনুপস্থিত → "
                                        f"{proxy.get('Name','')} (Joining: {proxy.get('JoiningDate','')}) প্রতিস্থাপিত")
                            duties_all.loc[mask, "Notes"] = duties_all.loc[mask, "Notes"].astype(str) + note_add
                            upsert_rows("ExamDuties", duties_all[mask], key_cols=["DutyID"])
                            st.success(f"✅ {teacher_lookup.get(absent_tid, absent_tid)}-এর বদলে "
                                       f"{proxy.get('Name','')} (সবচেয়ে জুনিয়র, খালি) অটো-এসাইন করা হলো।")
                            st.rerun()

        show = day_duties.copy()
        for c in ["Invigilator1_ID", "Invigilator2_ID", "Invigilator3_ID"]:
            show[c] = show[c].map(lambda t: teacher_lookup.get(t, t))
        cols = [c for c in ["Shift", "RoomNo", "AssignedClasses", "TotalScriptsNeeded", "Invigilator1_ID",
                             "Invigilator2_ID", "Invigilator3_ID", "PrivateTutorTeacherIDs", "Status"] if c in show.columns]
        st.dataframe(show[cols], use_container_width=True, hide_index=True)

        html = render_guard_list(
            read_df("Schools")[read_df("Schools")["SchoolID"] == school_id].iloc[0],
            print_date, f"Exam: {exam_name}", day_duties, teacher_lookup,
        )
        st.markdown(html, unsafe_allow_html=True)
        print_button()
    else:
        st.info("No duties assigned yet for this exam.")


# =============================================================================
# SEAT PLAN (auto bench allocation, per exam / room)
# =============================================================================
def page_seatplan():
    school_id, role, actor = st.session_state["school_id"], st.session_state["role"], st.session_state["user_name"]
    hero("🪑 Seat Plan", "Auto-generate exam bench/seat allocation by roll number.")

    exams = read_df("Exams", school_id=school_id)
    students, classes = _class_section_options(school_id)
    if exams.empty or not classes:
        st.info("Add Exams and Students first.")
        return

    c1, c2, c3, c4 = st.columns(4)
    exam_label = c1.selectbox("Exam", (exams["ExamName"] + " (" + exams["ExamID"] + ")").tolist())
    exam_id = exam_label.split("(")[-1].rstrip(")")
    exam_name = exams[exams["ExamID"] == exam_id].iloc[0]["ExamName"]
    klass = c2.selectbox("Class", classes)
    sections = sorted(students[students["Class"] == klass]["Section"].dropna().unique().tolist())
    section = c3.selectbox("Section", sections if sections else ["A"])
    room_no = c4.text_input("Room No.", value="1")
    capacity = st.number_input("Benches per row / seats used for numbering", min_value=1, value=2)

    if role in ADMIN_LIKE_ROLES and st.button("⚙️ Auto-Generate Seat Plan", type="primary"):
        cs = students[(students["Class"] == klass) & (students["Section"] == section)
                      & (students.get("Status", "Active") == "Active")].copy()
        cs = cs.sort_values("Roll", key=lambda s: pd.to_numeric(s, errors="coerce")).reset_index(drop=True)
        rows = []
        for i, r in cs.iterrows():
            bench_no = (i // int(capacity)) + 1
            seat_pos = "Left" if i % int(capacity) == 0 else "Right"
            rows.append({
                "SchoolID": school_id, "SeatPlanID": next_seq_id(school_id, "SeatPlans_Attendance", "SeatPlanID", "SEAT"),
                "ExamID": exam_id, "RoomNo": room_no, "StudentID": r["StudentID"], "Class": klass,
                "Section": section, "Roll": r["Roll"], "Gender": r.get("Gender", ""),
                "BenchNo": bench_no, "SeatPosition": seat_pos, "ExamDate": "", "IsPresent": "", "AttendanceRemarks": "",
            })
        df = pd.DataFrame(rows)
        for _, row in df.iterrows():
            append_row("SeatPlans_Attendance", row.to_dict())
        st.session_state["_last_seatplan"] = df
        st.success(f"Generated seats for {len(df)} students in Room {room_no}.")

    plan = st.session_state.get("_last_seatplan")
    if plan is None:
        all_plans = read_df("SeatPlans_Attendance", school_id=school_id)
        plan = all_plans[(all_plans["ExamID"] == exam_id) & (all_plans["RoomNo"] == room_no)] if not all_plans.empty else pd.DataFrame()
    if plan is not None and not plan.empty:
        st.dataframe(plan, use_container_width=True, hide_index=True)
        schools = read_df("Schools")
        school_row = schools[schools["SchoolID"] == school_id].iloc[0] if not schools.empty else pd.Series()
        st.markdown(render_seat_plan(school_row, exam_name, room_no, plan), unsafe_allow_html=True)
        print_button()
    else:
        st.info("No seat plan generated yet for this Exam / Room.")


# =============================================================================
# SCRIPT (ANSWER-SHEET) HANDOVER / RETURN REGISTER — খাতা জমা-নেওয়া
# =============================================================================
def page_script_distribution():
    school_id, role, actor = st.session_state["school_id"], st.session_state["role"], st.session_state["user_name"]
    hero("📒 Script Register", "Log answer-script handover to teachers and their return.")

    exams = read_df("Exams", school_id=school_id)
    subjects = read_df("Subjects", school_id=school_id)
    teachers = read_df("Teachers", school_id=school_id)
    if exams.empty or teachers.empty:
        st.info("Add Exams and Teachers first.")
        return
    subject_lookup = subjects.set_index("SubjectID")["SubjectName"].to_dict() if not subjects.empty else {}

    exam_label = st.selectbox("Exam", (exams["ExamName"] + " (" + exams["ExamID"] + ")").tolist())
    exam_id = exam_label.split("(")[-1].rstrip(")")
    exam_name = exams[exams["ExamID"] == exam_id].iloc[0]["ExamName"]

    with st.expander("➕ Hand Over Scripts to a Teacher", expanded=True):
        with st.form("handover_scripts"):
            c1, c2, c3 = st.columns(3)
            subj_label = c1.selectbox("Subject", (subjects["SubjectName"] + " (" + subjects["SubjectID"] + ")").tolist()) if not subjects.empty else None
            klass = c2.text_input("Class")
            section = c3.text_input("Section")
            teacher_label = st.selectbox("Teacher", (teachers["Name"] + " (" + teachers["TeacherID"] + ")").tolist())
            count = st.number_input("Total Scripts Handed Over", min_value=0, step=1)
            if st.form_submit_button("Save Handover", type="primary") and subj_label:
                subject_id = subj_label.split("(")[-1].rstrip(")")
                tid = teacher_label.split("(")[-1].rstrip(")")
                did = next_seq_id(school_id, "ScriptDistribution", "DistributionID", "SCR")
                append_row("ScriptDistribution", {
                    "SchoolID": school_id, "DistributionID": did, "ExamID": exam_id, "SubjectID": subject_id,
                    "Class": klass, "Section": section, "TeacherID": tid, "TotalScriptsHandedOver": count,
                    "HandoverDate": datetime.now().strftime("%Y-%m-%d"), "ReturnStatus": "Pending",
                })
                st.success("Handover logged.")
                st.rerun()

    dist = read_df("ScriptDistribution", school_id=school_id)
    dist = dist[dist["ExamID"] == exam_id] if not dist.empty else dist
    if dist is not None and not dist.empty:
        pending = dist[dist["ReturnStatus"] != "Returned"]
        if not pending.empty:
            with st.expander("↩️ Mark Scripts as Returned"):
                labels = (pending["DistributionID"] + " — " + pending["TeacherID"] + " — " + pending["SubjectID"]).tolist()
                pick = st.selectbox("Pick a handover record", labels)
                dist_id = pick.split(" — ")[0]
                returned_count = st.number_input("Returned Scripts Count", min_value=0, step=1)
                if st.button("Mark Returned"):
                    dist.loc[dist["DistributionID"] == dist_id, "ReturnedScriptsCount"] = returned_count
                    dist.loc[dist["DistributionID"] == dist_id, "ReturnDate"] = datetime.now().strftime("%Y-%m-%d")
                    dist.loc[dist["DistributionID"] == dist_id, "ReturnStatus"] = "Returned"
                    upsert_rows("ScriptDistribution", dist[dist["DistributionID"] == dist_id], key_cols=["DistributionID"])
                    st.success("Marked returned.")
                    st.rerun()
        st.dataframe(dist, use_container_width=True, hide_index=True)
        schools = read_df("Schools")
        school_row = schools[schools["SchoolID"] == school_id].iloc[0] if not schools.empty else pd.Series()
        st.markdown(render_script_sheet(school_row, exam_name, dist, subject_lookup), unsafe_allow_html=True)
        print_button()
    else:
        st.info("No script handovers logged yet for this exam.")


# =============================================================================
# CERTIFICATES — প্রত্যয়নপত্র / প্রশংসাপত্র / ছাড়পত্র (auto-generated)
# =============================================================================
def page_certificates():
    school_id = st.session_state["school_id"]
    hero("📜 Certificates", "Auto-generate Testimonial, Transfer Certificate & Certification letters.")

    schools = read_df("Schools")
    school_row = schools[schools["SchoolID"] == school_id].iloc[0] if not schools.empty else pd.Series()
    students, classes = _class_section_options(school_id)
    if not classes:
        st.info("Add Students first.")
        return

    c1, c2 = st.columns(2)
    klass = c1.selectbox("Class", classes)
    cs = students[students["Class"] == klass]
    student_label = c2.selectbox(
        "Student", (cs["StudentName"] + " — Roll " + cs["Roll"] + " (" + cs["StudentID"] + ")").tolist()
    ) if not cs.empty else None
    cert_type = st.radio(
        "Certificate Type",
        ["testimonial", "transfer", "certification"],
        format_func=lambda k: {"testimonial": "প্রশংসাপত্র (Testimonial)", "transfer": "ছাড়পত্র (Transfer Certificate)",
                                "certification": "প্রত্যয়নপত্র (Certification)"}[k],
        horizontal=True,
    )
    extra = {}
    if cert_type == "transfer":
        c3, c4 = st.columns(2)
        extra["reason"] = c3.text_input("Reason for leaving (optional — overrides Students sheet)")
        extra["character"] = c4.text_input("Character remark", value="Good")
        extra["dues"] = st.text_input("Dues status", value="cleared")
    elif cert_type == "certification":
        extra["purpose_line"] = st.text_input(
            "Purpose line", value="This certificate is issued on the student/guardian's request for necessary purposes."
        )
    else:
        extra["character"] = st.text_input("Character remark", value="good and satisfactory")

    if student_label and st.button("🖨️ Generate Certificate", type="primary"):
        sid = student_label.split("(")[-1].rstrip(")")
        student_row = cs[cs["StudentID"] == sid].iloc[0]
        html = render_certificate(school_row, student_row, cert_type, extra)
        st.markdown(html, unsafe_allow_html=True)
        print_button()


GENERIC_SHEETS = {
    "Job Applications": ("JobApplications", "JobAppID"),
    "Admission Applications": ("Applications", "AppID"),
    "Contact Messages": ("ContactMessages", "MessageID"),
}
# Note: Seat Plans, Exam Duties (Guard List) and Script Distribution now have
# their own dedicated, rule-aware pages (see page_seatplan / page_exam_duties /
# page_script_distribution below) instead of the raw generic editor.


def page_other_records():
    school_id = st.session_state["school_id"]
    hero("🗂️ Other Records", "Admissions, duties, seat plans, script logs & messages.")
    choice = st.selectbox("Choose a record type", list(GENERIC_SHEETS.keys()))
    sheet_name, id_col = GENERIC_SHEETS[choice]
    df = read_df(sheet_name, school_id=school_id)
    edited = st.data_editor(df, use_container_width=True, num_rows="dynamic", hide_index=True, key=f"editor_{sheet_name}")
    if st.button("💾 Save Changes"):
        edited["SchoolID"] = school_id
        for i, r in edited.iterrows():
            if not str(r.get(id_col, "")).strip():
                edited.at[i, id_col] = next_seq_id(school_id, sheet_name, id_col, id_col[:3].upper())
        upsert_rows(sheet_name, edited, key_cols=[id_col])
        st.success("Saved.")
        st.rerun()


def page_guardian_result():
    school_id = st.session_state["school_id"]
    student_id = st.session_state.get("guardian_student_id", "")
    roll_input = str(st.session_state.get("guardian_roll", "")).strip()
    hero("🎒 My Result", st.session_state["school_name"])
    schools = read_df("Schools")
    school_row = schools[schools["SchoolID"] == school_id].iloc[0] if not schools.empty else pd.Series()
    students = read_df("Students", school_id=school_id)
    student_row_df = students[students["StudentID"] == student_id]
    if student_row_df.empty:
        st.error("Student ID not found. দয়া করে সঠিক Student ID দিন।")
        if st.button("← Back to Login"):
            do_logout()
        return
    student_row = student_row_df.iloc[0]
    if roll_input and str(student_row.get("Roll", "")).strip() != roll_input:
        st.error("Student ID ঠিক আছে, কিন্তু Roll নম্বর মিলছে না। আবার চেষ্টা করুন।")
        if st.button("← Back to Login"):
            do_logout()
        return
    st.success(f"Showing results for **{student_row.get('StudentName')}** (Class {student_row.get('Class')}, Roll {student_row.get('Roll')})")

    results = read_df("Results", school_id=school_id)
    my_results = results[(results["StudentID"] == student_id) & (results["Published"] == "Yes")] if not results.empty else pd.DataFrame()
    exams = read_df("Exams", school_id=school_id)
    subjects = read_df("Subjects", school_id=school_id)
    marks = read_df("Marks", school_id=school_id)

    if my_results.empty:
        st.info("No published results available yet.")
    for _, res in my_results.iterrows():
        exam_name = exams[exams["ExamID"] == res["ExamID"]].iloc[0]["ExamName"] if not exams[exams["ExamID"] == res["ExamID"]].empty else res["ExamID"]
        with st.expander(f"📄 {exam_name}", expanded=True):
            m = marks[(marks["StudentID"] == student_id) & (marks["ExamID"] == res["ExamID"])]
            html = render_marksheet(school_row, student_row, res, m, subjects, exam_name)
            st.markdown(html, unsafe_allow_html=True)
            print_button()

    cons = read_df("ConsolidatedResults", school_id=school_id)
    my_cons = cons[(cons["StudentID"] == student_id) & (cons["Published"] == "Yes")] if not cons.empty else pd.DataFrame()
    if not my_cons.empty:
        st.markdown("### 🏆 Consolidated Annual Result")
        st.dataframe(
            my_cons[["Session", "ConsolidatedPercentage", "ConsolidatedGPA", "ConsolidatedGrade", "Status", "PromotedClass"]],
            use_container_width=True, hide_index=True,
        )
    if st.button("← Back to Login"):
        do_logout()


# =============================================================================
# SIDEBAR NAVIGATION & MAIN ROUTER
# =============================================================================
NAV_BY_ROLE = {
    ROLE_SUPERADMIN: [("🏠 Dashboard", "dashboard"), ("🛡️ SuperAdmin Console", "superadmin")],
    ROLE_ADMIN: [
        ("🏠 Dashboard", "dashboard"), ("📝 Marks Entry", "marks"),
        ("📊 Generate & Publish Results", "results"), ("🏆 Consolidated & Promotion", "consolidated"),
        ("🖨️ Print Center", "print"), ("📜 Certificates", "certificates"),
        ("🎒 Students", "students"), ("👩‍🏫 Teachers", "teachers"),
        ("🗓️ Routines", "routines"), ("🛡️ Guard List / Duties", "duties"),
        ("🪑 Seat Plan", "seatplan"), ("📒 Script Register", "scripts"),
        ("📢 Notices", "notices"), ("🗂️ Other Records", "other"),
        ("💳 Subscription & Billing", "billing"), ("💬 Support Chat", "chat"), ("👤 My Profile", "profile"),
    ],
    ROLE_TEACHER: [
        ("🏠 Dashboard", "dashboard"), ("📝 Marks Entry", "marks"),
        ("🖨️ Print Center", "print"), ("🗓️ Routines", "routines"),
        ("🛡️ Guard List / Duties", "duties"), ("📢 Notices", "notices"),
        ("🎒 Students", "students"), ("👤 My Profile", "profile"),
    ],
    ROLE_CLERK: [("🏠 Dashboard", "dashboard"), ("📢 Notices", "notices"), ("👤 My Profile", "profile")],
    # Blank-Role staff: view-only limited access, per spec ("তাদের কোনো role
    # থাকবে না" -> সীমিত তথ্য দেখতে পারবে)।
    ROLE_STAFF: [("🏠 Dashboard", "dashboard"), ("🗓️ Routines", "routines"),
                 ("📢 Notices", "notices"), ("👤 My Profile", "profile")],
}
# Headteacher/Headmaster get the same menu as Admin
NAV_BY_ROLE["Headmaster"] = NAV_BY_ROLE[ROLE_ADMIN]
NAV_BY_ROLE["Headteacher"] = NAV_BY_ROLE[ROLE_ADMIN]
# The ONLY menu a school is allowed to see once its subscription has expired —
# everything else is locked until a SuperAdmin approves a renewal payment.
NAV_WHEN_LOCKED = [("💳 Subscription & Billing", "billing"), ("💬 Support Chat", "chat"), ("👤 My Profile", "profile")]

PAGE_FUNCS = {
    "dashboard": page_dashboard, "marks": page_marks_entry, "results": page_generate_results,
    "consolidated": page_consolidated, "print": page_print_center, "students": page_students,
    "teachers": page_teachers, "routines": page_routines, "notices": page_notices,
    "other": page_other_records, "superadmin": page_superadmin,
    "duties": page_exam_duties, "seatplan": page_seatplan, "scripts": page_script_distribution,
    "certificates": page_certificates, "profile": page_my_profile,
    "billing": page_billing, "chat": page_support_chat,
}


def sidebar_nav():
    role = st.session_state["role"]
    with st.sidebar:
        st.markdown(f"### 🎓 School Manager BD")
        st.caption(f"{st.session_state['school_name']}")
        st.markdown(f"**{st.session_state['user_name']}**  \n`{role or 'Staff'}`")

        if st.session_state.get("impersonating"):
            st.warning("🎭 **Preview Mode চালু আছে** — এটা শুধু টেস্টিং, আসল ডাটা এডিট করবেন না।")
            if st.button("↩️ Exit Preview — Back to SuperAdmin", use_container_width=True, type="primary"):
                st.session_state["role"] = st.session_state.pop("real_role")
                st.session_state["school_id"] = st.session_state.pop("real_school_id")
                st.session_state["school_name"] = st.session_state.pop("real_school_name")
                st.session_state["user_name"] = st.session_state.pop("real_user_name")
                st.session_state["impersonating"] = False
                st.session_state["nav"] = "superadmin"
                st.rerun()

        st.markdown("---")
        locked = False
        if role != ROLE_SUPERADMIN and not st.session_state.get("impersonating"):
            locked, _, _ = get_subscription_status(st.session_state["school_id"])
        items = NAV_WHEN_LOCKED if locked else NAV_BY_ROLE.get(role, NAV_BY_ROLE[ROLE_STAFF])
        if locked and st.session_state.get("nav") not in [k for _, k in NAV_WHEN_LOCKED]:
            st.session_state["nav"] = "billing"
        if "nav" not in st.session_state:
            st.session_state["nav"] = items[0][1]
        for label, key in items:
            if st.button(label, key=f"nav_{key}", use_container_width=True):
                st.session_state["nav"] = key
        st.markdown("---")
        if st.button("🚪 Logout", use_container_width=True):
            do_logout()


def main():
    init_session()
    if st.session_state.get("guardian_mode"):
        page_guardian_result()
        return
    if not st.session_state["logged_in"]:
        login_page()
        return
    sidebar_nav()
    page = PAGE_FUNCS.get(st.session_state.get("nav", "dashboard"), page_dashboard)
    page()


if __name__ == "__main__":
    main()
