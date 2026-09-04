# -*- coding: utf-8 -*-
"""
EduManage BD — Multi-Tenant School Management & Automated Result System
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
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# =============================================================================
# APP CONFIG
# =============================================================================
st.set_page_config(
    page_title="EduManage BD | School Result & Management System",
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
}

# =============================================================================
# STYLING — "beautiful" themed UI
# =============================================================================
CUSTOM_CSS = """
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    .stApp {
        background: linear-gradient(180deg, #f4f7fb 0%, #eef2f9 100%);
    }
    .em-hero {
        background: linear-gradient(120deg, #1e3a8a 0%, #2563eb 55%, #0ea5e9 100%);
        padding: 28px 32px; border-radius: 18px; color: white;
        margin-bottom: 22px; box-shadow: 0 10px 30px rgba(30,58,138,0.25);
    }
    .em-hero h1 { margin: 0; font-size: 1.7rem; font-weight: 800; }
    .em-hero p { margin: 4px 0 0 0; opacity: 0.92; font-size: 0.95rem; }

    .em-card {
        background: white; border-radius: 16px; padding: 20px 22px;
        box-shadow: 0 4px 18px rgba(15, 23, 42, 0.06);
        border: 1px solid rgba(15,23,42,0.05); margin-bottom: 16px;
    }
    .em-stat {
        background: white; border-radius: 14px; padding: 16px 18px;
        border-left: 5px solid #2563eb; box-shadow: 0 4px 14px rgba(15,23,42,0.05);
    }
    .em-stat .val { font-size: 1.6rem; font-weight: 800; color: #1e3a8a; }
    .em-stat .lbl { font-size: 0.82rem; color: #64748b; font-weight: 600; text-transform: uppercase; letter-spacing: .04em;}

    .em-badge { display:inline-block; padding: 3px 12px; border-radius: 999px; font-size: 0.78rem; font-weight: 700; }
    .em-pass { background:#dcfce7; color:#15803d; }
    .em-fail { background:#fee2e2; color:#b91c1c; }
    .em-pending { background:#fef9c3; color:#a16207; }

    div[data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0f172a 0%, #1e293b 100%);
    }
    div[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
    div[data-testid="stSidebar"] .stButton>button {
        background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.12);
        width: 100%; text-align: left; border-radius: 10px; margin-bottom: 4px;
    }
    div[data-testid="stSidebar"] .stButton>button:hover { background: #2563eb; border-color:#2563eb; }

    .stButton>button {
        border-radius: 10px; font-weight: 600; padding: 0.5rem 1.1rem;
    }
    .stButton>button[kind="primary"] { background: #2563eb; }

    .marksheet {
        background: white; padding: 26px 34px; border-radius: 10px;
        border: 2px solid #1e3a8a; font-family: 'Georgia', serif;
    }
    .marksheet h2 { text-align:center; color:#1e3a8a; margin: 2px 0; }
    .marksheet .sub { text-align:center; color:#475569; font-size: 0.9rem; margin-bottom: 10px;}
    .ms-table { width:100%; border-collapse: collapse; margin-top: 14px; }
    .ms-table th, .ms-table td { border: 1px solid #94a3b8; padding: 7px 10px; font-size: 0.92rem; }
    .ms-table th { background: #1e3a8a; color: white; }
    .ms-table tr:nth-child(even) { background: #f1f5f9; }

    @media print {
        div[data-testid="stSidebar"], .stButton, header, footer { display:none !important; }
        .em-hero { display:none !important; }
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def hero(title: str, subtitle: str = ""):
    st.markdown(
        f"""<div class="em-hero"><h1>🎓 {title}</h1><p>{subtitle}</p></div>""",
        unsafe_allow_html=True,
    )


def stat_card(col, label, value):
    col.markdown(
        f"""<div class="em-stat"><div class="val">{value}</div><div class="lbl">{label}</div></div>""",
        unsafe_allow_html=True,
    )


def badge(text, kind="pending"):
    cls = {"pass": "em-pass", "fail": "em-fail", "pending": "em-pending"}.get(kind, "em-pending")
    return f'<span class="em-badge {cls}">{text}</span>'


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
# CLASS ORDER (for promotion) & GRADING ENGINE
# =============================================================================
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
              "user_name", "guardian_mode"]:
        st.session_state[k] = False if k in ("logged_in", "guardian_mode") else None
    st.rerun()


def login_page():
    hero("EduManage BD", "Multi-Tenant School Management & Automated Result System")
    tab_school, tab_super, tab_guardian = st.tabs(
        ["🏫 School Staff Login", "🛡️ SuperAdmin", "🎒 Check Result (Student/Guardian)"]
    )

    with tab_school:
        st.markdown('<div class="em-card">', unsafe_allow_html=True)
        schools = read_df("Schools")
        if schools.empty:
            st.warning("No schools found in the sheet yet.")
        else:
            active = schools[schools.get("IsActive", "Yes") == "Yes"] if "IsActive" in schools.columns else schools
            options = active["SchoolName"] + "  (" + active["SchoolID"] + ")"
            choice = st.selectbox("Select your School", options.tolist() if not active.empty else [])
            teacher_id = st.text_input("Teacher / Staff ID")
            pin = st.text_input("PIN", type="password")
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
        pw = st.text_input("SuperAdmin Password", type="password", key="sa_pw")
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

    with tab_guardian:
        st.markdown('<div class="em-card">', unsafe_allow_html=True)
        st.caption("View a published result without logging in.")
        schools = read_df("Schools")
        if not schools.empty:
            options = schools["SchoolName"] + "  (" + schools["SchoolID"] + ")"
            g_choice = st.selectbox("School", options.tolist(), key="g_school")
            g_student_id = st.text_input("Student ID (e.g. 000001-26-S0001)")
            if st.button("View Result"):
                g_school_id = g_choice.split("(")[-1].rstrip(")")
                st.session_state.update({
                    "guardian_mode": True, "school_id": g_school_id,
                    "school_name": g_choice.split("  (")[0], "role": "Guardian",
                })
                st.session_state["guardian_student_id"] = g_student_id.strip()
                st.rerun()
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
    # ---- Merit ranking + Next Roll, grouped per PromotedClass ----
    cdf["NextRoll"] = ""
    for pclass, grp in cdf.groupby("PromotedClass"):
        ranked = grp[grp["Status"] == "Passed"].sort_values(
            ["ConsolidatedGPA", "ConsolidatedTotal"], ascending=[False, False]
        )
        for i, idx in enumerate(ranked.index, start=1):
            cdf.at[idx, "NextRoll"] = str(i)

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
        <h2>{school_row.get('SchoolName','')}</h2>
        <div class="sub">{school_row.get('Address','')} &nbsp;|&nbsp; EIIN: {school_row.get('SchoolEIIN','N/A')}</div>
        <h3 style="text-align:center;text-decoration:underline;">ACADEMIC MARKSHEET — {exam_name}</h3>
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
    sched_rows = ""
    for _, r in routine_df.iterrows():
        sched_rows += f"<tr><td>{r.get('ExamDate','')}</td><td>{r.get('DayName','')}</td><td>{r.get('StartTime','')} - {r.get('EndTime','')}</td></tr>"
    html = f"""
    <div class="marksheet" style="max-width:640px;margin:auto;">
        <h2>{school_row.get('SchoolName','')}</h2>
        <div class="sub">Admit Card — {exam_name}</div>
        <table style="width:100%;margin-top:10px;">
            <tr><td><b>Name:</b> {student_row.get('StudentName','')}</td><td><b>ID:</b> {student_row.get('StudentID','')}</td></tr>
            <tr><td><b>Class:</b> {student_row.get('Class','')}</td><td><b>Section:</b> {student_row.get('Section','')}</td></tr>
            <tr><td><b>Roll:</b> {student_row.get('Roll','')}</td><td><b>Session:</b> {student_row.get('Session','')}</td></tr>
        </table>
        <table class="ms-table"><tr><th>Date</th><th>Day</th><th>Time</th></tr>{sched_rows}</table>
        <div style="display:flex;justify-content:space-between;margin-top:50px;">
            <div>_____________________<br/>Student Signature</div>
            <div>_____________________<br/>Headteacher</div>
        </div>
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
        stat_card(c1, "Total Schools", len(schools))
        stat_card(c2, "Active Schools", (schools.get("IsActive", pd.Series(dtype=str)) == "Yes").sum() if not schools.empty else 0)
        stat_card(c3, "Active Subscriptions", (subs.get("Status", pd.Series(dtype=str)) == "Active").sum() if not subs.empty else 0)
        total_paid = num(pays["Amount"]).sum() if not pays.empty and "Amount" in pays.columns else 0
        stat_card(c4, "Total Collected (৳)", f"{total_paid:,.0f}")
        st.markdown("### Recent Schools")
        st.dataframe(schools, use_container_width=True, hide_index=True)
        return

    students = read_df("Students", school_id=school_id)
    teachers = read_df("Teachers", school_id=school_id)
    results = read_df("Results", school_id=school_id)
    c1, c2, c3, c4 = st.columns(4)
    stat_card(c1, "Total Students", len(students) if not students.empty else 0)
    stat_card(c2, "Total Teachers", len(teachers) if not teachers.empty else 0)
    stat_card(c3, "Results Generated", len(results) if not results.empty else 0)
    published = (results.get("Published", pd.Series(dtype=str)) == "Yes").sum() if not results.empty else 0
    stat_card(c4, "Published Results", published)

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


def _class_section_options(school_id):
    students = read_df("Students", school_id=school_id)
    classes = sorted(students["Class"].dropna().unique().tolist()) if not students.empty else []
    return students, classes


def page_marks_entry():
    role, school_id, actor = st.session_state["role"], st.session_state["school_id"], st.session_state["user_name"]
    hero("📝 Marks Entry", "Enter marks — grade, GPA and pass/fail are calculated live.")

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
    if cls_students.empty:
        st.info("No active students in this Class / Section.")
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
            c1, c2 = st.columns(2)
            exam_label2 = c1.selectbox("Exam", (exams2["ExamName"] + " (" + exams2["ExamID"] + ")").tolist(), key="ac_exam")
            exam_id2 = exam_label2.split("(")[-1].rstrip(")")
            klass2 = c2.selectbox("Class", classes2, key="ac_class")
            cs2 = students2[students2["Class"] == klass2]
            student_label2 = st.selectbox(
                "Student", (cs2["StudentName"] + " — Roll " + cs2["Roll"] + " (" + cs2["StudentID"] + ")").tolist(), key="ac_student"
            ) if not cs2.empty else None
            if student_label2:
                sid2 = student_label2.split("(")[-1].rstrip(")")
                srow2 = cs2[cs2["StudentID"] == sid2].iloc[0]
                routines = read_df("Routines", school_id=school_id)
                routines = routines[routines["ExamID"] == exam_id2]
                exam_name2 = exams2[exams2["ExamID"] == exam_id2].iloc[0]["ExamName"]
                html2 = render_admit_card(school_row, srow2, exam_name2, routines)
                st.markdown(html2, unsafe_allow_html=True)
                print_button()


def page_teachers():
    school_id = st.session_state["school_id"]
    hero("👩‍🏫 Teachers & Staff", "Manage staff accounts. PINs are masked for security.")
    teachers = read_df("Teachers", school_id=school_id)

    with st.expander("➕ Add New Teacher / Staff"):
        with st.form("add_teacher"):
            c1, c2, c3 = st.columns(3)
            name = c1.text_input("Name")
            role_ = c2.selectbox("Role", [ROLE_ADMIN, ROLE_TEACHER, ROLE_CLERK, "Headmaster"])
            designation = c3.text_input("Designation")
            c4, c5, c6 = st.columns(3)
            klass = c4.text_input("Class (if class teacher)")
            section = c5.text_input("Section")
            subject = c6.text_input("Subject")
            c7, c8, c9 = st.columns(3)
            phone = c7.text_input("Phone")
            email = c8.text_input("Email")
            pin = c9.text_input("Set PIN", value=str(np.random.randint(1000, 9999)))
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
    students = read_df("Students", school_id=school_id)

    with st.expander("➕ Add New Student"):
        with st.form("add_student"):
            c1, c2, c3 = st.columns(3)
            name = c1.text_input("Student Name")
            klass = c2.text_input("Class")
            section = c3.text_input("Section")
            c4, c5, c6 = st.columns(3)
            roll = c4.text_input("Roll")
            session_year = c5.text_input("Session", value=datetime.now().strftime("%Y"))
            gender = c6.selectbox("Gender", ["Male", "Female", "Other"])
            c7, c8 = st.columns(2)
            father = c7.text_input("Father's Name")
            mother = c8.text_input("Mother's Name")
            phone = st.text_input("Guardian Phone")
            if st.form_submit_button("Add Student", type="primary"):
                sid = next_seq_id(school_id, "Students", "StudentID", "S")
                new = pd.DataFrame([{
                    "SchoolID": school_id, "StudentID": sid, "Roll": roll, "Class": klass,
                    "Section": section, "Group": "Core", "StudentName": name, "FatherName": father,
                    "MotherName": mother, "Phone": phone, "Status": "Active", "Gender": gender,
                    "Session": session_year,
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
        with st.form("add_notice"):
            title = st.text_input("Title")
            desc = st.text_area("Description")
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
            with st.form("add_routine"):
                c1, c2 = st.columns(2)
                exam_label = c1.selectbox("Exam", (exams["ExamName"] + " (" + exams["ExamID"] + ")").tolist())
                exam_id = exam_label.split("(")[-1].rstrip(")")
                shift = c2.text_input("Shift", value="SHIFT - 1: MORNING")
                c3, c4, c5 = st.columns(3)
                classes_txt = c3.text_input("Classes (label)")
                class_group = c4.text_input("Class Group (comma list)")
                exam_date = c5.date_input("Exam Date")
                c6, c7 = st.columns(2)
                start_t = c6.time_input("Start Time")
                end_t = c7.time_input("End Time")
                if st.form_submit_button("Add", type="primary"):
                    rid = next_seq_id(school_id, "Routines", "RoutineID", "ROU")
                    append_row("Routines", {
                        "SchoolID": school_id, "RoutineID": rid, "ExamID": exam_id, "Shift": shift,
                        "Classes": classes_txt, "ClassGroup": class_group,
                        "ExamDate": exam_date.strftime("%Y-%m-%d"), "DayName": exam_date.strftime("%A"),
                        "StartTime": start_t.strftime("%H:%M"), "EndTime": end_t.strftime("%H:%M"),
                    })
                    st.success("Routine added.")
                    st.rerun()

    routines = read_df("Routines", school_id=school_id)
    if not routines.empty:
        cols = [c for c in ["ExamID", "Shift", "Classes", "ExamDate", "DayName", "StartTime", "EndTime"] if c in routines.columns]
        st.dataframe(routines[cols], use_container_width=True, hide_index=True)
    else:
        st.info("No routines added yet.")


def page_superadmin():
    hero("🛡️ SuperAdmin Console", "Schools, subscriptions & payment tracking (bKash / Nagad).")
    tab1, tab2, tab3 = st.tabs(["🏫 Schools", "💳 Subscriptions", "🧾 Payment Logs"])

    with tab1:
        schools = read_df("Schools")
        st.dataframe(schools, use_container_width=True, hide_index=True)
        with st.expander("➕ Add School"):
            with st.form("add_school"):
                c1, c2 = st.columns(2)
                name = c1.text_input("School Name")
                phone = c2.text_input("Phone")
                address = st.text_input("Address")
                exam_system = st.selectbox("Exam System", ["3-Term", "2-Term", "Semester"])
                if st.form_submit_button("Create School", type="primary"):
                    n = len(schools) + 1
                    sid = f"SCH-{n:06d}"
                    append_row("Schools", {
                        "SchoolID": sid, "SchoolName": name, "Address": address, "Phone": phone,
                        "IsActive": "Yes", "ExamSystem": exam_system,
                    })
                    st.success(f"School created: {sid}")
                    st.rerun()

    with tab2:
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

    with tab3:
        pays = read_df("PaymentLogs")
        st.dataframe(pays, use_container_width=True, hide_index=True)
        with st.expander("➕ Log a Payment (bKash / Nagad)"):
            with st.form("add_pay"):
                schools = read_df("Schools")
                sch_label = st.selectbox("School", (schools["SchoolName"] + " (" + schools["SchoolID"] + ")").tolist(), key="pay_school") if not schools.empty else None
                gateway = st.selectbox("Gateway", ["bKash", "Nagad", "Bank", "Cash"])
                trx = st.text_input("Transaction ID")
                amount = st.number_input("Amount (৳) ", min_value=0)
                if st.form_submit_button("Log Payment", type="primary") and sch_label:
                    sid = sch_label.split("(")[-1].rstrip(")")
                    payid = f"{sid.replace('SCH-','')}-{datetime.now().strftime('%y')}-PAY{len(pays)+1:03d}"
                    append_row("PaymentLogs", {
                        "SchoolID": sid, "PaymentID": payid, "TrxID": trx, "Gateway": gateway,
                        "Amount": amount, "PaymentDate": datetime.now().strftime("%Y-%m-%d"), "Status": "SUCCESS",
                    })
                    st.success("Payment logged.")
                    st.rerun()


GENERIC_SHEETS = {
    "Job Applications": ("JobApplications", "JobAppID"),
    "Admission Applications": ("Applications", "AppID"),
    "Seat Plans / Attendance": ("SeatPlans_Attendance", "SeatPlanID"),
    "Exam Duties": ("ExamDuties", "DutyID"),
    "Script Distribution": ("ScriptDistribution", "DistributionID"),
    "Contact Messages": ("ContactMessages", "MessageID"),
}


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
    hero("🎒 My Result", st.session_state["school_name"])
    schools = read_df("Schools")
    school_row = schools[schools["SchoolID"] == school_id].iloc[0] if not schools.empty else pd.Series()
    students = read_df("Students", school_id=school_id)
    student_row_df = students[students["StudentID"] == student_id]
    if student_row_df.empty:
        st.error("Student ID not found. Please check and try again.")
        if st.button("← Back to Login"):
            do_logout()
        return
    student_row = student_row_df.iloc[0]
    st.success(f"Showing results for **{student_row.get('StudentName')}** (Class {student_row.get('Class')})")

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
        ("🖨️ Print Center", "print"), ("🎒 Students", "students"), ("👩‍🏫 Teachers", "teachers"),
        ("🗓️ Routines", "routines"), ("📢 Notices", "notices"), ("🗂️ Other Records", "other"),
    ],
    ROLE_TEACHER: [
        ("🏠 Dashboard", "dashboard"), ("📝 Marks Entry", "marks"),
        ("🖨️ Print Center", "print"), ("🗓️ Routines", "routines"), ("📢 Notices", "notices"),
    ],
    ROLE_CLERK: [("🏠 Dashboard", "dashboard"), ("📢 Notices", "notices"), ("🗓️ Routines", "routines")],
}
# Headteacher/Headmaster get the same menu as Admin
NAV_BY_ROLE["Headmaster"] = NAV_BY_ROLE[ROLE_ADMIN]
NAV_BY_ROLE["Headteacher"] = NAV_BY_ROLE[ROLE_ADMIN]

PAGE_FUNCS = {
    "dashboard": page_dashboard, "marks": page_marks_entry, "results": page_generate_results,
    "consolidated": page_consolidated, "print": page_print_center, "students": page_students,
    "teachers": page_teachers, "routines": page_routines, "notices": page_notices,
    "other": page_other_records, "superadmin": page_superadmin,
}


def sidebar_nav():
    role = st.session_state["role"]
    with st.sidebar:
        st.markdown(f"### 🎓 EduManage BD")
        st.caption(f"{st.session_state['school_name']}")
        st.markdown(f"**{st.session_state['user_name']}**  \n`{role}`")
        st.markdown("---")
        items = NAV_BY_ROLE.get(role, NAV_BY_ROLE[ROLE_TEACHER])
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
