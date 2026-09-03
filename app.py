import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Zero-Cost Multi-Tenant School Management System",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- GOOGLE SHEETS API CONNECTION SETUP ---
@st.cache_resource
def init_connection():
    try:
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            client = gspread.authorize(creds)
            # Open spreadsheet 'ResultManager'
            sheet = client.open("ResultManager")
            return sheet
        else:
            return None
    except Exception as e:
        return None

spreadsheet = init_connection()

# --- HELPER FUNCTIONS TO LOAD/SAVE DATA ---
def load_data(worksheet_name):
    if spreadsheet is None:
        return pd.DataFrame()
    try:
        ws = spreadsheet.worksheet(worksheet_name)
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        return pd.DataFrame()

def save_data(worksheet_name, df):
    if spreadsheet is None:
        return False
    try:
        ws = spreadsheet.worksheet(worksheet_name)
        ws.clear()
        if not df.empty:
            ws.update([df.columns.values.tolist()] + df.values.tolist())
        return True
    except Exception as e:
        return False

def append_row(worksheet_name, row_dict):
    if spreadsheet is None:
        return False
    try:
        ws = spreadsheet.worksheet(worksheet_name)
        df = load_data(worksheet_name)
        if df.empty:
            new_df = pd.DataFrame([row_dict])
        else:
            new_df = pd.concat([df, pd.DataFrame([row_dict])], ignore_index=True)
        return save_data(worksheet_name, new_df)
    except Exception as e:
        return False

# --- SESSION STATE INITIALIZATION ---
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
if "user_role" not in st.session_state:
    st.session_state.user_role = None
if "teacher_id" not in st.session_state:
    st.session_state.teacher_id = None
if "school_id" not in st.session_state:
    st.session_state.school_id = None
if "teacher_name" not in st.session_state:
    st.session_state.teacher_name = None

# --- STYLING CUSTOMIZATION (CSS) ---
st.markdown("""
    <style>
    .main { background-color: #f8f9fa; }
    .stButton>button { width: 100%; border-radius: 6px; font-weight: bold; }
    .card { background: white; padding: 20px; border-radius: 10px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); margin-bottom: 20px; }
    .metric-card { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 15px; border-radius: 8px; text-align: center; }
    </style>
""", unsafe_allow_html=True)

# --- SIDEBAR NAVIGATION ---
st.sidebar.title("🎓 School Management Portal")
st.sidebar.markdown("---")

# Load Schools for multi-tenant selector
schools_df = load_data("Schools")
if not schools_df.empty and "SchoolID" in schools_df.columns:
    school_options = dict(zip(schools_df["SchoolName"], schools_df["SchoolID"]))
    selected_school_name = st.sidebar.selectbox("Select Institution", list(school_options.keys()))
    current_school_id = school_options[selected_school_name]
    st.session_state.school_id = current_school_id
else:
    current_school_id = "SCH-000001"
    st.session_state.school_id = current_school_id

st.sidebar.markdown("---")

# --- AUTHENTICATION & LOGIN SYSTEM ---
if not st.session_state.logged_in:
    st.sidebar.subheader("🔒 Staff PIN Login")
    login_teacher_id = st.sidebar.text_input("Teacher ID / Username", placeholder="e.g. ADM-000001-T001")
    login_pin = st.sidebar.text_input("4-Digit Secret PIN", type="password", placeholder="e.g. 6029")
    
    if st.sidebar.button("Login"):
        teachers_df = load_data("Teachers")
        if not teachers_df.empty:
            match = teachers_df[
                (teachers_df["SchoolID"] == current_school_id) & 
                (teachers_df["TeacherID"].astype(str) == str(login_teacher_id)) & 
                (teachers_df["PIN"].astype(str) == str(login_pin))
            ]
            if not match.empty:
                st.session_state.logged_in = True
                st.session_state.user_role = match.iloc[0]["Role"]
                st.session_state.teacher_id = match.iloc[0]["TeacherID"]
                st.session_state.teacher_name = match.iloc[0]["Name"]
                st.sidebar.success(f"Welcome, {st.session_state.teacher_name}!")
                st.rerun()
            else:
                st.sidebar.error("Invalid Teacher ID or Secret PIN!")
        else:
            st.sidebar.error("Teacher database is empty or connection failed.")
    
    st.sidebar.markdown("---")
    menu = st.sidebar.radio("Public Navigation", ["Home / Dashboard", "Check Result", "Student Admission Apply", "Job Application", "Contact Us"])
else:
    st.sidebar.success(f"Logged in as: **{st.session_state.teacher_name}** ({st.session_state.user_role})")
    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.user_role = None
        st.session_state.teacher_id = None
        st.session_state.teacher_name = None
        st.rerun()
    
    st.sidebar.markdown("---")
    if st.session_state.user_role == "Admin":
        menu = st.sidebar.radio("Admin Dashboard", [
            "Home / Dashboard", "Manage Students", "Manage Teachers", "Exam Management", 
            "Marks Entry", "Result Processing", "Consolidated Results", "Routine & Seat Plan", 
            "Exam Duty Assignment", "Script Distribution", "Notices", "Settings"
        ])
    else:
        menu = st.sidebar.radio("Teacher Dashboard", [
            "Home / Dashboard", "Marks Entry", "Check Result", "Notices"
        ])

# --- MAIN APP ROUTING & MODULES ---

if menu == "Home / Dashboard":
    st.title("🏫 Welcome to Smart School Management System")
    st.markdown("A Zero-Cost, Multi-Tenant Cloud Powered School & Result Management Solution.")
    
    col1, col2, col3, col4 = st.columns(4)
    students_df = load_data("Students")
    teachers_df = load_data("Teachers")
    exams_df = load_data("Exams")
    notices_df = load_data("Notices")
    
    total_students = len(students_df[students_df["SchoolID"] == current_school_id]) if not students_df.empty else 0
    total_teachers = len(teachers_df[teachers_df["SchoolID"] == current_school_id]) if not teachers_df.empty else 0
    total_exams = len(exams_df[exams_df["SchoolID"] == current_school_id]) if not exams_df.empty else 0
    total_notices = len(notices_df[notices_df["SchoolID"] == current_school_id]) if not notices_df.empty else 0
    
    with col1:
        st.metric("Total Students", total_students)
    with col2:
        st.metric("Total Teachers", total_teachers)
    with col3:
        st.metric("Active Exams", total_exams)
    with col4:
        st.metric("Published Notices", total_notices)
        
    st.markdown("---")
    st.subheader("📢 Recent Notices & Announcements")
    if not notices_df.empty:
        school_notices = notices_df[notices_df["SchoolID"] == current_school_id]
        for idx, row in school_notices.iterrows():
            st.info(f"**{row.get('Title', '')}** ({row.get('Date', '')})\n\n{row.get('Description', '')}")
    else:
        st.write("No notices available at the moment.")

elif menu == "Check Result":
    st.title("📊 Student Result & Marksheet Portal (SSC Standard)")
    st.markdown("Enter student details to view term-wise marks, GPA, and grade breakdown.")
    
    exams_df = load_data("Exams")
    students_df = load_data("Students")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        selected_class = st.selectbox("Select Class", ["Play", "Nursery", "Class-1", "Class-2", "Class-3", "Class-4", "Class-5", "Class-6", "Class-7", "Class-8", "Class-9", "Class-10"])
    with col2:
        selected_section = st.selectbox("Select Section", ["A", "B", "C"])
    with col3:
        student_roll = st.number_input("Enter Roll Number", min_value=1, step=1, value=1)
        
    if not exams_df.empty:
        exam_opts = exams_df[exams_df["SchoolID"] == current_school_id]["ExamName"].tolist()
        selected_exam_name = st.selectbox("Select Examination", exam_opts if exam_opts else ["1st Term (30%)"])
    else:
        selected_exam_name = "1st Term (30%)"
        
    if st.button("Search Result"):
        results_df = load_data("Results")
        marks_df = load_data("Marks")
        
        if not students_df.empty and not results_df.empty:
            stud_match = students_df[
                (students_df["SchoolID"] == current_school_id) & 
                (students_df["Class"] == selected_class) & 
                (students_df["Section"] == selected_section) & 
                (students_df["Roll"] == student_roll)
            ]
            if not stud_match.empty:
                student_id = stud_match.iloc[0]["StudentID"]
                student_name = stud_match.iloc[0]["StudentName"]
                
                st.success(f"Result Found for: **{student_name}** (Roll: {student_roll}, Class: {selected_class})")
                
                # Filter marks
                stu_marks = marks_df[(marks_df["SchoolID"] == current_school_id) & (marks_df["StudentID"] == student_id)]
                if not stu_marks.empty:
                    st.dataframe(stu_marks[["SubjectID", "FullMarks", "Written", "MCQ", "Practical", "ObtainedMarks", "Grade", "SubjectGPA", "IsPass"]])
                else:
                    st.warning("No subject marks found for this student.")
            else:
                st.error("No student found with this Roll/Class/Section combination.")
        else:
            st.warning("Database records are currently empty.")

elif menu == "Manage Students" and st.session_state.user_role == "Admin":
    st.title("👩‍🎓 Student Management")
    st.subheader("Add New Student")
    
    with st.form("add_student_form"):
        col1, col2, col3 = st.columns(3)
        with col1:
            s_name = st.text_input("Student Name")
            s_class = st.selectbox("Class", ["Play", "Nursery", "Class-1", "Class-2", "Class-3", "Class-4", "Class-5", "Class-6", "Class-7", "Class-8", "Class-9", "Class-10"])
            s_roll = st.number_input("Roll No", min_value=1, step=1)
        with col2:
            s_father = st.text_input("Father's Name")
            s_section = st.selectbox("Section", ["A", "B", "C"])
            s_gender = st.selectbox("Gender", ["Male", "Female", "Other"])
        with col3:
            s_mother = st.text_input("Mother's Name")
            s_phone = st.text_input("Phone Number")
            s_session = st.text_input("Session", value="2026")
            
        submitted = st.form_submit_button("Register Student")
        if submitted:
            student_id = f"{current_school_id}-26-S{int(datetime.now().timestamp())}"
            new_student = {
                "SchoolID": current_school_id,
                "StudentID": student_id,
                "Roll": s_roll,
                "Class": s_class,
                "Section": s_section,
                "Group": "Core",
                "OptionalSubjectID": "",
                "StudentName": s_name,
                "FatherName": s_father,
                "MotherName": s_mother,
                "StudentAddress": "",
                "Phone": s_phone,
                "Status": "Active",
                "DateOfBirth": "2015-01-01",
                "Gender": s_gender,
                "Religion": "Islam",
                "BloodGroup": "A+",
                "Session": s_session,
                "PassingYear": "",
                "TCReason": "",
                "CharacterStatus": "Good",
                "DuesStatus": "Paid",
                "Photo": "",
                "IssueDate": str(datetime.now().date())
            }
            if append_row("Students", new_student):
                st.success("Student added successfully!")
            else:
                st.error("Failed to add student to Google Sheet.")
                
    st.subheader("Existing Students List")
    students_df = load_data("Students")
    if not students_df.empty:
        st.dataframe(students_df[students_df["SchoolID"] == current_school_id])

elif menu == "Manage Teachers" and st.session_state.user_role == "Admin":
    st.title("👨‍🏫 Teacher Management & PIN Control")
    teachers_df = load_data("Teachers")
    if not teachers_df.empty:
        st.dataframe(teachers_df[teachers_df["SchoolID"] == current_school_id])
    else:
        st.write("No teacher records found.")

elif menu == "Marks Entry":
    st.title("📝 Subject-wise Marks Entry")
    st.markdown("Teachers can input written, MCQ, and practical marks securely.")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        m_class = st.selectbox("Class", ["Play", "Nursery", "Class-1", "Class-2", "Class-3", "Class-4", "Class-5", "Class-6", "Class-7", "Class-8", "Class-9", "Class-10"], key="m_class")
    with col2:
        m_section = st.selectbox("Section", ["A", "B", "C"], key="m_section")
    with col3:
        m_subject = st.text_input("Subject Name / ID", value="Bangla")
        
    with st.form("marks_form"):
        student_roll_input = st.number_input("Student Roll No", min_value=1, value=1)
        written_marks = st.number_input("Written Marks", min_value=0.0, max_value=100.0, value=80.0)
        mcq_marks = st.number_input("MCQ Marks", min_value=0.0, max_value=50.0, value=0.0)
        practical_marks = st.number_input("Practical Marks", min_value=0.0, max_value=50.0, value=0.0)
        
        save_marks = st.form_submit_button("Save Marks Entry")
        if save_marks:
            total_obtained = written_marks + mcq_marks + practical_marks
            mark_row = {
                "SchoolID": current_school_id,
                "MarkID": f"{current_school_id}-M-{int(datetime.now().timestamp())}",
                "StudentID": f"{current_school_id}-26-S0001",
                "ExamID": f"{current_school_id}-26-EX1",
                "Class": m_class,
                "Section": m_section,
                "Session": "2026",
                "SubjectID": m_subject,
                "FullMarks": 100,
                "Written": written_marks,
                "MCQ": mcq_marks,
                "Practical": practical_marks,
                "ObtainedMarks": total_obtained,
                "Percentage": total_obtained,
                "SubjectGPA": 5.0,
                "Grade": "A+",
                "IsPass": "Yes",
                "IsPresent": "Yes"
            }
            if append_row("Marks", mark_row):
                st.success("Marks saved successfully!")
            else:
                st.error("Failed to save marks.")

elif menu == "Consolidated Results" and st.session_state.user_role == "Admin":
    st.title("📈 Consolidated Result & Multi-Term Promotion")
    consolidated_df = load_data("ConsolidatedResults")
    if not consolidated_df.empty:
        st.dataframe(consolidated_df[consolidated_df["SchoolID"] == current_school_id])
    else:
        st.info("No consolidated result records found yet.")

elif menu == "Exam Duty Assignment" and st.session_state.user_role == "Admin":
    st.title("📋 Exam Duty & Private Tuition Conflict Prevention")
    st.markdown("Ensure teachers are not assigned duty in rooms where their private tuition students sit.")
    duties_df = load_data("ExamDuties")
    if not duties_df.empty:
        st.dataframe(duties_df[duties_df["SchoolID"] == current_school_id])
    else:
        st.info("No exam duties assigned yet.")

elif menu == "Script Distribution" and st.session_state.user_role == "Admin":
    st.title("📦 Script Handover & Return Tracking")
    script_df = load_data("ScriptDistribution")
    if not script_df.empty:
        st.dataframe(script_df[script_df["SchoolID"] == current_school_id])
    else:
        st.info("No script distribution records found.")

elif menu == "Notices":
    st.title("📢 School Notice Board")
    notices_df = load_data("Notices")
    if not notices_df.empty:
        for idx, row in notices_df[notices_df["SchoolID"] == current_school_id].iterrows():
            st.markdown(f"### {row.get('Title', '')}")
            st.write(f"**Date:** {row.get('Date', '')}")
            st.write(row.get('Description', ''))
            st.markdown("---")
    else:
        st.write("No notices posted.")

elif menu == "Settings" and st.session_state.user_role == "Admin":
    st.title("⚙️ School Settings & Configuration")
    settings_df = load_data("Settings")
    if not settings_df.empty:
        st.dataframe(settings_df[settings_df["SchoolID"] == current_school_id])
    else:
        st.write("No specific settings configurations found.")

elif menu == "Student Admission Apply":
    st.title("📝 Online Student Admission Application")
    with st.form("admission_apply"):
        app_name = st.text_input("Applicant Name")
        guardian_name = st.text_input("Guardian Name")
        app_phone = st.text_input("Phone Number")
        app_class = st.selectbox("Applying for Class", ["Play", "Nursery", "Class-1", "Class-2", "Class-3"])
        prev_school = st.text_input("Previous School (if any)")
        birth_cert = st.text_input("Birth Certificate No")
        
        submitted_app = st.form_submit_button("Submit Application")
        if submitted_app:
            new_app = {
                "SchoolID": current_school_id,
                "AppID": f"{current_school_id}-26-A{int(datetime.now().timestamp())}",
                "StudentName": app_name,
                "GuardianName": guardian_name,
                "Phone": app_phone,
                "ApplyingClass": app_class,
                "PreviousSchool": prev_school,
                "BirthCertNo": birth_cert,
                "Photo": "",
                "ApplyDate": str(datetime.now().date()),
                "Status": "Pending"
            }
            if append_row("Applications", new_app):
                st.success("Admission application submitted successfully! Admin will review soon.")
            else:
                st.error("Submission failed.")

elif menu == "Job Application":
    st.title("💼 Teacher / Staff Job Application Portal")
    with st.form("job_apply"):
        j_name = st.text_input("Full Name")
        j_phone = st.text_input("Mobile Number")
        j_email = st.text_input("Email Address")
        j_subject = st.text_input("Applying Subject / Post")
        j_qual = st.text_input("Educational Qualification")
        
        submitted_job = st.form_submit_button("Submit Job Application")
        if submitted_job:
            new_job = {
                "SchoolID": current_school_id,
                "JobAppID": f"{current_school_id}-26-J{int(datetime.now().timestamp())}",
                "ApplicantName": j_name,
                "Phone": j_phone,
                "Email": j_email,
                "Subject": j_subject,
                "Qualification": j_qual,
                "CV_File": "",
                "ApplyDate": str(datetime.now().date()),
                "Status": "Pending"
            }
            if append_row("JobApplications", new_job):
                st.success("Job application submitted successfully!")
            else:
                st.error("Job application submission failed.")

elif menu == "Contact Us":
    st.title("📞 Contact School Authority")
    with st.form("contact_form"):
        c_name = st.text_input("Your Name")
        c_phone = st.text_input("Your Phone")
        c_role = st.selectbox("You are", ["Guardian", "Student", "Teacher", "Other"])
        c_subject = st.text_input("Subject")
        c_message = st.text_area("Message")
        
        submitted_msg = st.form_submit_button("Send Message")
        if submitted_msg:
            msg_row = {
                "SchoolID": current_school_id,
                "MessageID": f"{current_school_id}-MSG{int(datetime.now().timestamp())}",
                "SenderName": c_name,
                "Phone": c_phone,
                "SenderRole": c_role,
                "Subject": c_subject,
                "Message": c_message,
                "Date": str(datetime.now().date()),
                "Status": "Unread"
            }
            if append_row("ContactMessages", msg_row):
                st.success("Message sent successfully! We will contact you soon.")
            else:
                st.error("Failed to send message.")

# --- FOOTER ---
st.markdown("---")
st.markdown("<p style='text-align: center; color: gray;'>Zero-Cost School Management System | Powered by Streamlit & Google Sheets API</p>", unsafe_allow_html=True)
