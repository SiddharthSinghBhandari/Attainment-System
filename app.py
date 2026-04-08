import streamlit as st
import matplotlib.pyplot as plt
import pandas as pd
from db import students, users, uploads
from calculate import calculate_attainment
import os
import time   # ✅ ADDED
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ---------------- LOAD CUSTOM CSS ----------------
def load_css():
    """Load external CSS file and hide default Streamlit UI elements"""
    with open("styles.css") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

    st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    </style>
    """, unsafe_allow_html=True)

load_css()


# ---------------- FOOTER ----------------
def show_footer():
    st.markdown("""
        <div class="custom-footer">
            Developed by Siddharth Singh Bhandari In Collaboration With Amit,Nishant,Tushar | Attainment System 🚀
        </div>
    """, unsafe_allow_html=True)


# ---------------- SESSION STATE ----------------
if "page" not in st.session_state:
    st.session_state.page = "home"

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# ✅ ADD (for reset)
if "reset_key" not in st.session_state:
    st.session_state.reset_key = 0

# ✅ ADD (flicker fix)
if "just_logged_in" not in st.session_state:
    st.session_state.just_logged_in = False


# ---------------- NAVIGATION BAR ----------------
def navbar():
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if st.button("Admin Register", use_container_width=True):
            st.session_state.page = "admin_register"

    with col2:
        if st.button("Admin Login", use_container_width=True):
            st.session_state.page = "admin_login"

    with col3:
        if st.button("Teacher Register", use_container_width=True):
            st.session_state.page = "teacher_register"

    with col4:
        if st.button("Teacher Login", use_container_width=True):
            st.session_state.page = "teacher_login"

navbar()

# ✅ FIX (no stop → no blank page)
if st.session_state.just_logged_in:
    st.session_state.just_logged_in = False


# ---------------- APP TITLE ----------------
st.markdown("<h1> Attainment System</h1>", unsafe_allow_html=True)


# ---------------- LOGO DISPLAY ----------------
logo_url = "https://fimt-ggsipu.org/images/flogo2025-1.jpg"
col1, col2, col3 = st.columns([1, 4, 1])
with col2:
    st.image(logo_url, width=450)


# ---------------- ADMIN REGISTER ----------------
if st.session_state.page == "admin_register":

    st.header("Create Admin Account")

    with st.form("admin_register_form"):
        username = st.text_input("Admin Username")
        password = st.text_input("Password", type="password")

        submitted = st.form_submit_button("Create Admin")

        if submitted:
            if users.find_one({"role": "admin"}):
                st.error("Admin already exists")
            elif username.strip() == "" or password.strip() == "":
                st.error("Fill all fields")
            else:
                users.insert_one({
                    "username": username,
                    "password": password,
                    "role": "admin"
                })
                st.success("Admin account created")


# ---------------- ADMIN LOGIN ----------------
elif st.session_state.page == "admin_login":

    st.header("Admin Login")

    with st.form("admin_login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        submitted = st.form_submit_button("Login")

        if submitted:
            user = users.find_one({
                "username": username,
                "password": password,
                "role": "admin"
            })

            if user:
                st.session_state.logged_in = True
                st.session_state.page = "admin_panel"
                st.session_state.just_logged_in = True
                st.rerun()
            else:
                st.error("Invalid credentials")


# ---------------- ADMIN DASHBOARD ----------------
elif st.session_state.page == "admin_panel" and st.session_state.logged_in:

    st.header("Admin Dashboard")

    if st.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.page = "home"
        st.rerun()

    st.subheader("Registered Teachers")

    teachers = list(users.find({"role": "teacher"}))

    if teachers:
        for t in teachers:
            col1, col2 = st.columns([6, 1])

            with col1:
                st.write(f"👨‍🏫 {t['username']}")

            with col2:
                if st.button("Delete", key=str(t["_id"])):
                    users.delete_one({"_id": t["_id"]})
                    st.success("Teacher deleted")
                    st.rerun()
    else:
        st.info("No teachers registered")


# ---------------- TEACHER REGISTER ----------------
elif st.session_state.page == "teacher_register":

    st.header("Create Teacher Account")

    with st.form("teacher_register_form"):
        username = st.text_input("Teacher Username")
        password = st.text_input("Password", type="password")

        submitted = st.form_submit_button("Register Teacher")

        if submitted:
            if username.strip() == "" or password.strip() == "":
                st.error("Fill all fields")
            else:
                users.insert_one({
                    "username": username,
                    "password": password,
                    "role": "teacher"
                })
                st.success("Teacher registered successfully")


# ---------------- TEACHER LOGIN ----------------
elif st.session_state.page == "teacher_login":

    st.header("Teacher Login")

    with st.form("teacher_login_form"):
        username = st.text_input("Username")
        password = st.text_input("Password", type="password")

        submitted = st.form_submit_button("Login")

        if submitted:
            user = users.find_one({
                "username": username,
                "password": password,
                "role": "teacher"
            })

            if user:
                st.session_state.logged_in = True
                st.session_state.teacher = username
                st.session_state.page = "teacher_panel"
                st.session_state.just_logged_in = True
                st.rerun()
            else:
                st.error("Invalid credentials")


# ---------------- TEACHER DASHBOARD ----------------
elif st.session_state.page == "teacher_panel" and st.session_state.logged_in:

    st.header("Teacher Dashboard")

    if st.button("Logout"):
        st.session_state.logged_in = False
        st.session_state.page = "home"
        st.rerun()

    # -------- TEMPLATE DOWNLOAD --------
    st.subheader("Download Template")
    if os.path.exists("Template clean.xlsx"):
        with open("Template clean.xlsx", "rb") as f:
            st.download_button("Download Template", f, file_name="Template clean.xlsx")

    # -------- CLASS CONFIGURATION --------
    st.subheader("📊 Class Configuration")
    total_students_input = st.number_input("Total Students", min_value=1, value=1)
    max_marks_input = st.number_input("Maximum Marks", min_value=1, value=1)

    use_weight = st.checkbox("Enable Weightage (Optional)")
    weight_input = 100
    if use_weight:
        weight_input = st.number_input("Weight (%)", min_value=1, max_value=100, value=100)

                 # -------- FILE UPLOAD --------
    st.subheader("📂 Upload Excel/CSV")
    uploaded_file = st.file_uploader("Upload student file", type=["xlsx", "csv"])

    if uploaded_file is not None:
        if uploaded_file.name.endswith(".csv"):
            df_upload = pd.read_csv(uploaded_file)
        else:
            df_upload = pd.read_excel(uploaded_file, engine="openpyxl")

        st.dataframe(df_upload)

        # session flag
        if "data_saved" not in st.session_state:
            st.session_state.data_saved = False

        if st.button("Save Uploaded Data"):
            for row in df_upload.to_dict(orient="records"):

                row["teacher"] = st.session_state.teacher

                # attendance fix
                if "Attendance" in row:
                    row["status"] = str(row["Attendance"]).strip()
                else:
                    row["status"] = "Present"

                # max marks logic (default + custom support)
                if "Max Marks" in row and str(row["Max Marks"]).strip() != "":
                    try:
                        row["Max Marks"] = float(row["Max Marks"])
                    except:
                        row["Max Marks"] = max_marks_input
                else:
                    row["Max Marks"] = max_marks_input

                uploads.insert_one(row)

            st.session_state.data_saved = True

        if st.session_state.data_saved:
            st.success("Data saved successfully")
            st.session_state.data_saved = False


        # -------- EDIT MAX MARKS IN UI --------
    st.subheader("✏️ Edit Max Marks Of Students")

    # ✅ FIX: ensure teacher exists
    if "teacher" not in st.session_state:
        st.session_state.teacher = None

    data = []
    if st.session_state.teacher:
        data = list(students.find({"teacher": st.session_state.teacher})) + \
               list(uploads.find({"teacher": st.session_state.teacher}))

    if data:
        names = []
        for d in data:
            name = d.get("student") or d.get("Student Name", "Unknown")
            names.append(name)

        selected_student = st.selectbox("Select Student", names)

        new_max = st.number_input("New Max Marks", min_value=0, value=0)

        if st.button("Update Max Marks"):

            students.update_many(
                {"teacher": st.session_state.teacher, "student": selected_student},
                {"$set": {"Max Marks": new_max}}
            )

            uploads.update_many(
                {"teacher": st.session_state.teacher, "student": selected_student},
                {"$set": {"Max Marks": new_max}}
            )

            st.success(f"Updated Max Marks for {selected_student}")

    else:
        st.info("No students available to edit")


    # -------- STUDENT ENTRY --------
    st.subheader("Student Details")

    box = st.container()

    with box:

        # -------- NUMBER OF CO --------
        co_no = st.number_input("Number of CO", min_value=0, max_value=6, step=1)

        # -------- CO INPUT FIELDS --------
        co_marks_inputs = []
        for i in range(int(co_no)):
            co_marks_inputs.append(
                st.text_input(f"CO{i+1} Marks", key=f"co_{i}_{st.session_state.reset_key}")
            )

        # -------- FORM --------
        with st.form(f"student_form_{st.session_state.reset_key}"):

            subject = st.text_input("Subject Name")
            code = st.text_input("Subject Code")
            threshold = st.number_input("Threshold Marks", min_value=0, value=0)

            name = st.text_input("Student Name")
            status = st.selectbox("Attendance", ["Present", "Absent"])

            submit_student = st.form_submit_button("Add Student")

            if submit_student:

                if name.strip() == "" or subject.strip() == "" or code.strip() == "":
                    st.error("❌ Student Name, Subject Name and Subject Code are required")

                else:
                    marks_int = [int(x) if x.isdigit() else 0 for x in co_marks_inputs]
                    total = sum(marks_int)

                    if use_weight:
                        final_total = (total / max_marks_input) * weight_input if max_marks_input > 0 else 0
                    else:
                        final_total = total

                    students.insert_one({
                        "student": name,
                        "teacher": st.session_state.teacher,
                        "subject": subject,
                        "code": code,
                        "threshold": int(threshold),
                        "co_marks": marks_int,
                        "total": total,
                        "final_total": final_total,
                        "status": status
                    })

                    st.success("✅ Student Added Successfully")

                    time.sleep(1)
                    st.session_state.reset_key += 1
                    st.rerun()
                # -------- BUTTON ALIGNMENT FIX --------
    st.subheader("📥 Download My Data (Excel)")

    col1, col2 = st.columns(2)

    with col1:
        download_clicked = st.button("Download My Students Data")

    with col2:
        calculate_clicked = st.button("Calculate Attainment")

    # -------- DOWNLOAD LOGIC --------
    if download_clicked:

        student_data = list(students.find({"teacher": st.session_state.teacher}))
        upload_data = list(uploads.find({"teacher": st.session_state.teacher}))

        all_data = student_data + upload_data

        if not all_data:
            st.warning("⚠️ No data available to download")
        else:

            # -------- DETECT MAX CO --------
            max_co = 0
            for d in all_data:
                if "co_marks" in d:
                    max_co = max(max_co, len(d.get("co_marks", [])))
                else:
                    co_keys = [k for k in d.keys() if str(k).startswith("CO")]
                    max_co = max(max_co, len(co_keys))

            co_headers = [f"CO{i+1}" for i in range(max_co)]

            formatted_data = []

            for d in all_data:

                # CO extraction
                if "co_marks" in d:
                    co_marks = d.get("co_marks", [])
                else:
                    co_marks = [d.get(f"CO{i+1}", 0) for i in range(max_co)]

                obtained = d.get("total", 0)
                max_marks = d.get("Max Marks", max_marks_input)

                # -------- ATTAINMENT % --------
                attainment_percent = min((obtained / max_marks * 100), 100) if max_marks > 0 else 0

                # -------- ATTAINMENT LEVEL --------
                if attainment_percent >= 70:
                    attainment_level = 3
                elif attainment_percent >= 60:
                    attainment_level = 2
                elif attainment_percent >= 50:
                    attainment_level = 1
                else:
                    attainment_level = 0

                row = {
                    "Student Name": d.get("student", ""),
                    "Subject Name": d.get("subject", ""),
                    "Subject Code": d.get("code", ""),
                    "Attendance": d.get("status") or d.get("Attendance", "Present"),
                    "Max Marks": max_marks,
                    "Obtained Marks": obtained,
                    "Final Marks": d.get("final_total", 0),
                    "Threshold": d.get("threshold", 0),
                    "Attainment %": round(attainment_percent, 2),
                    "Attainment Level": attainment_level
                }

                # -------- ADD CO VALUES --------
                for i in range(max_co):
                    row[f"CO{i+1}"] = co_marks[i] if i < len(co_marks) else 0

                formatted_data.append(row)

            file_name = f"{st.session_state.teacher}_students_data.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.title = "Student Data"

            headers = [
                "Student Name", "Subject Name", "Subject Code",
                "Attendance", "Max Marks", "Obtained Marks", "Final Marks",
                "Threshold", "Attainment %", "Attainment Level"
            ] + co_headers

            # -------- HEADER STYLE --------
            for col_num, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_num, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")

            # -------- DATA --------
            for row_num, row_data in enumerate(formatted_data, start=2):
                for col_num, key in enumerate(headers, 1):
                    ws.cell(row=row_num, column=col_num, value=row_data.get(key, ""))

            # -------- AUTO WIDTH --------
            for col in ws.columns:
                max_length = 0
                col_letter = get_column_letter(col[0].column)

                for cell in col:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))

                ws.column_dimensions[col_letter].width = max_length + 3

            wb.save(file_name)

            with open(file_name, "rb") as f:
                st.download_button("⬇️ Download Excel File", f, file_name=file_name)

    # -------- ATTAINMENT --------
    if calculate_clicked:

        data_students = list(students.find({"teacher": st.session_state.teacher}))
        data_uploads = list(uploads.find({"teacher": st.session_state.teacher}))

        if len(data_students) == 0 and len(data_uploads) == 0:
            st.warning("⚠️ No students found in database")
            st.stop()

        threshold_val = int(threshold)

        df, summary = calculate_attainment(threshold_val, max_marks_input)

        st.write(summary)

        marks = df.get("final_total", df["Total Marks"])

        passed = (marks >= threshold_val).sum()
        failed = (marks < threshold_val).sum()

        fig, ax = plt.subplots()
        ax.bar(["Passed", "Failed"], [passed, failed])
        st.pyplot(fig)

# ---------------- FOOTER ----------------
show_footer()