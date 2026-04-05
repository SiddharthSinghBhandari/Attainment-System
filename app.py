import streamlit as st
import matplotlib.pyplot as plt
import pandas as pd
from db import students, users, uploads
from calculate import calculate_attainment
import os
import time   # ✅ ADDED


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

        if st.button("Save Uploaded Data"):
            for row in df_upload.to_dict(orient="records"):
                row["teacher"] = st.session_state.teacher
                uploads.insert_one(row)
            st.success("Data saved successfully")

    # -------- STUDENT ENTRY --------
    st.subheader("Student Details")

    with st.form(f"student_form_{st.session_state.reset_key}"):

        subject = st.text_input("Subject Name")
        code = st.text_input("Subject Code")
        threshold = st.number_input("Threshold Marks", min_value=0, value=0)

        co_no = st.number_input("Number of CO", min_value=0, max_value=6)

        name = st.text_input("Student Name")
        status = st.selectbox("Attendance", ["Present", "Absent"])

        marks = [st.text_input(f"CO{i+1} Marks") for i in range(int(co_no))]

        submit_student = st.form_submit_button("Add Student")

        if submit_student:

            if name.strip() == "" or subject.strip() == "" or code.strip() == "":
                st.error("❌ Student Name, Subject Name and Subject Code are required")

            else:
                marks_int = [int(x) if x.isdigit() else 0 for x in marks]
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
            formatted_data = []

            for d in all_data:
                co_marks = d.get("co_marks", [])

                row = {
                    "Student Name": d.get("student", ""),
                    "Subject Name": d.get("subject", ""),
                    "Subject Code": d.get("code", ""),
                    "Attendance": d.get("status", ""),
                    "Max Marks": max_marks_input,
                    "Obtained Marks": d.get("total", 0),
                    "Final Marks": d.get("final_total", 0),
                }

                for i, mark in enumerate(co_marks):
                    row[f"CO{i+1}"] = mark

                formatted_data.append(row)

            df_download = pd.DataFrame(formatted_data)

            file_name = f"{st.session_state.teacher}_students_data.xlsx"

            with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
                df_download.to_excel(writer, index=False)

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