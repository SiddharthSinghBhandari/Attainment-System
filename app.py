import streamlit as st
st.set_page_config(layout="wide")
import matplotlib.pyplot as plt
import pandas as pd
from db import students, users, uploads
from calculate import calculate_attainment


# -------- LOAD CSS --------
def load_css():
    with open("styles.css") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css()


# -------- SESSION STATE --------
if "page" not in st.session_state:
    st.session_state.page = "home"

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False


# -------- NAVBAR --------
def navbar():

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        if st.button("Admin Register", use_container_width=True):
            st.session_state.page = "admin_register"
            st.rerun()

    with col2:
        if st.button("Admin Login", use_container_width=True):
            st.session_state.page = "admin_login"
            st.rerun()

    with col3:
        if st.button("Teacher Register", use_container_width=True):
            st.session_state.page = "teacher_register"
            st.rerun()

    with col4:
        if st.button("Teacher Login", use_container_width=True):
            st.session_state.page = "teacher_login"
            st.rerun()

navbar()


# -------- TITLE --------
st.markdown("<h1>CO-PO Attainment System</h1>", unsafe_allow_html=True)


# -------- LOGO --------
logo_url = "https://fimt-ggsipu.org/images/flogo2025-1.jpg"

col1, col2, col3 = st.columns([1,4,1])

with col2:
    st.image(logo_url, width=450)


# -------- ADMIN REGISTER --------
if st.session_state.page == "admin_register":

    st.header("Create Admin Account")

    username = st.text_input("Admin Username")
    password = st.text_input("Password", type="password")

    if st.button("Create Admin"):

        if users.find_one({"role": "admin"}):
            st.error("Admin already exists")

        elif username.strip()=="" or password.strip()=="":
            st.error("Fill all fields")

        else:

            users.insert_one({
                "username": username,
                "password": password,
                "role": "admin"
            })

            st.success("Admin account created")


# -------- ADMIN LOGIN --------
elif st.session_state.page == "admin_login":

    st.header("Admin Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):

        user = users.find_one({
            "username": username,
            "password": password,
            "role": "admin"
        })

        if user:
            st.session_state.logged_in = True
            st.session_state.page = "admin_panel"
            st.rerun()
        else:
            st.error("Invalid credentials")


# -------- ADMIN PANEL --------
elif st.session_state.page == "admin_panel" and st.session_state.logged_in:

    col1, col2 = st.columns([8,1])

    with col1:
        st.header("Admin Dashboard")

    with col2:
        if st.button("Logout"):
            st.session_state.logged_in = False
            st.session_state.page = "home"
            st.rerun()

    st.subheader("Registered Teachers")

    teachers = list(users.find({"role":"teacher"}))

    if teachers:

        for t in teachers:

            col1,col2 = st.columns([6,1])

            with col1:
                st.write(f"👨‍🏫 {t['username']}")

            with col2:
                if st.button("Delete", key=str(t["_id"])):

                    users.delete_one({"_id":t["_id"]})
                    st.success("Teacher deleted")
                    st.rerun()

    else:
        st.info("No teachers registered")


# -------- TEACHER REGISTER --------
elif st.session_state.page == "teacher_register":

    st.header("Create Teacher Account")

    username = st.text_input("Teacher Username")
    password = st.text_input("Password", type="password")

    if st.button("Register Teacher"):

        if username.strip()=="" or password.strip()=="":
            st.error("Fill all fields")

        else:

            users.insert_one({
                "username": username,
                "password": password,
                "role": "teacher"
            })

            st.success("Teacher registered successfully")


# -------- TEACHER LOGIN --------
elif st.session_state.page == "teacher_login":

    st.header("Teacher Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):

        user = users.find_one({
            "username": username,
            "password": password,
            "role": "teacher"
        })

        if user:

            st.session_state.logged_in = True
            st.session_state.teacher = username
            st.session_state.page = "teacher_panel"
            st.rerun()

        else:
            st.error("Invalid credentials")


# -------- TEACHER PANEL --------
elif st.session_state.page == "teacher_panel" and st.session_state.logged_in:

    col1, col2 = st.columns([8,1])

    with col1:
        st.header("Teacher Dashboard")

    with col2:
        if st.button("Logout"):
            st.session_state.logged_in = False
            st.session_state.page = "home"
            st.rerun()


    subject = st.text_input("Subject Name")
    code = st.text_input("Subject Code")

    total_students = st.number_input("Total Students", min_value=1)
    max_marks = st.number_input("Maximum Marks", min_value=1)
    threshold = st.number_input("Threshold Marks", min_value=0)

    co_no = st.number_input("Number of CO", min_value=1, max_value=6)

    st.subheader("Student Details")

    name = st.text_input("Student Name")
    status = st.selectbox("Attendance", ["Present","Absent"])

    marks=[]

    for i in range(int(co_no)):
        m = st.number_input(f"CO{i+1} Marks", min_value=0)
        marks.append(m)

    total=sum(marks)


    if st.button("Add Student"):

        if name.strip()=="" or subject.strip()=="" or code.strip()=="":
            st.error("Student Name, Subject Name and Subject Code are mandatory")

        else:

            students.insert_one({
                "student": name,
                "teacher": st.session_state.teacher,
                "subject": subject,
                "code": code,
                "total_students": total_students,
                "max_marks": max_marks,
                "threshold": threshold,
                "co_marks": marks,
                "total": total,
                "status": status
            })

            st.success("Student Added Successfully")

    if st.button("Calculate Attainment"):

        df, summary = calculate_attainment(threshold,max_marks)

        st.write(summary)

        fig,ax = plt.subplots()
        ax.hist(df["Total Marks"])

        st.pyplot(fig)

        with open("CO_Attainment.xlsx","rb") as f:

            st.download_button(
                "Download Excel",
                f,
                file_name="CO_Attainment.xlsx"
            )


    # -------- FILE UPLOAD --------

    st.divider()
    st.subheader("Upload Excel/CSV Student Data")

    uploaded_file = st.file_uploader(
        "Upload Excel or CSV",
        type=["xlsx","csv"]
    )

    if uploaded_file is not None:

        try:

            if uploaded_file.name.endswith(".csv"):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)

            st.write("Preview of Uploaded File")
            st.dataframe(df)

            if st.button("Save File Data to MongoDB"):

                data = df.to_dict(orient="records")

                for row in data:

                    row["teacher"] = st.session_state.teacher
                    row["file_name"] = uploaded_file.name

                    uploads.insert_one(row)

                st.success("All students saved separately in MongoDB")

        except Exception as e:
            st.error(f"Error reading file: {e}")
