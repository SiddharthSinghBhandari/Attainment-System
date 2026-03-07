import streamlit as st
import matplotlib.pyplot as plt
from db import students, users
from calculate import calculate_attainment


# -------- LOAD CSS --------
def load_css():
    with open("styles.css") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css()


# -------- SESSION STATE --------
if "page" not in st.session_state:
    st.session_state.page = "home"


# -------- NAVBAR --------
def navbar():

    col1, col2, col3, col4 = st.columns(4, gap="large")

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

    st.markdown('</div>', unsafe_allow_html=True)


navbar()

# -------- TITLE --------
st.markdown("<h1>CO-PO Attainment System</h1>", unsafe_allow_html=True)


# -------- CENTER LOGO IMAGE -------- asdasdasd

# -------- CENTER LOGO IMAGE --------

logo_url = "https://fimt-ggsipu.org/images/flogo2025-1.jpg"

col1, col2, col3 = st.columns([1,4,1])

with col2:
    st.image(logo_url, width=450)

st.markdown('</div>', unsafe_allow_html=True)


# -------- ADMIN REGISTER --------
if st.session_state.page == "admin_register":

    st.header("Create Admin Account")

    username = st.text_input("Admin Username")
    password = st.text_input("Password", type="password")

    if st.button("Create Admin"):

        if users.find_one({"role": "admin"}):
            st.error("Admin already exists")

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
            st.session_state.page = "admin_panel"
            st.rerun()
        else:
            st.error("Invalid credentials")


# -------- TEACHER REGISTER --------
elif st.session_state.page == "teacher_register":

    st.header("Create Teacher Account")

    username = st.text_input("Teacher Username")
    password = st.text_input("Password", type="password")

    if st.button("Register Teacher"):

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
            st.session_state.page = "teacher_panel"
            st.rerun()
        else:
            st.error("Invalid credentials")


# -------- ADMIN PANEL --------
elif st.session_state.page == "admin_panel":

    st.header("Admin Dashboard")

    st.subheader("Create Teacher")

    t_user = st.text_input("Teacher Username", key="add_teacher")
    t_pass = st.text_input("Teacher Password", type="password")

    if st.button("Add Teacher"):

        users.insert_one({
            "username": t_user,
            "password": t_pass,
            "role": "teacher"
        })

        st.success("Teacher Added")
        st.rerun()

    st.subheader("Teacher List")

    teacher_list = list(users.find({"role": "teacher"}))

    if len(teacher_list) == 0:
        st.info("No teachers available")

    else:

        for teacher in teacher_list:

            col1, col2 = st.columns([3,1])

            with col1:
                st.write(teacher["username"])

            with col2:

                if st.button("Delete", key=str(teacher["_id"])):

                    users.delete_one({"_id": teacher["_id"]})
                    st.success("Teacher Removed")
                    st.rerun()


# -------- TEACHER PANEL --------
elif st.session_state.page == "teacher_panel":

    st.header("Teacher Dashboard")

    subject = st.text_input("Subject Name")
    code = st.text_input("Subject Code")

    total_students = st.number_input("Total Students", min_value=1)
    max_marks = st.number_input("Maximum Marks", min_value=1)
    threshold = st.number_input("Threshold Marks", min_value=0)

    co_no = st.number_input("Number of CO", min_value=1, max_value=6)

    st.subheader("Student Details")

    name = st.text_input("Student Name")
    status = st.selectbox("Attendance", ["Present", "Absent"])

    marks = []

    for i in range(int(co_no)):
        m = st.number_input(f"CO{i+1} Marks", min_value=0)
        marks.append(m)

    total = sum(marks)

    if st.button("Add Student"):

        students.insert_one({
            "student": name,
            "subject": subject,
            "code": code,
            "total_students": total_students,
            "max_marks": max_marks,
            "threshold": threshold,
            "co_marks": marks,
            "total": total,
            "status": status
        })

        st.success("Student Added")

    if st.button("Calculate Attainment"):

        df, summary = calculate_attainment(threshold, max_marks)

        st.write(summary)

        fig, ax = plt.subplots()
        ax.hist(df["Total Marks"])
        st.pyplot(fig)

        with open("CO_Attainment.xlsx", "rb") as f:
            st.download_button(
                "Download Excel",
                f,
                file_name="CO_Attainment.xlsx"
            )