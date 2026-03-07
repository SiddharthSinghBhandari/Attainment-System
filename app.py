import streamlit as st
import matplotlib.pyplot as plt
from db import students, users
from calculate import calculate_attainment

# Load CSS
def load_css():
    with open("styles.css") as f:
        st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

load_css()

if "page" not in st.session_state:
    st.session_state.page="home"


# HOME
if st.session_state.page == "home":

    st.markdown("<h1>CO-PO Attainment System</h1>", unsafe_allow_html=True)

    if st.button("Admin Register"):
        st.session_state.page = "admin_register"
        st.rerun()

    if st.button("Admin Login"):
        st.session_state.page = "admin_login"
        st.rerun()

    if st.button("Teacher Login"):
        st.session_state.page = "teacher_login"
        st.rerun()


# ADMIN REGISTER
elif st.session_state.page=="admin_register":

    st.header("Create Admin Account")

    username=st.text_input("Admin Username",key="admin_reg_user")
    password=st.text_input("Password",type="password",key="admin_reg_pass")

    if st.button("Create Admin"):

        if users.find_one({"role":"admin"}):
            st.error("Admin already exists")

        else:

            users.insert_one({
                "username":username,
                "password":password,
                "role":"admin"
            })

            st.success("Admin created")

    if st.button("Back"):
        st.session_state.page="home"
        st.rerun()


# ADMIN LOGIN
elif st.session_state.page=="admin_login":

    st.header("Admin Login")

    username=st.text_input("Username",key="admin_login_user")
    password=st.text_input("Password",type="password",key="admin_login_pass")

    if st.button("Login"):

        user=users.find_one({
            "username":username,
            "password":password,
            "role":"admin"
        })

        if user:
            st.session_state.page="admin_panel"
            st.rerun()
        else:
            st.error("Invalid login")

    if st.button("Back"):
        st.session_state.page="home"
        st.rerun()


# TEACHER LOGIN
elif st.session_state.page=="teacher_login":

    st.header("Teacher Login")

    username=st.text_input("Username",key="teacher_login_user")
    password=st.text_input("Password",type="password",key="teacher_login_pass")

    if st.button("Login"):

        user=users.find_one({
            "username":username,
            "password":password,
            "role":"teacher"
        })

        if user:
            st.session_state.page="teacher_panel"
            st.rerun()
        else:
            st.error("Invalid login")

    if st.button("Back"):
        st.session_state.page="home"
        st.rerun()


# ADMIN PANEL
elif st.session_state.page == "admin_panel":

    st.header("Admin Dashboard")

    # ---------------- ADD TEACHER ----------------

    st.subheader("Create Teacher")

    t_user = st.text_input("Teacher Username", key="add_teacher_user")
    t_pass = st.text_input("Teacher Password", type="password", key="add_teacher_pass")

    if st.button("Add Teacher"):

        if t_user == "" or t_pass == "":
            st.warning("Please enter username and password")

        else:

            # check if teacher already exists
            existing = users.find_one({"username": t_user, "role": "teacher"})

            if existing:
                st.error("Teacher already exists")

            else:
                users.insert_one({
                    "username": t_user,
                    "password": t_pass,
                    "role": "teacher"
                })

                st.success("Teacher Added")
                st.rerun()


    # ---------------- TEACHER LIST ----------------

    st.subheader("Teacher List")

    teacher_list = list(users.find({"role": "teacher"}))

    if len(teacher_list) == 0:
        st.info("No teachers added yet")

    else:

        for teacher in teacher_list:

            col1, col2 = st.columns([3,1])

            with col1:
                st.write(teacher["username"])

            with col2:

                if st.button("Delete", key=teacher["_id"]):

                    users.delete_one({"_id": teacher["_id"]})
                    st.success("Teacher Removed")
                    st.rerun()


    # ---------------- LOGOUT ----------------

    if st.button("Logout"):
        st.session_state.page = "home"
        st.rerun()


# TEACHER PANEL
elif st.session_state.page=="teacher_panel":

    st.header("Teacher Dashboard")

    subject=st.text_input("Subject Name")
    code=st.text_input("Subject Code")

    total_students=st.number_input("Total Students",min_value=1)
    max_marks=st.number_input("Maximum Marks",min_value=1)
    threshold=st.number_input("Threshold Marks",min_value=0)

    co_no=st.number_input("Number of CO",min_value=1,max_value=6)

    name=st.text_input("Student Name")
    status=st.selectbox("Attendance",["Present","Absent"])

    marks=[]

    for i in range(int(co_no)):
        m=st.number_input(f"CO{i+1} Marks",min_value=0)
        marks.append(m)

    total=sum(marks)

    if st.button("Add Student"):

        students.insert_one({
            "student":name,
            "subject":subject,
            "code":code,
            "total_students":total_students,
            "max_marks":max_marks,
            "threshold":threshold,
            "co_marks":marks,
            "total":total,
            "status":status
        })

        st.success("Student Added")

    if st.button("Calculate Attainment"):

        df,summary=calculate_attainment(threshold,max_marks)

        st.write(summary)

        fig,ax=plt.subplots()
        ax.hist(df["Total Marks"])
        st.pyplot(fig)

        with open("CO_Attainment.xlsx","rb") as f:
            st.download_button("Download Excel",f,file_name="CO_Attainment.xlsx")

    if st.button("Logout"):
        st.session_state.page="home"
        st.rerun()