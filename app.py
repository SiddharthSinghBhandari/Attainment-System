import streamlit as st
import matplotlib.pyplot as plt
from db import students
from calculate import calculate_attainment

st.title("CO-PO ttainment System")



st.subheader("Subject Details")

subject = st.text_input("Enter Subject Name")
code = st.text_input("Enter Subject Code")

total_students = st.number_input("Total number of Students", min_value=1)
max_marks = st.number_input("Maximum Marks", min_value=1)
threshold = st.number_input("Threshold Marks", min_value=0)

co_no = st.number_input("How many CO (Course Outcome) do you want?", min_value=1, max_value=6)

method = st.radio(
    "Do you want CO separately or weightage according to marks?",
    ("Separate CO Input", "Weightage by Marks")
)

# ---------------- STUDENT INPUT ----------------

st.subheader("Student Details")

name = st.text_input("Student Name")

status = st.selectbox("Attendance", ["Present","Absent"])

marks = []

for i in range(int(co_no)):
    m = st.number_input(f"Enter CO{i+1} Marks", min_value=0)
    marks.append(m)

total = sum(marks)

# ---------------- STORE DATA ----------------

if st.button("Submit Student Data"):

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

    st.success("Student Data Stored Successfully!")

if st.button("Calculate Attainment"):

    df, summary = calculate_attainment(threshold, max_marks)

    st.write(df)
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