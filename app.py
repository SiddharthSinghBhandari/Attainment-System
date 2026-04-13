import streamlit as st
import matplotlib.pyplot as plt
import pandas as pd
from db import students, users, uploads
from calculate import calculate_attainment
import os
import time   # ✅ ADDED
import io
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter


# ---------------- TEMPLATE BUILDER ----------------
def _build_template():
    """Generate the student data entry template as a BytesIO object."""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side, Protection

    NUM_STUDENTS = 30
    BLUE_HDR = "1E3A5F"
    ALT1     = "EBF5FB"
    ALT2     = "FFFFFF"

    def S(): return Side(style="thin")
    def bdr(): return Border(left=S(), right=S(), top=S(), bottom=S())

    def c(ws, r, col, value=None, bold=False, size=11, h="center", v="center",
          bg=None, locked=True, color="000000"):
        cl = ws.cell(row=r, column=col, value=value)
        cl.font       = Font(name="Arial", bold=bold, size=size, color=color)
        cl.alignment  = Alignment(horizontal=h, vertical=v)
        if bg: cl.fill = PatternFill("solid", fgColor=bg)
        cl.border     = bdr()
        cl.protection = Protection(locked=locked)
        return cl

    def m(ws, r1, c1, r2, c2):
        ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)

    wb = Workbook()
    ws = wb.active
    ws.title = "Student Data"

    for col, w in zip("ABCDEFG", [6, 14, 26, 8, 12, 12, 12]):
        ws.column_dimensions[col].width = w

    # Row 1: Headers
    ws.row_dimensions[1].height = 22
    for c1, c2, label in [
        (1,1,"Sr.No"),(2,2,"Reg No"),(3,4,"Student Name"),
        (5,5,"CO1 Marks"),(6,6,"CO2 Marks"),(7,7,"CO3 Marks")
    ]:
        if c1 != c2: m(ws, 1, c1, 1, c2)
        c(ws, 1, c1, label, bold=True, size=11, bg=BLUE_HDR, color="FFFFFF")

    # Student rows
    for idx in range(NUM_STUDENTS):
        r   = 2 + idx
        alt = ALT1 if idx % 2 == 0 else ALT2
        ws.row_dimensions[r].height = 17
        c(ws, r, 1, idx+1, size=10, bg=alt)
        c(ws, r, 2, None,  size=10, bg=ALT2, locked=False, h="center")
        m(ws, r, 3, r, 4)
        c(ws, r, 3, None,  size=10, bg=ALT2, locked=False, h="left")
        c(ws, r, 5, None,  size=10, bg=ALT2, locked=False)
        c(ws, r, 6, None,  size=10, bg=ALT2, locked=False)
        c(ws, r, 7, None,  size=10, bg=ALT2, locked=False)

    ws.freeze_panes = "A2"
    ws.protection.sheet = True
    ws.protection.password = ""
    ws.protection.enable()

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


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
    st.subheader("📥 Download Template")
    st.download_button(
        "⬇️ Download Student Data Template",
        data=_build_template(),
        file_name="Student_Data_Template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # -------- CLASS CONFIGURATION --------
    st.subheader("📊 Class Configuration")
    total_students_input = st.number_input("Total Students", min_value=1, value=1)
    max_marks_input = st.number_input("Maximum Marks (Total)", min_value=1, value=15)



    # -------- THRESHOLD --------
    st.subheader("🎯 Threshold & CO Marks Distribution")
    threshold_global = st.number_input(
        "Threshold % (students scoring above this are counted as passed)",
        min_value=0, max_value=100, value=50,
        help="e.g. 50 means students who score ≥50% in a CO are counted as attained"
    )

    # -------- CO MAX MARKS --------
    co_no_config = st.number_input("Number of COs", min_value=1, max_value=6, value=3, step=1)
    co_no_config = int(co_no_config)

    st.markdown("**Set max marks for each CO** (must add up to total max marks)")
    co_max_cols = st.columns(co_no_config)
    co_maxes_input = []
    for i in range(co_no_config):
        default_val = round(max_marks_input / co_no_config, 1)
        with co_max_cols[i]:
            val = st.number_input(
                f"CO{i+1} Max",
                min_value=0.0, max_value=float(max_marks_input),
                value=float(round(max_marks_input / co_no_config, 1)),
                step=0.5, key=f"co_max_{i}"
            )
            co_maxes_input.append(val)

    co_sum = sum(co_maxes_input)
    if abs(co_sum - max_marks_input) > 0.01:
        st.warning(f"⚠️ CO max marks sum = {co_sum}, but Total Max Marks = {max_marks_input}. Please adjust so they match.")
    else:
        st.success(f"✅ CO marks distribution: {' + '.join(str(v) for v in co_maxes_input)} = {co_sum}")

    # -------- COURSE / INSTITUTE DETAILS --------
    st.subheader("🏫 Course & Institute Details (for Excel Report)")
    col_a, col_b = st.columns(2)
    with col_a:
        university_name = st.text_input("University Name", value="CT UNIVERSITY")
        department_name = st.text_input("Department", value="Department of Management Studies")
        program_name    = st.text_input("Program (e.g. BBA)", value="BBA")
    with col_b:
        semester_val    = st.text_input("Semester", value="II")
        course_code_val = st.text_input("Course Code")
        course_name_val = st.text_input("Course Name")
    exam_label_val  = st.text_input("Exam Label", value="Mid Term Examination")

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
        upload_data  = list(uploads.find({"teacher": st.session_state.teacher}))

        all_data = student_data + upload_data

        if not all_data:
            st.warning("⚠️ No data available to download")
        else:

            # -------- NORMALISE EACH RECORD --------
            # Manual entries use: student, reg_no, co_marks[], total, subject, code, status, threshold
            # Uploaded entries use: Name, Reg No, CO1, CO2..., Attendance, Max Marks
            def normalise(d):
                name = (d.get("student") or d.get("Student Name") or
                        d.get("Name") or d.get("student_name") or "")
                reg_no = (d.get("reg_no") or d.get("Reg No") or
                          d.get("RegNo") or d.get("Registration No") or "")
                status = (d.get("status") or d.get("Attendance") or "Present")

                # CO marks — support "CO1", "CO1 Marks", co_marks[] formats
                if "co_marks" in d:
                    co_vals = [float(v or 0) for v in d["co_marks"]]
                else:
                    co_vals = []
                    i = 1
                    while True:
                        v = (d.get(f"CO{i}") if d.get(f"CO{i}") is not None else
                             d.get(f"CO{i} Marks") if d.get(f"CO{i} Marks") is not None else
                             d.get(f"co{i}") if d.get(f"co{i}") is not None else None)
                        if v is None:
                            break
                        co_vals.append(float(v or 0))
                        i += 1

                total = d.get("total") if d.get("total") else sum(float(v or 0) for v in co_vals)

                return {
                    "name":      name,
                    "reg_no":    reg_no,
                    "status":    status,
                    "subject":   d.get("subject") or d.get("Subject Name", ""),
                    "code":      d.get("code")    or d.get("Subject Code", ""),
                    "co_vals":   co_vals,
                    "total":     total,
                    "final_total": d.get("final_total", total),
                    "threshold": d.get("threshold", 0),
                    "max_marks": d.get("Max Marks", max_marks_input),
                }

            normalised = [normalise(d) for d in all_data]

            # -------- DETECT MAX CO --------
            max_co = max((len(n["co_vals"]) for n in normalised), default=0)
            co_headers = [f"CO{i+1}" for i in range(max_co)]

            formatted_data = []

            for n in normalised:
                co_vals = n["co_vals"]
                obtained   = n["total"]
                max_marks  = n["max_marks"]

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
                    "Student Name": n["name"],
                    "Reg No":       n["reg_no"],
                    "Subject Name": n["subject"],
                    "Subject Code": n["code"],
                    "Attendance":   n["status"],
                    "Max Marks":    max_marks,
                    "Obtained Marks": obtained,
                    "Final Marks":  n["final_total"],
                    "Threshold":    n["threshold"],
                    "Attainment %": round(attainment_percent, 2),
                    "Attainment Level": attainment_level,
                }

                # -------- ADD CO VALUES --------
                for i in range(max_co):
                    row[f"CO{i+1}"] = co_vals[i] if i < len(co_vals) else 0

                formatted_data.append(row)

            file_name = f"{st.session_state.teacher}_students_data.xlsx"

            wb = Workbook()
            ws = wb.active
            ws.title = "Student Data"

            headers = [
                "Student Name", "Reg No", "Subject Name", "Subject Code",
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

            excel_buffer = io.BytesIO()
            wb.save(excel_buffer)
            excel_buffer.seek(0)
            st.download_button("⬇️ Download Excel File", excel_buffer, file_name=file_name,
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # -------- ATTAINMENT --------
    if calculate_clicked:

        data_students = list(students.find({"teacher": st.session_state.teacher}))
        data_uploads = list(uploads.find({"teacher": st.session_state.teacher}))

        if len(data_students) == 0 and len(data_uploads) == 0:
            st.warning("⚠️ No students found in database")
            st.stop()

        threshold_val = int(threshold_global)

        df, summary, co_excel_buffer = calculate_attainment(
            threshold_val, max_marks_input,
            co_maxes=co_maxes_input,
            teacher=st.session_state.teacher,
            program=program_name,
            semester=semester_val,
            course_code=course_code_val,
            course_name=course_name_val,
            department=department_name,
            university=university_name,
            exam_label=exam_label_val
        )
        st.session_state["co_excel_buffer"] = co_excel_buffer

        # ── fix pass/fail: threshold_val is %, total marks is raw
        # convert raw total to % using max_marks_input for comparison
        df["Total %"] = df["Total Marks"].apply(
            lambda m: round(m / max_marks_input * 100, 2) if max_marks_input > 0 else 0
        )
        df["Status"] = df.apply(
            lambda row: "Absent" if row["Status"] == "Absent"
                        else ("Pass" if row["Total %"] >= threshold_val else "Fail"),
            axis=1
        )

        passed = (df["Status"] == "Pass").sum()
        failed = (df["Status"] == "Fail").sum()
        absent = (df["Status"] == "Absent").sum()
        total_present = passed + failed
        attainment_pct = round(passed / total_present * 100, 2) if total_present > 0 else 0



        # ── CHART 1: Pass/Fail bar chart ──────────────────────────────────────
        st.subheader("📊 Attainment Overview")
        col_c1, col_c2, col_c3, col_c4 = st.columns(4)
        col_c1.metric("✅ Passed",  passed)
        col_c2.metric("❌ Failed",  failed)
        col_c3.metric("🚫 Absent",  absent)
        col_c4.metric("🎯 Attainment", f"{attainment_pct}%")

        present_df = df[df["Status"] != "Absent"].copy()
        present_df = present_df.sort_values("Total %", ascending=False).reset_index(drop=True)
        scores     = present_df["Total %"].tolist()
        avg        = round(sum(scores) / len(scores), 1) if scores else 0

        fig, axes = plt.subplots(1, 2, figsize=(13, 5))
        fig.patch.set_facecolor("#1e293b")

        # ── LEFT: simple bar chart passed vs failed ────────────────────────────
        ax1 = axes[0]
        ax1.set_facecolor("#1e293b")
        bars = ax1.bar(["Passed", "Failed", "Absent"],
                       [passed, failed, absent],
                       color=["#22c55e", "#ef4444", "#94a3b8"],
                       width=0.5, edgecolor="none")
        for bar, val in zip(bars, [passed, failed, absent]):
            ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.3,
                     str(int(val)), ha="center", va="bottom",
                     color="white", fontsize=13, fontweight="bold")
        ax1.set_title("Pass / Fail / Absent", color="white", fontsize=12, fontweight="bold", pad=12)
        ax1.set_ylabel("Number of Students", color="#94a3b8", fontsize=10)
        ax1.tick_params(colors="white", labelsize=11)
        ax1.set_ylim(0, max(passed, failed, absent, 1) * 1.25)
        for spine in ax1.spines.values(): spine.set_visible(False)
        ax1.yaxis.set_visible(False)
        ax1.grid(False)

        # ── RIGHT: score distribution bars coloured by threshold ──────────────
        ax2 = axes[1]
        ax2.set_facecolor("#1e293b")
        bar_colors = ["#22c55e" if s >= threshold_val else "#ef4444" for s in scores]
        ax2.bar(range(len(scores)), scores, color=bar_colors,
                width=1.0, edgecolor="none", alpha=0.85)
        ax2.axhline(y=threshold_val, color="#facc15", linewidth=2,
                    linestyle="--", label=f"Threshold: {threshold_val}%", zorder=5)
        ax2.axhline(y=avg, color="#60a5fa", linewidth=2,
                    linestyle="-", label=f"Class Avg: {avg}%", zorder=5)
        ax2.set_title("Score Distribution (High → Low)", color="white", fontsize=12, fontweight="bold", pad=12)
        ax2.set_xlabel("Students ranked by score", color="#94a3b8", fontsize=10)
        ax2.set_ylabel("Score %", color="#94a3b8", fontsize=10)
        ax2.set_ylim(0, 110)
        ax2.tick_params(colors="#94a3b8", labelsize=9)
        ax2.set_xticks([])
        for spine in ax2.spines.values(): spine.set_visible(False)
        ax2.grid(axis="y", color="#374151", linewidth=0.6, alpha=0.5)
        leg = ax2.legend(facecolor="#0f172a", edgecolor="#334155",
                         labelcolor="white", fontsize=9, loc="upper right")

        # green/red legend patches
        import matplotlib.patches as mpatches
        pass_patch = mpatches.Patch(color="#22c55e", label=f"Above threshold ({passed})")
        fail_patch = mpatches.Patch(color="#ef4444", label=f"Below threshold ({failed})")
        ax2.legend(handles=[pass_patch, fail_patch,
                             plt.Line2D([0],[0], color="#facc15", linewidth=2, linestyle="--", label=f"Threshold: {threshold_val}%"),
                             plt.Line2D([0],[0], color="#60a5fa", linewidth=2, label=f"Class Avg: {avg}%")],
                   facecolor="#0f172a", edgecolor="#334155", labelcolor="white",
                   fontsize=8.5, loc="upper right")

        plt.tight_layout(pad=2)
        st.pyplot(fig)

        co_buf = st.session_state.get("co_excel_buffer")
        if co_buf:
            st.download_button(
                "⬇️ Download CO Attainment Report (Template Format)",
                co_buf,
                file_name="CO_Attainment.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

# ---------------- FOOTER ----------------
show_footer()