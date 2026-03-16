import pandas as pd
from db import students


def calculate_attainment(threshold, max_marks):

    data = list(students.find())

    present = 0
    absent = 0
    above = 0
    below = 0

    rows = []

    for s in data:

        name = s.get("student", "")
        total = s.get("total", 0)
        status = s.get("status", "Present")

        if status == "Absent":
            absent += 1
        else:
            present += 1

            if total >= threshold:
                above += 1
            else:
                below += 1

        rows.append({
            "Student": name,
            "Total Marks": total,
            "Status": status
        })


    df_students = pd.DataFrame(rows)

    percent = (above/present)*100 if present else 0

    if percent >= 70:
        attainment = 3
    elif percent >= 60:
        attainment = 2
    elif percent >= 50:
        attainment = 1
    else:
        attainment = 0


    summary = pd.DataFrame({
        "Parameter":[
            "Present Students",
            "Absent Students",
            "Students Above Threshold",
            "Students Below Threshold",
            "% Students Above Threshold",
            "CO Attainment Level"
        ],
        "Value":[
            present,
            absent,
            above,
            below,
            round(percent,2),
            attainment
        ]
    })


    with pd.ExcelWriter("CO_Attainment.xlsx") as writer:
        df_students.to_excel(writer, sheet_name="Student Data", index=False)
        summary.to_excel(writer, sheet_name="Attainment Summary", index=False)


    return df_students, summary