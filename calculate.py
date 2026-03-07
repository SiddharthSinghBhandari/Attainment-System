import pandas as pd
from db import students

def calculate_attainment(threshold, max_marks):

    data = list(students.find())

    present = 0
    absent = 0
    above = 0
    below = 0

    result = []

    for i in data:

        status = i.get("status","Present")

        if status == "Absent":
            absent += 1
            continue

        present += 1

        total = i.get("total",0)

        if total >= threshold:
            above += 1
        else:
            below += 1

        result.append({
            "Student": i.get("student"),
            "Total Marks": total,
            "Status": status
        })

    percent = (above/present)*100 if present>0 else 0

    if percent >= 70:
        attainment = 3
    elif percent >= 60:
        attainment = 2
    elif percent >= 50:
        attainment = 1
    else:
        attainment = 0

    df = pd.DataFrame(result)

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
            percent,
            attainment
        ]
    })

    with pd.ExcelWriter("CO_Attainment.xlsx", engine="openpyxl") as writer:

        df.to_excel(writer, sheet_name="Student Data", index=False)
        summary.to_excel(writer, sheet_name="Attainment Summary", index=False)

    return df, summary