import pandas as pd
from db import students

def calculate_attainment(subject, code, threshold, max_marks):

    data = list(students.find({
        "subject": subject,
        "code": code
    }))

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

    # ----------- CREATE EXCEL WITH HISTOGRAM -----------

    with pd.ExcelWriter("CO_Attainment.xlsx", engine='xlsxwriter') as writer:

        df.to_excel(writer, sheet_name="Student Data", index=False)
        summary.to_excel(writer, sheet_name="Attainment Summary", index=False)

        # ----------- ATTAINMENT CRITERIA SHEET -----------

        criteria = pd.DataFrame({
            "Percentage of Students Scoring Above Threshold":[
                ">= 70%",
                ">= 60% and < 70%",
                ">= 50% and < 60%",
                "< 50%"
            ],
            "CO Attainment Level":[
                "Level 3",
                "Level 2",
                "Level 1",
                "Level 0"
            ]
        })

        criteria.to_excel(writer, sheet_name="Attainment Criteria", index=False)

        # ----------- HISTOGRAM -----------

        workbook  = writer.book
        worksheet = writer.sheets["Student Data"]

        chart = workbook.add_chart({'type': 'column'})

        chart.add_series({
            'name': 'Marks Distribution',
            'categories': f'=Student Data!$B$2:$B${len(df)+1}',
            'values':     f'=Student Data!$B$2:$B${len(df)+1}',
        })

        chart.set_title({'name': 'Marks Distribution'})
        chart.set_x_axis({'name': 'Total Marks'})
        chart.set_y_axis({'name': 'Number of Students'})

        worksheet.insert_chart('E2', chart)

    return df, summary