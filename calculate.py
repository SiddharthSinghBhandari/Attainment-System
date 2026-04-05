import pandas as pd
from db import students, uploads
from openpyxl import load_workbook


# -------- SAFE WRITE FUNCTION --------
def write_safe(ws, row, col, value):
    """
    Write value safely into Excel even if cell is merged
    """
    cell = ws.cell(row=row, column=col)

    for merged in ws.merged_cells.ranges:
        if cell.coordinate in merged:
            row = merged.min_row
            col = merged.min_col
            break

    ws.cell(row=row, column=col).value = value


# -------- MAIN ATTAINMENT FUNCTION --------
def calculate_attainment(threshold, max_marks):
    """
    Calculate CO attainment and generate Excel output
    """

    wb = load_workbook("template.xlsx")
    ws = wb.active

    # Fetch data from both collections
    data_students = list(students.find())
    data_uploads = list(uploads.find())

    data = data_students + data_uploads

    start_row = 10
    rows = []

    passed_count = 0

    for i, s in enumerate(data):

        row = start_row + i

        # Handle different key formats (manual + uploaded)
        name = s.get("student") or s.get("Name", "")
        reg_no = s.get("reg_no") or s.get("Reg No", "")

        # Handle CO marks from both sources
        if "co_marks" in s:
            co = s.get("co_marks", [0, 0, 0])
            co1 = float(co[0]) if len(co) > 0 else 0
            co2 = float(co[1]) if len(co) > 1 else 0
            co3 = float(co[2]) if len(co) > 2 else 0
        else:
            co1 = float(s.get("CO1", 0))
            co2 = float(s.get("CO2", 0))
            co3 = float(s.get("CO3", 0))

        # Total marks
        total = co1 + co2 + co3

        # Percentage calculation
        co1_per = (co1 / 4.5 * 100) if co1 else 0
        co2_per = (co2 / 4.5 * 100) if co2 else 0
        co3_per = (co3 / 6 * 100) if co3 else 0

        # Pass / Fail logic
        result = "Pass" if total >= threshold else "Fail"
        if result == "Pass":
            passed_count += 1

        # Write into Excel template
        write_safe(ws, row, 1, i + 1)
        write_safe(ws, row, 2, reg_no)
        write_safe(ws, row, 3, name)

        write_safe(ws, row, 4, total)
        write_safe(ws, row, 5, co1)
        write_safe(ws, row, 6, co2)
        write_safe(ws, row, 7, co3)

        write_safe(ws, row, 8, round(co1_per, 2))
        write_safe(ws, row, 9, round(co2_per, 2))
        write_safe(ws, row, 10, round(co3_per, 2))

        rows.append({
            "Student": name,
            "Total Marks": total,
            "Status": result
        })

    total_students = len(rows)

    # Calculate attainment percentage
    attainment = (passed_count / total_students * 100) if total_students > 0 else 0

    # Write summary into Excel
    write_safe(ws, 21, 6, passed_count)
    write_safe(ws, 22, 6, total_students)
    write_safe(ws, 23, 6, round(attainment, 2))

    wb.save("CO_Attainment.xlsx")

    df = pd.DataFrame(rows)

    summary = pd.DataFrame({
        "Parameter": ["Total Students", "Passed", "Attainment %"],
        "Value": [total_students, passed_count, round(attainment, 2)]
    })

    return df, summary