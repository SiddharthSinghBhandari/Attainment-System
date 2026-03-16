import pandas as pd
from db import students
from openpyxl import load_workbook


def write_safe(ws, row, col, value):
    """
    Write to cell even if the position is inside a merged range.
    """
    cell = ws.cell(row=row, column=col)

    for merged in ws.merged_cells.ranges:
        if cell.coordinate in merged:
            row = merged.min_row
            col = merged.min_col
            break

    ws.cell(row=row, column=col).value = value


def calculate_attainment(threshold, max_marks):

    wb = load_workbook("template.xlsx")
    ws = wb.active

    data = list(students.find())

    start_row = 10
    rows = []

    for i, s in enumerate(data):

        row = start_row + i

        name = s.get("student", "")
        status = s.get("status", "Present")
        co = s.get("co_marks", [0, 0, 0])

        co1 = co[0] if len(co) > 0 else 0
        co2 = co[1] if len(co) > 1 else 0
        co3 = co[2] if len(co) > 2 else 0

        total = co1 + co2 + co3

        # write safely even if cells are merged
        write_safe(ws, row, 1, i + 1)     # Sr No
        write_safe(ws, row, 3, name)      # Name
        write_safe(ws, row, 5, total)     # Max Marks
        write_safe(ws, row, 6, co1)       # CO1
        write_safe(ws, row, 7, co2)       # CO2
        write_safe(ws, row, 8, co3)       # CO3

        rows.append({
            "Student": name,
            "Total Marks": total,
            "Status": status
        })

    wb.save("CO_Attainment.xlsx")

    df = pd.DataFrame(rows)

    summary = pd.DataFrame({
        "Parameter": ["Students"],
        "Value": [len(rows)]
    })

    return df, summary