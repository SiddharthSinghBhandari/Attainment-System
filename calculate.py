import pandas as pd
from db import students

def calculate_attainment(subject, code, threshold, max_marks):

    data = list(students.find({
        "subject": subject,
        "code": code
    }))

    if len(data) == 0:
        return pd.DataFrame(), {}

    df = pd.DataFrame(data)

    df.columns = df.columns.str.strip()

    # keep only present students
    df = df[df["status"] == "Present"]

    # ----------- CREATE TOTAL HERE -----------
    if "co_marks" in df.columns:
        df["total"] = df["co_marks"].apply(
            lambda x: sum(x) if isinstance(x, list) else 0
        )

    df["total"] = pd.to_numeric(df["total"], errors="coerce")

    # ---------- ATTAINMENT CALC ----------
    above = len(df[df["total"] >= threshold])
    below = len(df[df["total"] < threshold])
    total_present = len(df)

    percent = (above / total_present) * 100 if total_present > 0 else 0

    if percent >= 70:
        level = 3
    elif percent >= 60:
        level = 2
    else:
        level = 1

    summary = {
        "Students Above Threshold": above,
        "Students Below Threshold": below,
        "% Students Above Threshold": percent,
        "CO Attainment Level": level
    }

    return df, summary