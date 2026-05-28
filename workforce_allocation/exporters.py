from pathlib import Path

import pandas as pd


def build_dept_attendance_pivot(attendance_today: pd.DataFrame) -> pd.DataFrame:
    dept_attendance_summary = (
        attendance_today
        .groupby(["dept", "attendance_status"])
        .agg(headcount=("emp_id", "nunique"))
        .reset_index()
    )
    dept_attendance_pivot = (
        dept_attendance_summary
        .pivot_table(
            index="dept",
            columns="attendance_status",
            values="headcount",
            fill_value=0,
        )
        .reset_index()
    )

    for col in ["Present", "Late", "Absent"]:
        if col not in dept_attendance_pivot.columns:
            dept_attendance_pivot[col] = 0

    dept_attendance_pivot["Total_Come"] = (
        dept_attendance_pivot["Present"] + dept_attendance_pivot["Late"]
    )
    return dept_attendance_pivot[
        ["dept", "Present", "Late", "Total_Come", "Absent"]
    ].sort_values("Total_Come", ascending=False)


def build_no_scan_depts(
    employee: pd.DataFrame,
    scan_with_dept: pd.DataFrame,
) -> pd.DataFrame:
    all_depts = pd.DataFrame({
        "dept": employee["dept"].dropna().astype(str).str.strip().unique()
    })
    scan_depts = pd.DataFrame({
        "dept": (
            scan_with_dept[scan_with_dept["Timestamp"].notna()]["dept"]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )
    })
    return all_depts[~all_depts["dept"].isin(scan_depts["dept"])].sort_values("dept")


def export_result(
    output_folder: Path,
    target_date: str,
    allocation_result: pd.DataFrame,
    gap_summary: pd.DataFrame,
    attendance_today: pd.DataFrame,
    manpower_plan: pd.DataFrame,
    employee: pd.DataFrame,
    scan_with_dept: pd.DataFrame,
) -> Path:
    output_folder.mkdir(parents=True, exist_ok=True)
    date_text = target_date.replace("-", "")
    output_file = output_folder / f"workforce_allocation_result_{date_text}.xlsx"

    dept_attendance_pivot = build_dept_attendance_pivot(attendance_today)
    no_scan_depts = build_no_scan_depts(employee, scan_with_dept)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        allocation_result.to_excel(writer, sheet_name="Allocation Result", index=False)
        gap_summary.to_excel(writer, sheet_name="Gap Summary", index=False)
        attendance_today.to_excel(writer, sheet_name="Attendance Status", index=False)
        manpower_plan.to_excel(writer, sheet_name="Manpower Plan Clean", index=False)
        dept_attendance_pivot.to_excel(
            writer,
            sheet_name="Dept Attendance Summary",
            index=False,
        )
        no_scan_depts.to_excel(writer, sheet_name="Dept No Scan", index=False)
        scan_with_dept.to_excel(writer, sheet_name="Timestamp With Dept", index=False)

    return output_file

