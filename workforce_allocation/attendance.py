import datetime

import pandas as pd

from .utils import clean_emp_id, clean_name


def _redate_am_scans_for_pm_shift_workers(
    scan: pd.DataFrame,
    employee: pd.DataFrame,
    manpower_plan: pd.DataFrame,
) -> pd.DataFrame:
    """
    For employees whose dept has only PM shifts (all shift_start ≥ 12:00),
    move AM timestamps back one day instead of discarding them.

    Night-shift workers (e.g., 22:00–06:00) produce early-morning exit scans
    (e.g., 02:00 on date D). If kept on date D, scan_in = min(time) would pick
    02:00 as clock-in, making the worker look very late for their PM shift.
    By moving those scans to D-1, they become scan_out for the night shift on
    D-1 and are no longer available as clock-in on D.
    """
    dept_shifts = (
        manpower_plan[["dept", "shift_start"]]
        .dropna(subset=["shift_start"])
        .copy()
    )
    dept_shifts["_start_dt"] = pd.to_datetime(
        dept_shifts["shift_start"], format="%H:%M", errors="coerce"
    )
    dept_shifts = dept_shifts.dropna(subset=["_start_dt"])
    dept_shifts["_is_pm"] = dept_shifts["_start_dt"].dt.hour >= 12

    dept_all_pm = dept_shifts.groupby("dept")["_is_pm"].all()
    pm_depts = set(dept_all_pm[dept_all_pm].index)
    if not pm_depts:
        return scan

    pm_emp_ids = set(
        employee.loc[
            employee["dept"].astype(str).str.strip().isin(pm_depts), "emp_id"
        ].dropna()
    )
    if not pm_emp_ids:
        return scan

    scan = scan.copy()
    scan["_time_dt"] = pd.to_datetime(scan["time"], format="%H:%M", errors="coerce")
    is_pm_worker = scan["emp_id"].isin(pm_emp_ids)
    is_am_scan = scan["_time_dt"].dt.hour < 12
    redate_mask = is_pm_worker & is_am_scan

    scan.loc[redate_mask, "date"] = scan.loc[redate_mask, "date"].apply(
        lambda d: d - datetime.timedelta(days=1)
    )
    return scan.drop(columns=["_time_dt"])


def rebuild_attendance(
    scan: pd.DataFrame,
    employee: pd.DataFrame,
    manpower_plan: pd.DataFrame | None = None,
) -> pd.DataFrame:
    scan = scan.copy()
    scan["emp_id"] = scan["Employee ID"].apply(clean_emp_id)
    scan["name_key"] = scan["Employee Name"].apply(clean_name)

    test_merge = scan.merge(
        employee[["emp_id", "name", "dept", "section", "position", "name_key"]],
        on="emp_id",
        how="left",
        suffixes=("", "_master"),
    )
    name_lookup = (
        employee[["emp_id", "name", "dept", "section", "position", "name_key"]]
        .dropna(subset=["name_key"])
        .drop_duplicates("name_key")
    )
    name_to_emp = dict(zip(name_lookup["name_key"], name_lookup["emp_id"]))
    condition_missing = test_merge["dept"].isna()

    scan.loc[condition_missing, "emp_id_by_name"] = scan.loc[
        condition_missing,
        "name_key",
    ].map(name_to_emp)
    condition_exact_name = condition_missing & scan["emp_id_by_name"].notna()
    scan.loc[condition_exact_name, "emp_id"] = scan.loc[
        condition_exact_name,
        "emp_id_by_name",
    ]

    if manpower_plan is not None:
        scan = _redate_am_scans_for_pm_shift_workers(scan, employee, manpower_plan)

    attendance = (
        scan.groupby(["emp_id", "Employee Name", "date"])
        .agg(scan_in=("time", "min"), scan_out=("time", "max"))
        .reset_index()
        .rename(columns={"Employee Name": "scan_name"})
    )
    return attendance


def add_department_to_scan(scan: pd.DataFrame, employee: pd.DataFrame) -> pd.DataFrame:
    employee_dept = employee[["emp_id", "name", "dept", "section", "position"]].copy()
    scan_with_dept = scan.merge(employee_dept, on="emp_id", how="left")
    return scan_with_dept[
        [
            "Timestamp",
            "emp_id",
            "Employee Name",
            "name",
            "dept",
            "section",
            "position",
            "date",
            "time",
        ]
    ].copy()


def build_attendance_today(
    employee: pd.DataFrame,
    attendance: pd.DataFrame,
    target_date: str,
) -> pd.DataFrame:
    target_date_value = pd.to_datetime(target_date).date()
    attendance_scan_today = attendance[
        attendance["date"] == target_date_value
    ].copy()

    attendance_today = employee.merge(
        attendance_scan_today[["emp_id", "scan_name", "date", "scan_in", "scan_out"]],
        on="emp_id",
        how="left",
    )
    attendance_today["date"] = attendance_today["date"].fillna(target_date_value)
    return attendance_today


def detect_shift_and_availability(
    attendance_today: pd.DataFrame,
    manpower_plan: pd.DataFrame,
    late_grace_minutes: int,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    attendance_today = attendance_today.copy()
    dept_shift_map = (
        manpower_plan[["dept", "shift", "shift_start"]]
        .dropna()
        .drop_duplicates()
        .copy()
    )
    dept_shift_map["dept_key"] = dept_shift_map["dept"].astype(str).str.strip()
    dept_shift_map["shift_start_dt"] = pd.to_datetime(
        dept_shift_map["shift_start"],
        format="%H:%M",
        errors="coerce",
    )
    dept_shift_map = dept_shift_map.dropna(subset=["shift_start_dt"])

    def detect_shift_by_dept(row):
        if pd.isna(row["scan_in"]):
            return pd.Series({
                "shift": None,
                "shift_start": None,
                "shift_detect_reason": "Absent",
            })

        scan_time = pd.to_datetime(row["scan_in"], format="%H:%M", errors="coerce")
        if pd.isna(scan_time):
            return pd.Series({
                "shift": None,
                "shift_start": None,
                "shift_detect_reason": "Invalid scan_in",
            })

        emp_dept = str(row["dept"]).strip()
        possible_shifts = dept_shift_map[
            dept_shift_map["dept_key"] == emp_dept
        ].copy()
        if possible_shifts.empty:
            return pd.Series({
                "shift": None,
                "shift_start": None,
                "shift_detect_reason": "No shift found for dept",
            })

        possible_shifts["diff_minutes"] = (
            possible_shifts["shift_start_dt"] - scan_time
        ).abs().dt.total_seconds() / 60
        closest = possible_shifts.sort_values("diff_minutes").iloc[0]
        return pd.Series({
            "shift": closest["shift"],
            "shift_start": closest["shift_start"],
            "shift_detect_reason": "Matched by dept closest time",
        })

    def check_attendance_status(row):
        if pd.isna(row["scan_in"]):
            return "Absent"
        if pd.isna(row["shift_start"]):
            return "Unknown Shift"

        scan_in = pd.to_datetime(row["scan_in"], format="%H:%M", errors="coerce")
        shift_start = pd.to_datetime(
            row["shift_start"],
            format="%H:%M",
            errors="coerce",
        )
        if pd.isna(scan_in) or pd.isna(shift_start):
            return "Invalid Time"

        late_time = shift_start + pd.Timedelta(minutes=late_grace_minutes)
        return "Late" if scan_in > late_time else "Present"

    attendance_today[["shift", "shift_start", "shift_detect_reason"]] = (
        attendance_today.apply(detect_shift_by_dept, axis=1)
    )
    attendance_today["attendance_status"] = attendance_today.apply(
        check_attendance_status,
        axis=1,
    )
    available = attendance_today[
        attendance_today["attendance_status"].isin(["Present", "Late"])
    ].copy()
    return attendance_today, available


def summarize_attendance_by_dept(attendance_today: pd.DataFrame) -> pd.DataFrame:
    return (
        attendance_today
        .groupby(["dept", "attendance_status"])
        .agg(headcount=("emp_id", "nunique"))
        .reset_index()
    )

