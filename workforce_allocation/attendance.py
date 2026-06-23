import datetime

import pandas as pd

from .utils import clean_emp_id, clean_name


def _get_night_shift_windows(manpower_plan: pd.DataFrame) -> list[tuple[str, str]]:
    """Returns (shift_start, shift_end) pairs for shifts that cross midnight."""
    shifts = manpower_plan.dropna(subset=["shift_start", "shift_end"]).copy()
    shifts["start_dt"] = pd.to_datetime(shifts["shift_start"], format="%H:%M", errors="coerce")
    shifts["end_dt"] = pd.to_datetime(shifts["shift_end"], format="%H:%M", errors="coerce")
    shifts = shifts.dropna(subset=["start_dt", "end_dt"])
    night = shifts[shifts["end_dt"] < shifts["start_dt"]]
    return list(
        night[["shift_start", "shift_end"]].drop_duplicates().itertuples(index=False, name=None)
    )


def _adjust_dates_for_night_shifts(
    scan: pd.DataFrame,
    manpower_plan: pd.DataFrame,
) -> pd.DataFrame:
    """
    Re-date early-morning clock-out timestamps to the previous day so they
    group with the night shift's clock-in.

    A scan at time T on date D is moved to D-1 when:
      - T is within 2 h after a night shift's shift_end (early morning window)
      - The same employee has a scan within 2 h before shift_start on D-1
        (confirming they actually worked that night shift)

    Each scan row is only reassigned once (first matching night shift wins).
    """
    night_shifts = _get_night_shift_windows(manpower_plan)
    if not night_shifts:
        return scan

    scan = scan.copy()
    scan["time_dt"] = pd.to_datetime(scan["time"], format="%H:%M", errors="coerce")
    already_reassigned = pd.Series(False, index=scan.index)

    for shift_start_str, shift_end_str in night_shifts:
        shift_start_dt = pd.to_datetime(shift_start_str, format="%H:%M")
        shift_end_dt = pd.to_datetime(shift_end_str, format="%H:%M")
        buffer_before = pd.Timedelta(hours=2)
        buffer_after = pd.Timedelta(hours=1)

        # (emp_id, D+1) pairs where D is a date the employee has an evening scan
        # — their D+1 early-morning scan is the clock-out for this night shift.
        # Upper bound uses a tighter window (+1h) to reduce false positives from
        # afternoon-shift overtime workers who stay past shift_start.
        evening_mask = (
            (scan["time_dt"] >= (shift_start_dt - buffer_before))
            & (scan["time_dt"] <= (shift_start_dt + buffer_after))
        )
        evening_df = scan.loc[evening_mask, ["emp_id", "date"]].copy()
        evening_df["clockout_date"] = evening_df["date"].apply(
            lambda d: d + datetime.timedelta(days=1)
        )
        confirmed_pairs: set[tuple] = set(
            zip(evening_df["emp_id"], evening_df["clockout_date"])
        )

        # Early-morning scans eligible for reassignment
        morning_mask = (
            (scan["time_dt"] <= shift_end_dt + buffer)
            & (scan["time_dt"].dt.hour < 12)
            & ~already_reassigned
        )
        scan_keys = list(zip(scan["emp_id"], scan["date"]))
        in_confirmed = pd.Series(
            [k in confirmed_pairs for k in scan_keys],
            index=scan.index,
        )

        reassign_this = morning_mask & in_confirmed
        already_reassigned |= reassign_this

    scan.loc[already_reassigned, "date"] = scan.loc[already_reassigned, "date"].apply(
        lambda d: d - datetime.timedelta(days=1)
    )
    return scan.drop(columns=["time_dt"])


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
        scan = _adjust_dates_for_night_shifts(scan, manpower_plan)

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

