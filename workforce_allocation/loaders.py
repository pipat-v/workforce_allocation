from pathlib import Path

import pandas as pd

from .utils import clean_emp_id, clean_name


def load_scan(scan_file: Path) -> tuple[pd.DataFrame, str]:
    if scan_file.suffix.lower() == ".csv":
        try:
            scan = pd.read_csv(scan_file, skiprows=10, encoding="utf-8-sig")
        except UnicodeDecodeError:
            scan = pd.read_csv(scan_file, skiprows=10, encoding="cp874")
    else:
        scan = pd.read_excel(scan_file, skiprows=10)

    scan.columns = scan.iloc[0]
    scan = scan[1:].reset_index(drop=True)
    scan = scan[["Timestamp", "Employee ID", "Employee Name"]].copy()
    scan["Timestamp"] = pd.to_datetime(
        scan["Timestamp"],
        format="%d-%m-%Y %H:%M:%S",
        errors="coerce",
    )
    scan = scan.dropna(subset=["Timestamp"])
    scan["emp_id"] = scan["Employee ID"].astype(str).str.strip()
    scan["date"] = scan["Timestamp"].dt.date
    scan["time"] = scan["Timestamp"].dt.strftime("%H:%M")

    target_date = str(scan["date"].max())
    return scan, target_date


def load_employee(master_file: Path) -> pd.DataFrame:
    employee_raw = pd.read_excel(master_file)
    employee = employee_raw.rename(columns={
        "User ID (Job Information)": "emp_id",
        "First Name (Local)": "first_name",
        "Last Name (Local)": "last_name",
        "Employment Status": "status",
        "Name (Employment Type)": "employment_type",
        "Code (Position)": "position_code",
        "Title (Position)": "position",
        "หน่วยงาน": "dept",
        "Name (Section)": "section",
    })
    employee.columns = employee.columns.astype(str).str.strip()

    required_cols = [
        "emp_id",
        "first_name",
        "last_name",
        "status",
        "employment_type",
        "position_code",
        "position",
        "dept",
        "section",
    ]
    missing_cols = [col for col in required_cols if col not in employee.columns]
    if missing_cols:
        raise ValueError(
            f"Master employee ขาด column: {missing_cols}\n"
            f"Column ที่มีคือ: {employee.columns.tolist()}"
        )

    employee["emp_id"] = employee["emp_id"].apply(clean_emp_id)
    employee["name"] = (
        employee["first_name"].fillna("").astype(str).str.strip()
        + " "
        + employee["last_name"].fillna("").astype(str).str.strip()
    )
    employee["name_key"] = employee["name"].apply(clean_name)

    return employee[
        [
            "emp_id",
            "name",
            "name_key",
            "status",
            "employment_type",
            "position_code",
            "position",
            "dept",
            "section",
        ]
    ].copy()


def load_manpower_plan(manpower_file: Path, target_date: str) -> pd.DataFrame:
    manpower_raw = pd.read_excel(manpower_file)
    manpower = manpower_raw.rename(columns={
        "หน่วยงาน": "dept",
        "หน้างาน": "work_station",
        "กะ": "shift",
        "จำนวนคน": "required_hc",
        "เวลาเข้า": "shift_start",
        "เวลาออก": "shift_end",
    })
    manpower = manpower.dropna(subset=["required_hc", "work_station"]).copy()
    manpower["date"] = target_date
    manpower["required_skill"] = manpower["work_station"]
    manpower["dept"] = manpower["dept"].astype(str).str.strip()
    manpower["work_station"] = manpower["work_station"].astype(str).str.strip()
    manpower["shift"] = manpower["shift"].astype(str).str.strip()
    manpower["required_hc"] = pd.to_numeric(
        manpower["required_hc"],
        errors="coerce",
    )
    manpower = manpower.dropna(subset=["required_hc"]).copy()
    manpower["required_hc"] = manpower["required_hc"].astype(int)
    manpower["shift_start"] = pd.to_datetime(
        manpower["shift_start"].astype(str),
        errors="coerce",
    ).dt.strftime("%H:%M")
    manpower["shift_end"] = pd.to_datetime(
        manpower["shift_end"].astype(str),
        errors="coerce",
    ).dt.strftime("%H:%M")

    return manpower[
        [
            "date",
            "dept",
            "work_station",
            "shift",
            "required_hc",
            "shift_start",
            "shift_end",
            "required_skill",
        ]
    ].copy()


def load_skill_matrix(skill_file: Path) -> pd.DataFrame:
    skill_matrix = pd.read_excel(skill_file)
    skill_matrix = skill_matrix.rename(columns={
        "Employee ID": "emp_id",
        "Emp ID": "emp_id",
        "รหัสพนักงาน": "emp_id",
        "Skill": "skill",
        "ทักษะ": "skill",
        "Level": "level",
        "ระดับ": "level",
        "Can Do": "can_do",
        "ทำได้": "can_do",
    })
    skill_matrix["emp_id"] = skill_matrix["emp_id"].apply(clean_emp_id)
    skill_matrix["skill"] = skill_matrix["skill"].astype(str).str.strip()

    if "level" not in skill_matrix.columns:
        skill_matrix["level"] = 3
    if "can_do" not in skill_matrix.columns:
        skill_matrix["can_do"] = 1

    skill_matrix["level"] = (
        pd.to_numeric(skill_matrix["level"], errors="coerce")
        .fillna(3)
        .astype(int)
    )
    skill_matrix["can_do"] = (
        pd.to_numeric(skill_matrix["can_do"], errors="coerce")
        .fillna(1)
        .astype(int)
    )

    return skill_matrix[["emp_id", "skill", "level", "can_do"]].copy()

