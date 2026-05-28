from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from .allocation import (
    build_gap_summary,
    build_options,
    rebalance_within_same_dept_shift,
    solve_allocation,
)
from .attendance import (
    add_department_to_scan,
    build_attendance_today,
    detect_shift_and_availability,
    rebuild_attendance,
)
from .config import AppConfig, ask_scan_file, find_input_files
from .exporters import export_result
from .loaders import (
    load_employee,
    load_manpower_plan,
    load_scan,
    load_skill_matrix,
)


@dataclass
class PipelineResult:
    target_date: str
    solver_status: str
    moved_count: int
    output_file: Path
    allocation_result: pd.DataFrame
    gap_summary: pd.DataFrame
    attendance_today: pd.DataFrame
    manpower_plan: pd.DataFrame
    options_df: pd.DataFrame


def run_pipeline(
    config: AppConfig | None = None,
    scan_file: Path | None = None,
) -> PipelineResult:
    config = config or AppConfig()
    input_files = find_input_files(config.input_folder)
    scan_file = scan_file or ask_scan_file()

    scan, target_date = load_scan(scan_file)
    employee = load_employee(input_files["master_file"])
    attendance = rebuild_attendance(scan, employee)
    scan_with_dept = add_department_to_scan(scan, employee)
    manpower_plan = load_manpower_plan(input_files["manpower_file"], target_date)
    attendance_today = build_attendance_today(employee, attendance, target_date)
    attendance_today, available = detect_shift_and_availability(
        attendance_today,
        manpower_plan,
        config.late_grace_minutes,
    )
    skill_matrix = load_skill_matrix(input_files["skill_file"])
    options_df = build_options(manpower_plan, available, skill_matrix)
    allocation_result, solver_status = solve_allocation(options_df, manpower_plan)

    gap_summary = build_gap_summary(
        manpower_plan,
        allocation_result,
        target_date,
        match_by_plan_id=True,
    )
    allocation_result, moved_count = rebalance_within_same_dept_shift(
        allocation_result,
        gap_summary,
    )
    gap_summary = build_gap_summary(
        manpower_plan,
        allocation_result,
        target_date,
        match_by_plan_id=False,
    )
    output_file = export_result(
        config.output_folder,
        target_date,
        allocation_result,
        gap_summary,
        attendance_today,
        manpower_plan,
        employee,
        scan_with_dept,
    )

    return PipelineResult(
        target_date=target_date,
        solver_status=solver_status,
        moved_count=moved_count,
        output_file=output_file,
        allocation_result=allocation_result,
        gap_summary=gap_summary,
        attendance_today=attendance_today,
        manpower_plan=manpower_plan,
        options_df=options_df,
    )

