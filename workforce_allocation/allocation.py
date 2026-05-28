import pandas as pd
from pulp import LpBinary, LpMinimize, LpProblem, LpStatus, LpVariable, lpSum


def build_options(
    manpower_plan: pd.DataFrame,
    available: pd.DataFrame,
    skill_matrix: pd.DataFrame,
) -> pd.DataFrame:
    options = []

    for plan_id, plan_row in manpower_plan.iterrows():
        target_dept = str(plan_row["dept"]).strip()
        required_skill = str(plan_row["required_skill"]).strip()
        target_shift = str(plan_row["shift"]).strip()

        dept_candidates = available[
            (available["dept"].astype(str).str.strip() == target_dept)
            & (available["shift"].astype(str).str.strip() == target_shift)
        ].copy()
        dept_candidates = dept_candidates.merge(skill_matrix, on="emp_id", how="left")
        dept_candidates["skill"] = dept_candidates["skill"].fillna("")
        dept_candidates["level"] = dept_candidates["level"].fillna(1).astype(int)
        dept_candidates["skill_match"] = (
            dept_candidates["skill"].astype(str).str.strip() == required_skill
        )

        for _, emp in dept_candidates.iterrows():
            options.append({
                "plan_id": plan_id,
                "emp_id": emp["emp_id"],
                "name": emp["name"],
                "home_dept": emp["dept"],
                "target_dept": target_dept,
                "work_station": plan_row["work_station"],
                "shift": target_shift,
                "required_skill": required_skill,
                "emp_skill": emp["skill"],
                "skill_match": emp["skill_match"],
                "skill_level": emp["level"] if emp["skill_match"] else 1,
                "attendance_status": emp["attendance_status"],
                "scan_in": emp["scan_in"],
                "scan_out": emp["scan_out"],
                "home_dept_match": True,
            })

    return pd.DataFrame(options)


def solve_allocation(
    options_df: pd.DataFrame,
    manpower_plan: pd.DataFrame,
) -> tuple[pd.DataFrame, str]:
    if options_df.empty:
        return pd.DataFrame(), "No candidates"

    model = LpProblem("Workforce_Allocation", LpMinimize)
    x = {i: LpVariable(f"x_{i}", cat=LpBinary) for i in options_df.index}

    penalties = []
    for i, row in options_df.iterrows():
        penalty = 0
        if row["attendance_status"] == "Late":
            penalty += 5
        penalty += max(0, 5 - int(row["skill_level"]))
        if "skill_match" in options_df.columns and row["skill_match"] is False:
            penalty += 20
        penalties.append(penalty * x[i])

    model += lpSum(penalties)

    for emp_id in options_df["emp_id"].unique():
        emp_options = options_df[options_df["emp_id"] == emp_id].index
        model += lpSum(x[i] for i in emp_options) <= 1

    for plan_id, plan_row in manpower_plan.iterrows():
        required_hc = int(plan_row["required_hc"])
        plan_options = options_df[options_df["plan_id"] == plan_id].index
        target_hc = min(required_hc, len(plan_options))
        model += lpSum(x[i] for i in plan_options) == target_hc

    model.solve()
    assigned_rows = [
        row.to_dict()
        for i, row in options_df.iterrows()
        if x[i].value() == 1
    ]
    return pd.DataFrame(assigned_rows), LpStatus[model.status]


def build_gap_summary(
    manpower_plan: pd.DataFrame,
    allocation_result: pd.DataFrame,
    target_date: str,
    match_by_plan_id: bool = True,
) -> pd.DataFrame:
    summary_rows = []

    for plan_id, plan_row in manpower_plan.iterrows():
        required_hc = int(plan_row["required_hc"])
        if allocation_result.empty:
            assigned_hc = 0
        elif match_by_plan_id:
            assigned_hc = len(allocation_result[allocation_result["plan_id"] == plan_id])
        else:
            assigned_hc = len(
                allocation_result[
                    (
                        allocation_result["target_dept"].astype(str).str.strip()
                        == str(plan_row["dept"]).strip()
                    )
                    & (
                        allocation_result["work_station"].astype(str).str.strip()
                        == str(plan_row["work_station"]).strip()
                    )
                    & (
                        allocation_result["shift"].astype(str).str.strip()
                        == str(plan_row["shift"]).strip()
                    )
                ]
            )

        summary_rows.append({
            "date": target_date,
            "dept": plan_row["dept"],
            "work_station": plan_row["work_station"],
            "shift": plan_row["shift"],
            "required_hc": required_hc,
            "assigned_hc": assigned_hc,
            "gap": assigned_hc - required_hc,
            "shortage": max(required_hc - assigned_hc, 0),
            "surplus": max(assigned_hc - required_hc, 0),
        })

    return pd.DataFrame(summary_rows)


def rebalance_within_same_dept_shift(
    allocation_result: pd.DataFrame,
    gap_summary: pd.DataFrame,
) -> tuple[pd.DataFrame, int]:
    allocation_result = allocation_result.copy()
    if allocation_result.empty:
        return allocation_result, 0

    if "allocation_type" not in allocation_result.columns:
        allocation_result["allocation_type"] = "Solver Allocation"

    move_rows = []
    shortage_ws = gap_summary[gap_summary["gap"] < 0].copy()
    shortage_ws["shortage_hc"] = shortage_ws["gap"].abs()
    surplus_ws = gap_summary[gap_summary["gap"] > 0].copy()
    surplus_ws["surplus_hc"] = surplus_ws["gap"]

    for _, shortage in shortage_ws.iterrows():
        target_dept = str(shortage["dept"]).strip()
        target_ws = str(shortage["work_station"]).strip()
        target_shift = str(shortage["shift"]).strip()
        need_hc = int(shortage["shortage_hc"])

        donor_list = surplus_ws[
            (surplus_ws["dept"].astype(str).str.strip() == target_dept)
            & (surplus_ws["shift"].astype(str).str.strip() == target_shift)
        ].copy()

        for donor_idx, donor in donor_list.iterrows():
            if need_hc <= 0:
                break

            donor_ws = str(donor["work_station"]).strip()
            surplus_hc = int(surplus_ws.loc[donor_idx, "surplus_hc"])
            if surplus_hc <= 0:
                continue

            donor_people = allocation_result[
                (allocation_result["target_dept"].astype(str).str.strip() == target_dept)
                & (allocation_result["work_station"].astype(str).str.strip() == donor_ws)
                & (allocation_result["shift"].astype(str).str.strip() == target_shift)
            ].copy()
            if donor_people.empty:
                continue

            move_count = min(need_hc, surplus_hc, len(donor_people))
            selected_people = donor_people.head(move_count)

            for idx, person in selected_people.iterrows():
                new_row = person.copy()
                new_row["work_station"] = target_ws
                new_row["target_dept"] = target_dept
                new_row["shift"] = target_shift
                new_row["allocation_type"] = (
                    f"Moved from {donor_ws} to {target_ws} within same shift"
                )
                move_rows.append(new_row)
                allocation_result = allocation_result.drop(index=idx)

            need_hc -= move_count
            surplus_ws.loc[donor_idx, "surplus_hc"] -= move_count

    if move_rows:
        allocation_result = pd.concat(
            [allocation_result, pd.DataFrame(move_rows)],
            ignore_index=True,
        )

    return allocation_result, len(move_rows)

