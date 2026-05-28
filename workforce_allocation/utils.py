import re

import pandas as pd


def clean_emp_id(value) -> str:
    if pd.isna(value):
        return ""

    value = str(value).replace("\u00a0", "").strip()
    value = re.sub(r"\s+", "", value)
    value = re.sub(r"\.0$", "", value)
    value = re.sub(r"[^0-9]", "", value)
    return value


def clean_name(value) -> str:
    if pd.isna(value):
        return ""

    value = str(value).replace("\u00a0", " ").strip()
    value = re.sub(r"\s+", "", value)
    value = re.sub(
        r"^(นาย|นางสาว|นาง|mr|mrs|miss|ms)",
        "",
        value,
        flags=re.IGNORECASE,
    )
    return value


clean_id = clean_emp_id

