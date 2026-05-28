from dataclasses import dataclass
from pathlib import Path


DEFAULT_BASE_FOLDER = Path(
    r"C:\Users\chompoopan.jan\OneDrive - Charoen Pokphand Foods Group\Documents\workforce_system"
)


@dataclass(frozen=True)
class AppConfig:
    base_folder: Path = DEFAULT_BASE_FOLDER
    late_grace_minutes: int = 5

    @property
    def input_folder(self) -> Path:
        return self.base_folder / "Input"

    @property
    def output_folder(self) -> Path:
        return self.base_folder / "Output"


def find_input_files(input_folder: Path) -> dict[str, Path]:
    files = list(input_folder.iterdir())

    master_files = [
        f for f in files
        if f.suffix.lower() in (".xlsx", ".xls")
        and ("active" in f.name.lower() or "master" in f.name.lower())
    ]
    manpower_files = [
        f for f in files
        if f.suffix.lower() in (".xlsx", ".xls")
        and "manpower" in f.name.lower()
    ]
    skill_files = [
        f for f in files
        if f.suffix.lower() in (".xlsx", ".xls")
        and ("skill" in f.name.lower() or "matrix" in f.name.lower())
    ]

    if not master_files:
        raise FileNotFoundError("Master employee file not found in Input folder.")
    if not manpower_files:
        raise FileNotFoundError("Manpower file not found in Input folder.")
    if not skill_files:
        raise FileNotFoundError("Skill matrix file not found in Input folder.")

    return {
        "master_file": master_files[0],
        "manpower_file": manpower_files[0],
        "skill_file": skill_files[0],
    }


def ask_scan_file() -> Path:
    from tkinter import Tk, filedialog

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    scan_file = filedialog.askopenfilename(
        title="เลือกไฟล์ Timestamp / Time Record",
        filetypes=[
            ("Excel or CSV files", "*.csv *.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()

    if not scan_file:
        raise FileNotFoundError("ยังไม่ได้เลือกไฟล์ timestamp")

    return Path(scan_file)

