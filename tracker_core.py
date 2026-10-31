"""Read and write the tracker's Excel files. No UI code lives here.

Folder layout (a SharePoint library, synced to each PC with OneDrive):

    Project Tracker/
        Master.xlsx            PM/captain edits Key Dates; the app rewrites All Parts
        Subteams/
            Aero.xlsx          only the Aero lead can edit - one sheet, "Plan"
            Suspension.xlsx
            ...

Permissions are SharePoint's: a file you can't edit syncs as read-only, and the
app treats it as view-only.
"""

import os
from dataclasses import dataclass, field
from datetime import date, datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

MASTER = "Master.xlsx"
SUBTEAM_DIR = "Subteams"
PLAN_SHEET = "Plan"
FIXED_COLS = ["Part", "Owner", "Current Gate", "Flag", "Notes"]
DEFAULT_GATES = ["Concept", "Detailed Design", "Drawings", "Manufacturing", "Testing"]
FLAGS = ["OK", "At risk", "Blocked"]
DONE = "Done"
KEY_DATE_COLS = ["Date", "Event", "Type"]
KEY_DATE_TYPES = ["Deadline", "Design Review", "Milestone"]
ROLLUP_COLS = ["Subteam", "Part", "Owner", "Current Gate", "Due", "Status", "Flag", "Notes"]
DUE_SOON_DAYS = 7
FORMAT_ROWS = 300  # rows that get dropdowns / highlighting in the Excel files

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="3F51B5")  # Material indigo 500
GATE_FILL = PatternFill("solid", fgColor="5C6BC0")  # indigo 400


class SaveError(Exception):
    """Base for the reasons a save can be refused; the message is shown to the user."""


class ReadOnlyError(SaveError):
    """The file is synced read-only, i.e. you don't have edit rights on SharePoint."""


class FileBusyError(SaveError):
    """The file is open in Excel (or otherwise locked)."""


class ChangedOnDiskError(SaveError):
    """Someone saved the file since we loaded it."""


# --------------------------------------------------------------------------- model


@dataclass
class Part:
    name: str = ""
    owner: str = ""
    current: str = ""
    flag: str = "OK"
    notes: str = ""
    dates: dict = field(default_factory=dict)  # gate -> date | None


@dataclass
class Subteam:
    name: str
    path: Path
    gates: list  # gate names in column order
    parts: list
    writable: bool
    mtime: float


@dataclass
class KeyDate:
    date: date = None
    event: str = ""
    type: str = "Deadline"


@dataclass
class Tracker:
    folder: Path
    key_dates: list
    subteams: list
    master_writable: bool
    master_mtime: float


def current_gate(part, gates):
    if part.current == DONE or part.current in gates:
        return part.current
    return gates[0] if gates else DONE


def status(part, gates, today=None):
    """Done / Late / Due soon / On track / No date - from the current gate's due date."""
    today = today or date.today()
    gate = current_gate(part, gates)
    if gate == DONE:
        return DONE
    due = part.dates.get(gate)
    if due is None:
        return "No date"
    if today > due:
        return "Late"
    if (due - today).days <= DUE_SOON_DAYS:
        return "Due soon"
    return "On track"


def due_date(part, gates):
    gate = current_gate(part, gates)
    return None if gate == DONE else part.dates.get(gate)


def needs_attention(part, gates, today=None):
    return status(part, gates, today) == "Late" or part.flag != "OK"


def progress(part, gates):
    """(gates passed, gates in total)."""
    gate = current_gate(part, gates)
    return (len(gates) if gate == DONE else gates.index(gate)), len(gates)


# --------------------------------------------------------------------------- file checks


def is_writable(path):
    """OneDrive syncs files you can only view as read-only."""
    return os.access(path, os.W_OK)


def ensure_can_save(path, loaded_mtime=None):
    if not is_writable(path):
        raise ReadOnlyError(f"{path.name} is view-only for you (SharePoint permissions).")
    if loaded_mtime is not None and abs(path.stat().st_mtime - loaded_mtime) > 1e-6:
        raise ChangedOnDiskError(f"{path.name} was changed by someone else. Your edit wasn't saved; it has been reloaded.")
    try:
        with open(path, "r+b"):  # Excel holds a write lock while the file is open
            pass
    except PermissionError as e:
        raise FileBusyError(f"{path.name} is open in Excel. Close it there and save again.") from e


def _save_atomic(wb, path):
    # OneDrive ignores *.tmp files, so a half-written file never syncs.
    tmp = path.with_name(path.name + ".tmp")
    wb.save(tmp)
    os.replace(tmp, path)


# --------------------------------------------------------------------------- loading


def find_folder():
    """Look for a synced folder containing Master.xlsx under the user's OneDrive/SharePoint folders."""
    home = Path.home()
    bases = [Path(os.environ[e]) for e in ("OneDriveCommercial", "OneDrive") if os.environ.get(e)]
    bases += [p for p in home.iterdir() if p.is_dir() and (p.name.startswith("OneDrive") or p.name == "SharePoint")]
    for base in bases:
        for pattern in (MASTER, f"*/{MASTER}", f"*/*/{MASTER}"):
            for hit in base.glob(pattern):
                return hit.parent
    return None


def load(folder):
    folder = Path(folder)
    master = folder / MASTER
    ws = load_workbook(master, data_only=True)["Key Dates"]
    key_dates = []
    for row in ws.iter_rows(min_row=2, max_col=3, values_only=True):
        if row[1]:
            key_dates.append(KeyDate(_date(row[0]), str(row[1]).strip(), str(row[2] or "Deadline").strip()))
    teams = [load_subteam(p) for p in sorted((folder / SUBTEAM_DIR).glob("*.xlsx")) if not p.name.startswith("~$")]
    return Tracker(folder, key_dates, [t for t in teams if t], is_writable(master), master.stat().st_mtime)


def load_subteam(path):
    wb = load_workbook(path, data_only=True)
    if PLAN_SHEET not in wb.sheetnames:
        return None
    ws = wb[PLAN_SHEET]
    header = [str(c.value).strip() if c.value is not None else "" for c in ws[1]]
    gates = []
    for h in header[len(FIXED_COLS):]:
        if not h:
            break
        gates.append(h)
    parts = []
    n = len(FIXED_COLS)
    for row in ws.iter_rows(min_row=2, max_col=n + len(gates), values_only=True):
        if not row[0]:
            continue
        flag = str(row[3] or "OK").strip()
        parts.append(Part(
            name=str(row[0]).strip(),
            owner=str(row[1] or "").strip(),
            current=str(row[2] or "").strip(),
            flag=flag if flag in FLAGS else "OK",
            notes=str(row[4] or "").strip(),
            dates={g: _date(row[n + i]) for i, g in enumerate(gates)},
        ))
    return Subteam(path.stem, path, gates, parts, is_writable(path), path.stat().st_mtime)


def _date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, str) and v.strip():
        try:
            return date.fromisoformat(v.strip()[:10])
        except ValueError:
            return None
    return None


# --------------------------------------------------------------------------- writing


def _header(ws, names, fills=None):
    for i, h in enumerate(names, 1):
        c = ws.cell(1, i, h)
        c.font = HEADER_FONT
        c.fill = fills[i - 1] if fills else HEADER_FILL
        c.alignment = Alignment(horizontal="center")


def _list(ws, cells, values):
    dv = DataValidation(type="list", formula1='"' + ",".join(values) + '"', allow_blank=True)
    dv.add(cells)
    ws.add_data_validation(dv)


def create_subteam_file(path, gates=DEFAULT_GATES):
    """An empty subteam plan with headers, dropdowns and highlighting."""
    wb = Workbook()
    ws = wb.active
    ws.title = PLAN_SHEET
    n = len(FIXED_COLS)
    _header(ws, FIXED_COLS + gates, [HEADER_FILL] * n + [GATE_FILL] * len(gates))
    for col, width in zip("ABCDE", (28, 18, 16, 10, 40)):
        ws.column_dimensions[col].width = width
    first, last = get_column_letter(n + 1), get_column_letter(n + len(gates))
    for i in range(len(gates)):
        letter = get_column_letter(n + 1 + i)
        ws.column_dimensions[letter].width = 15
        for r in range(2, FORMAT_ROWS + 2):
            ws[f"{letter}{r}"].number_format = "yyyy-mm-dd"
    area = f"{first}2:{last}{FORMAT_ROWS + 1}"
    # The current gate's date: red when overdue, yellow otherwise. Gates to its left are done: grey.
    ws.conditional_formatting.add(area, FormulaRule(
        formula=[f'AND({first}$1=$C2,{first}2<>"",{first}2<TODAY())'],
        fill=PatternFill("solid", fgColor="F9DEDC"), font=Font(bold=True, color="8C1D18")))
    ws.conditional_formatting.add(area, FormulaRule(
        formula=[f"{first}$1=$C2"], fill=PatternFill("solid", fgColor="FFF3C4"), font=Font(bold=True)))
    ws.conditional_formatting.add(area, FormulaRule(
        formula=[f'OR($C2="{DONE}",IFERROR(COLUMN({first}2)<MATCH($C2,$1:$1,0),FALSE))'], font=Font(color="9E9E9E")))
    _list(ws, f"C2:C{FORMAT_ROWS + 1}", gates + [DONE])
    _list(ws, f"D2:D{FORMAT_ROWS + 1}", FLAGS)
    ws.freeze_panes = "B2"
    _save_atomic(wb, Path(path))


def create_master(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Key Dates"
    _header(ws, KEY_DATE_COLS)
    for col, width in zip("ABC", (12, 50, 16)):
        ws.column_dimensions[col].width = width
    for r in range(2, FORMAT_ROWS + 2):
        ws[f"A{r}"].number_format = "yyyy-mm-dd"
    _list(ws, f"C2:C{FORMAT_ROWS + 1}", KEY_DATE_TYPES)
    ws.freeze_panes = "A2"
    about = wb.create_sheet("Read Me")
    for line in [
        "Project Tracker: master file",
        "",
        "Key Dates: competition deadlines, design reviews (gate meetings) and milestones. Only the PM/captain edits this.",
        "All Parts: rebuilt by the tracker app from the subteam files. Don't edit it here.",
        "",
        "Each subteam's plan is in Subteams/<name>.xlsx, and only that subteam's lead can edit it.",
        "In a subteam file, the gates are the columns after Notes, in order, and each holds that gate's due date.",
        "'Current Gate' is the gate the part is working towards, so every gate to its left is done.",
        "Set Current Gate to Done when the part is finished. Flag a part 'At risk' or 'Blocked' to raise it at the next meeting.",
    ]:
        about.append([line])
    about["A1"].font = Font(bold=True, size=14)
    about.column_dimensions["A"].width = 110
    _save_atomic(wb, Path(path))


def save_subteam(team):
    """Write team.parts back to its file. Raises a SaveError subclass if it can't."""
    ensure_can_save(team.path, team.mtime)
    wb = load_workbook(team.path)
    ws = wb[PLAN_SHEET]
    width = len(FIXED_COLS) + len(team.gates)
    for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, 2), max_col=width):
        for c in row:
            c.value = None
    for r, p in enumerate(team.parts, 2):
        values = [p.name, p.owner or None, p.current or None, p.flag, p.notes or None] + [p.dates.get(g) for g in team.gates]
        for i, v in enumerate(values, 1):
            c = ws.cell(r, i, v)
            if i > len(FIXED_COLS):
                c.number_format = "yyyy-mm-dd"
    _save_atomic(wb, team.path)
    team.mtime = team.path.stat().st_mtime


def save_key_dates(tracker):
    path = tracker.folder / MASTER
    ensure_can_save(path, tracker.master_mtime)
    wb = load_workbook(path)
    ws = wb["Key Dates"]
    for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, 2), max_col=3):
        for c in row:
            c.value = None
    for r, kd in enumerate(sorted(tracker.key_dates, key=lambda k: k.date or date.max), 2):
        ws.cell(r, 1, kd.date).number_format = "yyyy-mm-dd"
        ws.cell(r, 2, kd.event)
        ws.cell(r, 3, kd.type)
    _save_atomic(wb, path)
    tracker.master_mtime = path.stat().st_mtime


def write_rollup(tracker, today=None):
    """Regenerate Master.xlsx > All Parts from every subteam file."""
    path = tracker.folder / MASTER
    ensure_can_save(path)  # regenerated from scratch, so a newer master isn't a conflict
    wb = load_workbook(path)
    if "All Parts" in wb.sheetnames:
        del wb["All Parts"]
    ws = wb.create_sheet("All Parts", 1)
    _header(ws, ROLLUP_COLS)
    for team in tracker.subteams:
        for p in team.parts:
            ws.append([team.name, p.name, p.owner or None, current_gate(p, team.gates), due_date(p, team.gates),
                       status(p, team.gates, today), p.flag, p.notes or None])
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
        row[0].number_format = "yyyy-mm-dd"
    for col, width in zip("ABCDEFGH", (24, 28, 18, 16, 11, 10, 9, 40)):
        ws.column_dimensions[col].width = width
    ws.cell(1, len(ROLLUP_COLS) + 2, f"Generated by the tracker app {datetime.now():%Y-%m-%d %H:%M}. Don't edit it here.").font = Font(italic=True, color="757575")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{max(ws.max_row, 2)}"
    _save_atomic(wb, path)
    tracker.master_mtime = path.stat().st_mtime
