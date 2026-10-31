"""Read and write the tracker's Excel files. No UI code lives here.

Folder layout (a SharePoint library, synced to each PC with OneDrive):

    Project Tracker/
        Master.xlsx            PM/captain edits Key Dates; the app rewrites All Parts / All Actions
        Subteams/
            Aero.xlsx          only the Aero lead can edit
            Suspension.xlsx
            ...

A subteam file has three sheets:
    Plan       one row per part: owner, current gate, flag, notes, a due date per gate
    Actions    follow-ups from design reviews: action, part, owner, due, from, closed
    Gate Log   written by the app when a gate is passed: part, gate, due, passed

Permissions are SharePoint's: if Windows won't let the app write a file (you don't have
Edit access), the app treats it as view-only.
"""

import contextlib
import json
import os
import re
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
SHAREPOINT_FOLDER = Path("UCalgary-Racing - Documents") / "Project Management"  # inside each member's OneDrive
PLAN_SHEET = "Plan"
ACTION_SHEET = "Actions"
LOG_SHEET = "Gate Log"
FIXED_COLS = ["Part", "Owner", "Current Gate", "Flag", "Notes"]
ACTION_COLS = ["Action", "Part", "Owner", "Due", "From", "Closed"]
LOG_COLS = ["Part", "Gate", "Due", "Passed"]
DEFAULT_GATES = ["DR01 - Concept", "DR02 - Detailed Design", "DR03 - Final Design", "Manufacturing", "Testing"]
# Files made with the original default gates are shown with the new names, and rewritten with them on the next save.
OLD_DEFAULT_GATES = ["Concept", "Detailed Design", "Drawings", "Manufacturing", "Testing"]
FLAGS = ["OK", "At risk", "Blocked"]
DONE = "Done"
KEY_DATE_COLS = ["Date", "Event", "Type", "People", "Done"]
KEY_DATE_TYPES = ["Deadline", "Design Review", "Milestone"]
ROLLUP_COLS = ["Subteam", "Part", "Owner", "Current Gate", "Due", "Status", "Flag", "Notes"]
DUE_SOON_DAYS = 7
FORMAT_ROWS = 300  # rows that get dropdowns / highlighting in the Excel files

CONFIG = Path(os.environ.get("APPDATA", Path.home())) / "FSAETracker" / "config.json"  # per-user app settings

HEADER_FONT = Font(bold=True, color="FFFFFF")
HEADER_FILL = PatternFill("solid", fgColor="3F51B5")  # Material indigo 500
GATE_FILL = PatternFill("solid", fgColor="5C6BC0")  # indigo 400


class SaveError(Exception):
    """Base for the reasons a save can be refused; the message is shown to the user."""


class ReadOnlyError(SaveError):
    """Windows refused the write, i.e. you don't have edit rights on SharePoint."""


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
class Action:
    text: str = ""
    part: str = ""
    owner: str = ""
    due: date = None
    source: str = ""  # which review / meeting it came from, e.g. "DR2"
    closed: date = None


@dataclass
class GatePass:
    part: str
    gate: str
    due: date
    passed: date


@dataclass
class Subteam:
    name: str
    path: Path
    gates: list  # gate names in column order
    parts: list
    writable: bool
    mtime: float
    actions: list = field(default_factory=list)
    log: list = field(default_factory=list)


@dataclass
class KeyDate:
    date: date = None
    event: str = ""
    type: str = "Deadline"
    people: list = field(default_factory=list)
    done: bool = False


@dataclass
class Tracker:
    folder: Path
    key_dates: list
    subteams: list
    master_writable: bool
    master_mtime: float


# --------------------------------------------------------------------------- rules


def date_status(due, done=False, today=None):
    """Done / Late / Due soon / On track / No date."""
    today = today or date.today()
    if done:
        return DONE
    if due is None:
        return "No date"
    if today > due:
        return "Late"
    if (due - today).days <= DUE_SOON_DAYS:
        return "Due soon"
    return "On track"


def current_gate(part, gates):
    if part.current == DONE or part.current in gates:
        return part.current
    return gates[0] if gates else DONE


def due_date(part, gates):
    gate = current_gate(part, gates)
    return None if gate == DONE else part.dates.get(gate)


def status(part, gates, today=None):
    """A part's status comes from its current gate's due date."""
    return date_status(due_date(part, gates), current_gate(part, gates) == DONE, today)


def needs_attention(part, gates, today=None):
    return status(part, gates, today) == "Late" or part.flag != "OK"


def progress(part, gates):
    """(gates passed, gates in total)."""
    gate = current_gate(part, gates)
    return (len(gates) if gate == DONE else gates.index(gate)), len(gates)


def action_status(action, today=None):
    return date_status(action.due, action.closed is not None, today)


def key_date_status(kd, today=None):
    """Deadlines and milestones can be late; a design review is a meeting, so once it's past it's just held."""
    today = today or date.today()
    if kd.type == "Design Review" and kd.date and kd.date < today:
        return "Held"
    return date_status(kd.date, kd.done, today)


def key_date_needs_attention(kd, today=None):
    return key_date_status(kd, today) in ("Late", "Due soon")


def split_names(text):
    """'Sam, Alex & Jo' -> ['Sam', 'Alex', 'Jo']."""
    return [n.strip() for n in re.split(r"[,;/&]| and ", text or "") if n.strip()]


def is_me(name, me):
    return bool(me) and name.strip().casefold() == me.strip().casefold()


def on_time(team):
    """(gates passed on or before their due date, gates passed that had a due date)."""
    timed = [g for g in team.log if g.due]
    return sum(1 for g in timed if g.passed <= g.due), len(timed)


def record_part_change(team, old, new, today=None):
    """Keep the Gate Log and Actions in step with an edited part.

    Moving the current gate forward logs each gate passed today; moving it back
    removes those log entries. Renaming a part renames it in the log and actions.
    """
    today = today or date.today()
    if old is None:
        return
    if old.name != new.name:
        for entry in team.log:
            if entry.part == old.name:
                entry.part = new.name
        for action in team.actions:
            if action.part == old.name:
                action.part = new.name
    index = lambda p: len(team.gates) if current_gate(p, team.gates) == DONE else team.gates.index(current_gate(p, team.gates))
    before, after = index(old), index(new)
    if after > before:
        for gate in team.gates[before:after]:
            team.log = [e for e in team.log if not (e.part == new.name and e.gate == gate)]
            team.log.append(GatePass(new.name, gate, new.dates.get(gate), today))
    elif after < before:
        undone = set(team.gates[after:before])
        team.log = [e for e in team.log if not (e.part == new.name and e.gate in undone)]


def check_gates(names):
    """The problem with a proposed gate list, or None if it's fine."""
    if not names:
        return "A subteam needs at least one gate."
    if any(not n for n in names):
        return "A gate can't be blank."
    if any(c in n for n in names for c in ',"'):
        return "Gate names can't contain commas or quotes."
    if DONE.casefold() in (n.casefold() for n in names):
        return f'"{DONE}" is reserved; it means every gate is passed.'
    if len({n.casefold() for n in names}) < len(names):
        return "Two gates have the same name."
    if len(",".join(names + [DONE])) > 250:
        return "The gate names are too long in total for Excel's dropdown."
    return None


def change_gates(team, gates):
    """Replace a subteam's gates. `gates` is a list of (old name or None for a new gate, new name), in order.

    Renamed gates keep their due dates and gate history. A removed gate loses its due dates;
    a part whose current gate was removed moves to the first gate it hasn't passed yet.
    """
    rename = {old: new for old, new in gates if old}
    names = [new for _, new in gates]
    for p in team.parts:
        cur = current_gate(p, team.gates)
        if cur != DONE:
            if cur in rename:
                cur = rename[cur]
            else:
                passed = {rename.get(g) for g in team.gates[:team.gates.index(cur)]}
                cur = next((n for n in names if n not in passed), DONE)
        p.current = cur
        p.dates = {new: p.dates.get(old) if old else None for old, new in gates}
    for entry in team.log:
        entry.gate = rename.get(entry.gate, entry.gate)
    team.gates = names


# --------------------------------------------------------------------------- file checks


def is_writable(path):
    """Whether Windows will let us write the file. OneDrive doesn't always mark files you can only view as
    read-only (the sync client blocks the write instead), so actually open it for writing; nothing is written."""
    if not os.access(path, os.W_OK):
        return False
    try:
        with open(path, "r+b"):
            return True
    except OSError:
        return False


def ensure_can_save(path, loaded_mtime=None):
    if not is_writable(path):
        raise ReadOnlyError("Edit failed.")
    if loaded_mtime is not None and abs(path.stat().st_mtime - loaded_mtime) > 1e-6:
        raise ChangedOnDiskError(f"{path.name} was changed by someone else. Your edit wasn't saved; it has been reloaded.")


def _save_atomic(wb, path):
    # OneDrive ignores *.tmp files, so a half-written file never syncs.
    tmp = path.with_name(path.name + ".tmp")
    try:
        wb.save(tmp)
        os.replace(tmp, path)
    except OSError as e:
        with contextlib.suppress(OSError):
            tmp.unlink()
        raise ReadOnlyError("Edit failed.") from e


# --------------------------------------------------------------------------- loading


def load_config():
    try:
        return json.loads(CONFIG.read_text())
    except (OSError, ValueError):
        return {}


def save_config(cfg):
    CONFIG.parent.mkdir(parents=True, exist_ok=True)
    CONFIG.write_text(json.dumps(cfg))


def default_folder():
    """The tracker folder the app last used, or one found under OneDrive."""
    folder = load_config().get("folder")
    if folder and (Path(folder) / MASTER).exists():
        return Path(folder)
    return find_folder()


def onedrive_roots():
    """The user's OneDrive folders, work/school account first (e.g. C:/Users/<you>/OneDrive - University of Calgary)."""
    roots = [os.environ.get("OneDriveCommercial"), os.environ.get("OneDrive")]
    roots += [str(p) for p in sorted(Path.home().glob("OneDrive - *"))]
    return list(dict.fromkeys(Path(r) for r in roots if r))


def default_location():
    """Where the team keeps the tracker: the 'UCalgary-Racing - Documents' SharePoint shortcut in your OneDrive.

    Every member adds that shortcut once ('Add shortcut to My files' in SharePoint), so the path only
    differs by the Windows user name, which OneDrive's environment variables already account for.
    """
    roots = onedrive_roots() or [Path.home() / "OneDrive - University of Calgary"]
    for root in roots:
        if (root / SHAREPOINT_FOLDER).is_dir():
            return root / SHAREPOINT_FOLDER
    return roots[0] / SHAREPOINT_FOLDER


def tracker_in(folder):
    """The tracker folder at or directly inside `folder` (the one holding Master.xlsx), or None."""
    folder = Path(folder)
    if (folder / MASTER).exists():
        return folder
    return next((hit.parent for hit in sorted(folder.glob(f"*/{MASTER}"))), None) if folder.is_dir() else None


def find_folder():
    """The tracker in the team's default location, else the first Master.xlsx found in the user's OneDrive folders."""
    hit = tracker_in(default_location())
    if hit:
        return hit
    bases = onedrive_roots() + [p for p in Path.home().iterdir() if p.is_dir() and p.name == "SharePoint"]
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
    for row in ws.iter_rows(min_row=2, max_col=len(KEY_DATE_COLS), values_only=True):
        row = list(row) + [None] * len(KEY_DATE_COLS)
        if row[1]:
            key_dates.append(KeyDate(_date(row[0]), str(row[1]).strip(), str(row[2] or "Deadline").strip(),
                                     split_names(str(row[3] or "")), _truthy(row[4])))
    teams = [load_subteam(p) for p in sorted((folder / SUBTEAM_DIR).glob("*.xlsx")) if not p.name.startswith("~$")]
    return Tracker(folder, key_dates, [t for t in teams if t], is_writable(master), master.stat().st_mtime)


def load_subteam(path):
    wb = load_workbook(path, data_only=True)
    if PLAN_SHEET not in wb.sheetnames:
        return None
    ws = wb[PLAN_SHEET]
    gates = _gate_headers(ws)
    renamed = dict(zip(OLD_DEFAULT_GATES, DEFAULT_GATES)) if gates == OLD_DEFAULT_GATES else {}
    gates = [renamed.get(g, g) for g in gates]
    parts = []
    n = len(FIXED_COLS)
    for row in ws.iter_rows(min_row=2, max_col=n + len(gates), values_only=True):
        if not row[0]:
            continue
        flag = str(row[3] or "OK").strip()
        parts.append(Part(
            name=str(row[0]).strip(),
            owner=str(row[1] or "").strip(),
            current=renamed.get(str(row[2] or "").strip(), str(row[2] or "").strip()),
            flag=flag if flag in FLAGS else "OK",
            notes=str(row[4] or "").strip(),
            dates={g: _date(row[n + i]) for i, g in enumerate(gates)},
        ))
    actions = [
        Action(str(r[0]).strip(), _text(r[1]), _text(r[2]), _date(r[3]), _text(r[4]), _date(r[5]))
        for r in _rows(wb, ACTION_SHEET, len(ACTION_COLS))
    ]
    log = [GatePass(str(r[0]).strip(), renamed.get(_text(r[1]), _text(r[1])), _date(r[2]), _date(r[3]))
           for r in _rows(wb, LOG_SHEET, len(LOG_COLS)) if r[1] and _date(r[3])]
    return Subteam(path.stem, path, gates, parts, is_writable(path), path.stat().st_mtime, actions, log)


def _rows(wb, sheet, width):
    if sheet not in wb.sheetnames:
        return []
    return [list(r) for r in wb[sheet].iter_rows(min_row=2, max_col=width, values_only=True) if r[0]]


def _text(v):
    return str(v).strip() if v is not None else ""


def _truthy(v):
    return isinstance(v, (date, datetime)) or str(v or "").strip().lower() in ("yes", "y", "x", "done", "true", "1")


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


def _list(ws, cells, values=None, source=None):
    dv = DataValidation(type="list", formula1=source or ('"' + ",".join(values) + '"'), allow_blank=True)
    dv.add(cells)
    ws.add_data_validation(dv)


def _ensure_side_sheets(wb):
    """Add the Actions and Gate Log sheets to a subteam file that doesn't have them yet."""
    if ACTION_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(ACTION_SHEET)
        _header(ws, ACTION_COLS)
        for col, width in zip("ABCDEF", (48, 24, 18, 12, 12, 12)):
            ws.column_dimensions[col].width = width
        _list(ws, f"B2:B{FORMAT_ROWS + 1}", source=f"={PLAN_SHEET}!$A$2:$A${FORMAT_ROWS + 1}")
        ws.freeze_panes = "A2"
    if LOG_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(LOG_SHEET)
        _header(ws, LOG_COLS)
        for col, width in zip("ABCD", (28, 18, 12, 12)):
            ws.column_dimensions[col].width = width
        ws["F1"] = "Written by the tracker app when a gate is passed."
        ws["F1"].font = Font(italic=True, color="757575")
        ws.freeze_panes = "A2"


def _write_rows(ws, rows, width, date_cols):
    for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, 2), max_col=width):
        for c in row:
            c.value = None
    for r, values in enumerate(rows, 2):
        for i, v in enumerate(values, 1):
            c = ws.cell(r, i, v if v not in ("", None) else None)
            if i in date_cols:
                c.number_format = "yyyy-mm-dd"


def _plan_sheet(wb, gates):
    """A fresh, empty Plan sheet for these gates, with headers, dropdowns and highlighting.
    Replaces the existing one, so changing the gates (or upgrading old default gates) keeps Excel's formatting right."""
    if PLAN_SHEET in wb.sheetnames:
        wb.remove(wb[PLAN_SHEET])
    ws = wb.create_sheet(PLAN_SHEET, 0)
    wb.active = 0
    n = len(FIXED_COLS)
    _header(ws, FIXED_COLS + gates, [HEADER_FILL] * n + [GATE_FILL] * len(gates))
    for col, width in zip("ABCDE", (28, 18, 16, 10, 40)):
        ws.column_dimensions[col].width = width
    first, last = get_column_letter(n + 1), get_column_letter(n + len(gates))
    for i in range(len(gates)):
        letter = get_column_letter(n + 1 + i)
        ws.column_dimensions[letter].width = max(15, len(gates[i]) + 4)
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


def _gate_headers(ws):
    n = len(FIXED_COLS)
    headers = [str(c.value).strip() if c.value is not None else "" for c in ws[1]][n:]
    return headers[:headers.index("")] if "" in headers else headers


def create_subteam_file(path, gates=DEFAULT_GATES):
    """An empty subteam plan with headers, dropdowns and highlighting."""
    wb = Workbook()
    wb.remove(wb.active)
    _plan_sheet(wb, gates)
    _ensure_side_sheets(wb)
    _save_atomic(wb, Path(path))


def create_master(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Key Dates"
    _header(ws, KEY_DATE_COLS)
    for col, width in zip("ABCDE", (12, 50, 16, 36, 8)):
        ws.column_dimensions[col].width = width
    for r in range(2, FORMAT_ROWS + 2):
        ws[f"A{r}"].number_format = "yyyy-mm-dd"
    _list(ws, f"C2:C{FORMAT_ROWS + 1}", KEY_DATE_TYPES)
    _list(ws, f"E2:E{FORMAT_ROWS + 1}", ["Yes"])
    ws.freeze_panes = "A2"
    about = wb.create_sheet("Read Me")
    for line in [
        "Project Tracker: master file",
        "",
        "Key Dates: competition deadlines, design reviews (gate meetings) and milestones. Only the PM/captain edits this.",
        "   People = who's responsible for meeting the date (comma-separated). They see it under 'My work' in the app.",
        "   Done = Yes once a deadline is submitted or a milestone is met.",
        "All Parts / All Actions: rebuilt by the tracker app from the subteam files. Don't edit them here.",
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
    """Write the parts, actions and gate log back to the file. Raises a SaveError subclass if it can't."""
    ensure_can_save(team.path, team.mtime)
    wb = load_workbook(team.path)
    _ensure_side_sheets(wb)
    if _gate_headers(wb[PLAN_SHEET]) != team.gates:
        _plan_sheet(wb, team.gates)
    _write_rows(
        wb[PLAN_SHEET],
        [[p.name, p.owner, p.current, p.flag, p.notes] + [p.dates.get(g) for g in team.gates] for p in team.parts],
        len(FIXED_COLS) + len(team.gates),
        set(range(len(FIXED_COLS) + 1, len(FIXED_COLS) + len(team.gates) + 1)),
    )
    _write_rows(wb[ACTION_SHEET], [[a.text, a.part, a.owner, a.due, a.source, a.closed] for a in team.actions],
                len(ACTION_COLS), {4, 6})
    _write_rows(wb[LOG_SHEET], [[g.part, g.gate, g.due, g.passed] for g in team.log], len(LOG_COLS), {3, 4})
    _save_atomic(wb, team.path)
    team.mtime = team.path.stat().st_mtime


def save_key_dates(tracker):
    path = tracker.folder / MASTER
    ensure_can_save(path, tracker.master_mtime)
    wb = load_workbook(path)
    ws = wb["Key Dates"]
    _header(ws, KEY_DATE_COLS)  # also upgrades a master made before People / Done existed
    ws.column_dimensions["D"].width = 36
    _write_rows(ws, [[k.date, k.event, k.type, ", ".join(k.people), "Yes" if k.done else None]
                     for k in sorted(tracker.key_dates, key=lambda k: k.date or date.max)],
                len(KEY_DATE_COLS), {1})
    _save_atomic(wb, path)
    tracker.master_mtime = path.stat().st_mtime


def write_rollup(tracker, today=None):
    """Regenerate Master.xlsx > All Parts and All Actions from every subteam file."""
    path = tracker.folder / MASTER
    ensure_can_save(path)  # regenerated from scratch, so a newer master isn't a conflict
    wb = load_workbook(path)
    stamp = f"Generated by the tracker app {datetime.now():%Y-%m-%d %H:%M}. Don't edit it here."
    sheets = [
        ("All Parts", ROLLUP_COLS, (24, 28, 18, 16, 11, 10, 9, 40), {5},
         [[t.name, p.name, p.owner, current_gate(p, t.gates), due_date(p, t.gates), status(p, t.gates, today), p.flag, p.notes]
          for t in tracker.subteams for p in t.parts]),
        ("All Actions", ["Subteam"] + ACTION_COLS + ["Status"], (24, 48, 24, 18, 11, 10, 11, 10), {5, 7},
         [[t.name, a.text, a.part, a.owner, a.due, a.source, a.closed, action_status(a, today)]
          for t in tracker.subteams for a in t.actions]),
    ]
    for i, (name, cols, widths, date_cols, rows) in enumerate(sheets, 1):
        if name in wb.sheetnames:
            del wb[name]
        ws = wb.create_sheet(name, i)
        _header(ws, cols)
        _write_rows(ws, rows, len(cols), date_cols)
        for c, width in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = width
        ws.cell(1, len(cols) + 2, stamp).font = Font(italic=True, color="757575")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{max(ws.max_row, 2)}"
    _save_atomic(wb, path)
    tracker.master_mtime = path.stat().st_mtime
