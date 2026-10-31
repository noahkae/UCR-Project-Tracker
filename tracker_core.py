"""Read and write the tracker's Excel files. No UI code lives here.

Folder layout (a SharePoint library, synced to each PC with OneDrive):

    Project Tracker/
        Master.xlsx            PM/captain edits Key Dates; the app rewrites All Parts / All Actions
        Subteams/
            Aero.xlsx          only the Aero lead can edit
            Suspension.xlsx
            ...

A subteam file has four sheets:
    Plan       one row per part: parent, owner, assignees, tags, current gate(s), flag, notes, a due date per gate
               ("N/A" in a gate's date cell skips that gate for the part)
    Actions    follow-ups from design reviews: action, part, owner, due, from, closed
    Gate Log   written by the app when a gate is passed: part, gate, due, passed
    Gates      the gate rules: which gates each gate comes after (gates can run in parallel), and the key
               date it must be due before. A file without it runs its gates in order, one after another.

Permissions are SharePoint's: if Windows won't let the app write a file (you don't have
Edit access), the app treats it as view-only.
"""

import contextlib
import json
import os
import re
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
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
GATE_SHEET = "Gates"
FIXED_COLS = ["Part", "Parent", "Owner", "Assignees", "Tags", "Current Gate", "Flag", "Notes"]
# Older files have a subset of these, in this order; they're read by header name and upgraded on save.
SEP = " / "  # a sub-part's full name is "Parent / Part"
SKIP = "N/A"  # in a gate's date cell: the part skips that gate
ACTION_COLS = ["Action", "Part", "Owner", "Due", "From", "Closed"]
LOG_COLS = ["Part", "Gate", "Due", "Passed"]
GATE_COLS = ["Gate", "After", "Due before"]
DEFAULT_GATES = ["DR01 - Concept", "DR02 - Detailed Design", "DR03 - Final Design", "Manufacturing", "Testing"]
# Files made with the original default gates are shown with the new names, and rewritten with them on the next save.
OLD_DEFAULT_GATES = ["Concept", "Detailed Design", "Drawings", "Manufacturing", "Testing"]
FLAGS = ["OK", "At risk", "Blocked"]
DONE = "Done"
KEY_DATE_COLS = ["Date", "Event", "Type", "People", "Done", "Internal deadline", "End date", "Track", "Competition", "Notes"]
KEY_DATE_TYPES = ["Deadline", "Design Review", "Milestone", "Freeze", "Blackout", "Event"]
PERIOD_TYPES = ("Design Review", "Blackout", "Event")  # things that happen, rather than targets you can miss
TRACKS = ["Team", "Mech", "Elec"]
ROLLUP_COLS = ["Subteam", "Part", "Parent", "Owner", "Assignees", "Tags", "Current Gate", "Due", "Status", "Flag", "Notes"]
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
    current: str = ""  # as read from the file; only used to work out passed gates in files the app didn't write
    flag: str = "OK"
    notes: str = ""
    dates: dict = field(default_factory=dict)  # gate -> due date | None
    parent: str = ""  # name of the part this is a sub-part of (one level only), or ""
    assignees: list = field(default_factory=list)  # people working on it besides the owner
    tags: list = field(default_factory=list)  # e.g. ["Development", "Aero-Testing"]
    passed: dict = field(default_factory=dict)  # gate -> (date passed | None if unknown, due date when passed)
    skipped: set = field(default_factory=set)  # gates that don't apply to this part

    @property
    def key(self):
        """What the Actions and Gate Log sheets call this part: "Front Wing / Element 1" for a sub-part."""
        return f"{self.parent}{SEP}{self.name}" if self.parent else self.name

    def people(self):
        return _unique([self.owner] + self.assignees)


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
    passed: date  # None when the gate was passed before the app recorded it


@dataclass
class Subteam:
    name: str
    path: Path
    gates: list  # gate names in column order (always an order in which every gate comes after its prerequisites)
    parts: list
    writable: bool
    mtime: float
    actions: list = field(default_factory=list)
    after: dict = field(default_factory=dict)  # gate -> gates it comes after
    due_before: dict = field(default_factory=dict)  # gate -> key date (event name) it must be due before
    unmatched_log: list = field(default_factory=list)  # Gate Log rows for parts/gates that no longer exist (e.g. renamed in Excel)

    @property
    def log(self):
        """Every gate passed, part by part (what the Gate Log sheet holds), then any rows the app couldn't match."""
        return [GatePass(p.key, g, *reversed(p.passed[g])) for p in self.parts for g in self.gates if g in p.passed] + self.unmatched_log


@dataclass
class KeyDate:
    date: date = None
    event: str = ""
    type: str = "Deadline"
    people: list = field(default_factory=list)
    done: bool = False
    internal: date = None  # the team's own deadline, ahead of the official one
    end: date = None  # last day, for a blackout or an event that spans days
    track: str = ""  # Team / Mech / Elec
    competition: str = ""
    notes: str = ""


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


def linear(gates):
    """The default rules: each gate comes after the one before it."""
    return {g: gates[i - 1:i] for i, g in enumerate(gates)}


def cleared(part, gate):
    return gate in part.passed or gate in part.skipped


def open_gates(part, team):
    return [g for g in team.gates if not cleared(part, g)]


def active_gates(part, team):
    """The gates being worked on now: not passed or skipped, and everything they come after is cleared (all the way back,
    so skipping a gate doesn't let the gates after it start before the gates before it)."""
    return [g for g in open_gates(part, team) if all(cleared(part, a) for a in ancestors(team, g))]


def gate_state(part, team, gate):
    """passed / skipped / active / waiting."""
    if gate in part.passed:
        return "passed"
    if gate in part.skipped:
        return "skipped"
    return "active" if gate in active_gates(part, team) else "waiting"


def is_done(part, team):
    return not open_gates(part, team)


def current_gate(part, team, sep=" + "):
    """What the part is working on: its active gate(s), or Done."""
    return DONE if is_done(part, team) else sep.join(active_gates(part, team))


def due_date(part, team):
    """The earliest due date among the active gates."""
    dates = [part.dates.get(g) for g in active_gates(part, team) if part.dates.get(g)]
    return min(dates) if dates else None


def status(part, team, today=None):
    """A part's status comes from the earliest due date among its active gates."""
    return date_status(due_date(part, team), is_done(part, team), today)


def needs_attention(part, team, today=None):
    return status(part, team, today) == "Late" or part.flag != "OK"


def progress(part, team):
    """(gates cleared, gates in total)."""
    return sum(cleared(part, g) for g in team.gates), len(team.gates)


def pass_gate(part, gate, today=None):
    part.passed[gate] = (today or date.today(), part.dates.get(gate))
    part.skipped.discard(gate)


def unpass_gate(part, team, gate):
    """Un-pass a gate, and with it every passed gate that comes after it."""
    for g in [gate] + descendants(team, gate):
        part.passed.pop(g, None)


def ancestors(team, gate):
    seen, todo = [], list(team.after.get(gate, []))
    while todo:
        g = todo.pop()
        if g in team.gates and g not in seen:
            seen.append(g)
            todo += team.after.get(g, [])
    return [g for g in team.gates if g in seen]


def descendants(team, gate):
    return [g for g in team.gates if gate in ancestors(team, g)]


def rule_warnings(part, team, key_dates):
    """Problems with a part's gate dates under the subteam's gate rules (warnings, not errors)."""
    fmt = lambda d: d.strftime("%a %d %b")
    events = {k.event.casefold(): k for k in sorted((k for k in key_dates if k.date), key=lambda k: k.date, reverse=True)}
    blackouts = [k for k in key_dates if k.type == "Blackout" and k.date]
    out = []
    for g in open_gates(part, team):
        due = part.dates.get(g)
        if not due:
            continue
        for a in team.after.get(g, []):
            if not cleared(part, a) and part.dates.get(a) and due < part.dates[a]:
                out.append(f"{g} is due before {a}, which comes first.")
        kd = events.get(team.due_before.get(g, "").casefold())
        if kd and due > kd.date:
            out.append(f"{g} is due after {kd.event} ({fmt(kd.date)}); it has to be due before it.")
        for b in blackouts:
            if b.date <= due <= (b.end or b.date):
                out.append(f"{g} is due during {b.event}.")
    return out


def action_status(action, today=None):
    return date_status(action.due, action.closed is not None, today)


def key_date_target(kd):
    """The date the team is working to: the internal deadline when there is one."""
    return kd.internal or kd.date


def key_date_status(kd, today=None):
    """Deadlines, milestones and freezes can be late (against the internal deadline if set). Reviews, blackouts
    and events just happen, so once they're over they're Held."""
    today = today or date.today()
    if kd.type in PERIOD_TYPES and kd.date and (kd.end or kd.date) < today:
        return "Held"
    if kd.type in PERIOD_TYPES and kd.date and kd.date <= today:
        return "Now"
    return date_status(kd.date if kd.type in PERIOD_TYPES else key_date_target(kd), kd.done, today)


def key_date_needs_attention(kd, today=None):
    return key_date_status(kd, today) in ("Late", "Due soon")


def split_names(text):
    """'Sam, Alex & Jo' -> ['Sam', 'Alex', 'Jo']."""
    return [n.strip() for n in re.split(r"[,;/&]| and ", text or "") if n.strip()]


def split_tags(text):
    return _unique(t.strip() for t in (text or "").split(","))


def _unique(items):
    """Non-blank items, first spelling wins, ignoring case."""
    seen, out = set(), []
    for x in items:
        if x and x.casefold() not in seen:
            seen.add(x.casefold())
            out.append(x)
    return out


def is_me(name, me):
    return bool(me) and name.strip().casefold() == me.strip().casefold()


def on_time(team):
    """(gates passed on or before their due date, gates passed that had a due date and a pass date)."""
    timed = [g for g in team.log if g.due and g.passed]
    return sum(1 for g in timed if g.passed <= g.due), len(timed)


def record_part_change(team, old, new):
    """Keep Actions in step with an edited part: renaming a part (or moving it under another part) renames it
    in the actions, and renaming a parent does the same for its sub-parts. (The Gate Log is written from the parts.)"""
    if old is None:
        return
    _rename_refs(team, old.key, new.key)
    if old.name != new.name and not old.parent:
        for child in children(team, old):
            before = child.key
            child.parent = new.name
            _rename_refs(team, before, child.key)


def _rename_refs(team, old_key, new_key):
    if old_key == new_key:
        return
    for action in team.actions:
        if action.part == old_key:
            action.part = new_key


def children(team, part):
    """The sub-parts of a top-level part."""
    return [p for p in team.parts if p.parent and p.parent == part.name] if not part.parent else []


def all_tags(team):
    return sorted(_unique(t for p in team.parts for t in p.tags), key=str.casefold)


def check_part(team, part, index=None):
    """The problem with saving `part` at position `index` (None = new), or None if it's fine."""
    if not part.name:
        return "Give the part a name."
    if SEP in part.name:
        return f'A part name can\'t contain "{SEP.strip()}" with spaces around it.'
    others = [p for i, p in enumerate(team.parts) if i != index]
    if any(p.key.casefold() == part.key.casefold() for p in others):
        return f'There\'s already a part called "{part.key}".'
    if part.parent:
        old = team.parts[index] if index is not None else None
        if part.parent == part.name or (old and part.parent == old.name):
            return "A part can't be its own parent."
        if not any(p.name == part.parent and not p.parent for p in others):
            return f'"{part.parent}" isn\'t a top-level part.'
        if old and children(team, old):
            return f"{old.name} has sub-parts, so it can't be a sub-part itself."
    return None


def check_gates(names):
    """The problem with a proposed gate list, or None if it's fine."""
    if not names:
        return "A subteam needs at least one gate."
    if any(not n for n in names):
        return "A gate can't be blank."
    if any(c in n for n in names for c in ',"+'):
        return "Gate names can't contain commas, quotes or +."
    reserved = [DONE, SKIP] + FIXED_COLS
    if any(n.casefold() == r.casefold() for n in names for r in reserved):
        return f"These names are reserved: {', '.join(reserved)}."
    if len({n.casefold() for n in names}) < len(names):
        return "Two gates have the same name."
    return None


def change_gates(team, gates, after=None, due_before=None):
    """Replace a subteam's gates. `gates` is a list of (old name or None for a new gate, new name), in order.

    `after` and `due_before` give the new rules by new name; without them, renamed gates keep their rules and a
    new gate comes after the one before it. Rules only point back up the list, so the order stays valid.
    Renamed gates keep their due dates, passes and skips; a removed gate's are dropped.
    """
    rename = {old: new for old, new in gates if old}
    names = [new for _, new in gates]
    if after is None:
        after = {new: [rename[a] for a in team.after.get(old, []) if a in rename] if old else names[i - 1:i]
                 for i, (old, new) in enumerate(gates)}
    if due_before is None:
        due_before = {new: team.due_before.get(old, "") for old, new in gates if old}
    team.after = {g: [a for a in after.get(g, []) if a in names[:i]] for i, g in enumerate(names)}
    team.due_before = {g: due_before.get(g, "") for g in names if due_before.get(g)}
    for p in team.parts:
        p.dates = {new: p.dates.get(old) if old else None for old, new in gates}
        p.passed = {rename[g]: v for g, v in p.passed.items() if g in rename}
        p.skipped = {rename[g] for g in p.skipped if g in rename}
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
    cols = _headers(ws)
    key_dates = []
    for row in ws.iter_rows(min_row=2, max_col=len(cols), values_only=True):
        cell = lambda name: row[cols.index(name)] if name in cols else None
        if _text(cell("Event")):
            key_dates.append(KeyDate(
                _date(cell("Date")), _text(cell("Event")), _text(cell("Type")) or "Deadline", split_names(_text(cell("People"))),
                _truthy(cell("Done")), _date(cell("Internal deadline")), _date(cell("End date")), _text(cell("Track")),
                _text(cell("Competition")), _text(cell("Notes"))))
    teams = [load_subteam(p) for p in sorted((folder / SUBTEAM_DIR).glob("*.xlsx")) if not p.name.startswith("~$")]
    return Tracker(folder, key_dates, [t for t in teams if t], is_writable(master), master.stat().st_mtime)


def load_subteam(path):
    wb = load_workbook(path, data_only=True)
    if PLAN_SHEET not in wb.sheetnames:
        return None
    ws = wb[PLAN_SHEET]
    fixed, gates = _plan_layout(ws)
    renamed = dict(zip(OLD_DEFAULT_GATES, DEFAULT_GATES)) if gates == OLD_DEFAULT_GATES else {}
    gates = [renamed.get(g, g) for g in gates]
    after, due_before = _gate_rules(wb, gates)
    team = Subteam(path.stem, path, gates, [], is_writable(path), path.stat().st_mtime, after=after, due_before=due_before)
    n = len(fixed)
    for row in ws.iter_rows(min_row=2, max_col=n + len(gates), values_only=True):
        cell = lambda col: _text(row[fixed.index(col)]) if col in fixed else ""
        if not cell("Part"):
            continue
        flag = cell("Flag") or "OK"
        cells = dict(zip(gates, row[n:]))
        team.parts.append(Part(
            name=cell("Part"), owner=cell("Owner"), current=cell("Current Gate"), flag=flag if flag in FLAGS else "OK",
            notes=cell("Notes"), dates={g: _date(v) for g, v in cells.items()}, parent=cell("Parent"),
            assignees=split_names(cell("Assignees")), tags=split_tags(cell("Tags")),
            skipped={g for g, v in cells.items() if _text(v).upper() in (SKIP, "NA", "SKIP")},
        ))
    team.actions = [Action(str(r[0]).strip(), _text(r[1]), _text(r[2]), _date(r[3]), _text(r[4]), _date(r[5]))
                    for r in _rows(wb, ACTION_SHEET, len(ACTION_COLS))]
    by_key = {p.key: p for p in team.parts}
    for r in _rows(wb, LOG_SHEET, len(LOG_COLS)):
        part, gate = by_key.get(_text(r[0])), renamed.get(_text(r[1]), _text(r[1]))
        if part and gate in gates:
            part.passed[gate] = (_date(r[3]), _date(r[2]))
        else:
            team.unmatched_log.append(GatePass(_text(r[0]), gate, _date(r[2]), _date(r[3])))
    for p in team.parts:
        _passed_from_current(p, team, renamed)
    return team


def _passed_from_current(part, team, renamed):
    """Gates passed before the app logged them: everything a listed current gate comes after (or every gate, for Done).
    This is how files edited by hand, or made by older versions, say which gates are passed."""
    listed = [renamed.get(g.strip(), g.strip()) for g in re.split(r"[,+]", part.current) if g.strip()]
    if DONE in listed:
        before = team.gates
    else:
        before = [a for g in listed if g in team.gates for a in ancestors(team, g)]
    for g in before:
        if g not in part.passed and g not in part.skipped:
            part.passed[g] = (None, part.dates.get(g))


def _gate_rules(wb, gates):
    """(after, due_before) from the Gates sheet, or the default linear rules when there isn't one."""
    if GATE_SHEET not in wb.sheetnames:
        return linear(gates), {}
    after, due_before = {g: [] for g in gates}, {}
    for r in _rows(wb, GATE_SHEET, len(GATE_COLS)):
        g = _text(r[0])
        if g in gates:
            earlier = gates[:gates.index(g)]
            after[g] = [a.strip() for a in _text(r[1]).split(",") if a.strip() in earlier]
            if _text(r[2]):
                due_before[g] = _text(r[2])
    return after, due_before


def _headers(ws):
    headers = [_text(c.value) for c in ws[1]]
    return headers[:headers.index("")] if "" in headers else headers


def _rows(wb, sheet, width):
    if sheet not in wb.sheetnames:
        return []
    return [list(r) + [None] * width for r in wb[sheet].iter_rows(min_row=2, max_col=width, values_only=True) if r[0]]


def _text(v):
    return str(v).strip() if v is not None else ""


def _truthy(v):
    return isinstance(v, (date, datetime)) or str(v or "").strip().lower() in ("yes", "y", "x", "done", "true", "1")


def _date(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 36526 <= v <= 73051:  # Excel serials for 2000..2099
        return (datetime(1899, 12, 30) + timedelta(days=int(v))).date()
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
    """Add the Actions, Gate Log and Gates sheets to a subteam file that doesn't have them yet."""
    if ACTION_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(ACTION_SHEET)
        _header(ws, ACTION_COLS)
        for col, width in zip("ABCDEF", (48, 24, 18, 12, 12, 12)):
            ws.column_dimensions[col].width = width
        ws.freeze_panes = "A2"
    if LOG_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(LOG_SHEET)
        _header(ws, LOG_COLS)
        for col, width in zip("ABCD", (34, 24, 12, 12)):
            ws.column_dimensions[col].width = width
        ws["F1"] = "Written by the tracker app when a gate is passed. A blank Passed date: passed before the app recorded it."
        ws["F1"].font = Font(italic=True, color="757575")
        ws.freeze_panes = "A2"
    if GATE_SHEET not in wb.sheetnames:
        ws = wb.create_sheet(GATE_SHEET)
        _header(ws, GATE_COLS)
        for col, width in zip("ABC", (30, 50, 28)):
            ws.column_dimensions[col].width = width
        ws["E1"] = "After: the gates (comma-separated, earlier in the list) that must be passed or skipped first. Blank = can start any time."
        ws["E2"] = "Due before: a key date (its Event name in Master.xlsx) this gate must be due before. The app warns if it isn't."
        for c in ("E1", "E2"):
            ws[c].font = Font(italic=True, color="757575")
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
    Replaces the existing one, so changing the gates (or upgrading an older layout) keeps Excel's formatting right."""
    if PLAN_SHEET in wb.sheetnames:
        wb.remove(wb[PLAN_SHEET])
    ws = wb.create_sheet(PLAN_SHEET, 0)
    wb.active = 0
    n = len(FIXED_COLS)
    _header(ws, FIXED_COLS + gates, [HEADER_FILL] * n + [GATE_FILL] * len(gates))
    for letter, width in zip("ABCDEFGH", (28, 20, 18, 26, 22, 30, 10, 40)):
        ws.column_dimensions[letter].width = width
    col = lambda name: get_column_letter(FIXED_COLS.index(name) + 1)
    cur = col("Current Gate")
    first, last = get_column_letter(n + 1), get_column_letter(n + len(gates))
    for i in range(len(gates)):
        letter = get_column_letter(n + 1 + i)
        ws.column_dimensions[letter].width = max(15, len(gates[i]) + 4)
        for r in range(2, FORMAT_ROWS + 2):
            ws[f"{letter}{r}"].number_format = "yyyy-mm-dd"
    area = f"{first}2:{last}{FORMAT_ROWS + 1}"
    # A gate listed in Current Gate: red when overdue, yellow otherwise. Skipped gates (N/A) and finished parts: grey.
    listed = f'ISNUMBER(SEARCH(", "&{first}$1&",", ", "&${cur}2&","))'
    ws.conditional_formatting.add(area, FormulaRule(
        formula=[f'AND({listed},ISNUMBER({first}2),{first}2<TODAY())'],
        fill=PatternFill("solid", fgColor="F9DEDC"), font=Font(bold=True, color="8C1D18")))
    ws.conditional_formatting.add(area, FormulaRule(formula=[listed], fill=PatternFill("solid", fgColor="FFF3C4"), font=Font(bold=True)))
    ws.conditional_formatting.add(area, FormulaRule(formula=[f'OR(${cur}2="{DONE}",{first}2="{SKIP}")'], font=Font(color="9E9E9E")))
    _list(ws, f"{col('Flag')}2:{col('Flag')}{FORMAT_ROWS + 1}", FLAGS)
    _list(ws, f"{col('Parent')}2:{col('Parent')}{FORMAT_ROWS + 1}", source=f"=$A$2:$A${FORMAT_ROWS + 1}")
    ws.freeze_panes = "B2"


def _plan_layout(ws):
    """(fixed columns, gates) from the Plan sheet's header row. Fixed columns are read by name, so older layouts
    (no Parent, Assignees or Tags) still load; the gates are the named columns after the last fixed one. A fixed
    column whose header was cleared is just ignored (and restored on the next save)."""
    headers = [_text(c.value) for c in ws[1]]
    n = max((i + 1 for i, h in enumerate(headers) if h in FIXED_COLS), default=0)
    fixed, gates = headers[:n], headers[n:]
    if fixed and fixed[0] == "":
        fixed[0] = "Part"  # column A is always the part name
    return fixed, gates[:gates.index("")] if "" in gates else gates


def create_subteam_file(path, gates=DEFAULT_GATES):
    """An empty subteam plan with headers, dropdowns and highlighting."""
    wb = Workbook()
    wb.remove(wb.active)
    _plan_sheet(wb, gates)
    _ensure_side_sheets(wb)
    _write_rows(wb[GATE_SHEET], [[g, ", ".join(a), None] for g, a in linear(gates).items()], len(GATE_COLS), set())
    _save_atomic(wb, Path(path))


def create_master(path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Key Dates"
    _key_date_sheet(ws)
    about = wb.create_sheet("Read Me")
    for line in [
        "Project Tracker: master file",
        "",
        "Key Dates: competition deadlines, design reviews (gate meetings), milestones, freezes, blackouts and events.",
        "   Only the PM/captain edits this.",
        "   People = who's responsible (comma-separated). They see it under 'My work' in the app.",
        "   Done = Yes once a deadline is submitted or a milestone is met.",
        "   Internal deadline = the team's own date ahead of the official one; the app tracks against it.",
        "   End date = last day of a blackout or an event that spans several days.",
        "All Parts / All Actions: rebuilt by the tracker app from the subteam files. Don't edit them here.",
        "",
        "Each subteam's plan is in Subteams/<name>.xlsx, and only that subteam's lead can edit it.",
        "In a subteam file, the gates are the columns after Notes, and each holds that gate's due date (or N/A to skip it).",
        "Parent: blank for a top-level part; for a sub-part, the name of the part it belongs to (e.g. Front Wing).",
        "Current Gate: the gate(s) the part is working on, comma-separated, or Done.",
        "The Gates sheet says which gates come after which (they can run in parallel) and which key date each must beat.",
    ]:
        about.append([line])
    about["A1"].font = Font(bold=True, size=14)
    about.column_dimensions["A"].width = 110
    _save_atomic(wb, Path(path))


def _key_date_sheet(ws):
    _header(ws, KEY_DATE_COLS)
    for col, width in zip("ABCDEFGHIJ", (12, 50, 16, 36, 8, 16, 12, 10, 24, 60)):
        ws.column_dimensions[col].width = width
    for letter in "AFG":
        for r in range(2, FORMAT_ROWS + 2):
            ws[f"{letter}{r}"].number_format = "yyyy-mm-dd"
    ws.data_validations.dataValidation = []
    _list(ws, f"C2:C{FORMAT_ROWS + 1}", KEY_DATE_TYPES)
    _list(ws, f"E2:E{FORMAT_ROWS + 1}", ["Yes"])
    _list(ws, f"H2:H{FORMAT_ROWS + 1}", TRACKS)
    ws.freeze_panes = "A2"


def save_subteam(team):
    """Write the parts, actions, gate log and gate rules back to the file. Raises a SaveError subclass if it can't."""
    ensure_can_save(team.path, team.mtime)
    wb = load_workbook(team.path)
    _ensure_side_sheets(wb)
    if _plan_layout(wb[PLAN_SHEET]) != (FIXED_COLS, team.gates):
        _plan_sheet(wb, team.gates)
    _write_rows(
        wb[PLAN_SHEET],
        [[p.name, p.parent, p.owner, ", ".join(p.assignees), ", ".join(p.tags), current_gate(p, team, ", "), p.flag, p.notes]
         + [SKIP if g in p.skipped else p.dates.get(g) for g in team.gates] for p in team.parts],
        len(FIXED_COLS) + len(team.gates),
        set(range(len(FIXED_COLS) + 1, len(FIXED_COLS) + len(team.gates) + 1)),
    )
    _write_rows(wb[ACTION_SHEET], [[a.text, a.part, a.owner, a.due, a.source, a.closed] for a in team.actions],
                len(ACTION_COLS), {4, 6})
    _write_rows(wb[LOG_SHEET], [[g.part, g.gate, g.due, g.passed] for g in team.log], len(LOG_COLS), {3, 4})
    _write_rows(wb[GATE_SHEET], [[g, ", ".join(team.after.get(g, [])), team.due_before.get(g)] for g in team.gates],
                len(GATE_COLS), set())
    _save_atomic(wb, team.path)
    team.mtime = team.path.stat().st_mtime


def save_key_dates(tracker):
    path = tracker.folder / MASTER
    ensure_can_save(path, tracker.master_mtime)
    wb = load_workbook(path)
    ws = wb["Key Dates"]
    _key_date_sheet(ws)  # also upgrades a master made before the newer columns existed
    _write_rows(ws, [[k.date, k.event, k.type, ", ".join(k.people), "Yes" if k.done else None, k.internal, k.end, k.track,
                      k.competition, k.notes]
                     for k in sorted(tracker.key_dates, key=lambda k: k.date or date.max)],
                len(KEY_DATE_COLS), {1, 6, 7})
    _save_atomic(wb, path)
    tracker.master_mtime = path.stat().st_mtime


def write_rollup(tracker, today=None):
    """Regenerate Master.xlsx > All Parts and All Actions from every subteam file."""
    path = tracker.folder / MASTER
    ensure_can_save(path)  # regenerated from scratch, so a newer master isn't a conflict
    wb = load_workbook(path)
    stamp = f"Generated by the tracker app {datetime.now():%Y-%m-%d %H:%M}. Don't edit it here."
    sheets = [
        ("All Parts", ROLLUP_COLS, (24, 28, 20, 18, 26, 22, 30, 11, 10, 9, 40), {8},
         [[t.name, p.name, p.parent, p.owner, ", ".join(p.assignees), ", ".join(p.tags), current_gate(p, t, ", "),
           due_date(p, t), status(p, t, today), p.flag, p.notes]
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


def key_date_days(kd, limit=62):
    """Every day a key date covers (one day, or start..end for a period)."""
    if not kd.date:
        return []
    end = kd.end if kd.end and kd.end >= kd.date else kd.date
    return [kd.date + timedelta(days=i) for i in range(min((end - kd.date).days + 1, limit))]
