"""Read and write the tracker's Excel files. No UI code lives here.

Folder layout (a SharePoint library, synced to each PC with OneDrive). Each car is its own tracker folder:

    Project Management/
        UCR-04/
            Master.xlsx        PM/captain: Key Dates, the general gate library and templates, Settings (car name);
                               the app rewrites All Parts / All Actions
            Subteams/
                Aero.xlsx      only the Aero lead can edit
                ...
        UCR-05/ ...

A subteam file has these sheets:
    Plan            one row per part: parent, owner, assignees, tags, gate template, current gate(s), flag, notes,
                    then a due date per gate ("N/A" = that gate is off for the part)
    Actions         follow-ups from design reviews: action, part, owner, due, from, closed
    Gate Log        gate states other than "not started": part, gate, due, passed (done date), status
    Design Reviews  DR1 Concept / DR2 Intermediate / DR3 Final per assembly: date, outcome, presenter, notes
    Gate Templates  this subteam's own gate templates: the gates in order, which gates each comes after, and the
                    key date each must be due before
    Gate Library    this subteam's own gates: code, name, description
    Settings        the subteam's default template
The master holds the same Gate Templates / Gate Library sheets for the general (all-subteam) ones.

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
REVIEW_SHEET = "Design Reviews"
TEMPLATE_SHEET = "Gate Templates"
LIBRARY_SHEET = "Gate Library"
SETTINGS_SHEET = "Settings"
LEGACY_GATE_SHEET = "Gates"  # one template per subteam, from the version before templates
FIXED_COLS = ["Part", "Parent", "Owner", "Assignees", "Tags", "Template", "Current Gate", "Flag", "Notes"]
# Older files have a subset of these, in this order; they're read by header name and upgraded on save.
SEP = " / "  # a sub-part's full name is "Parent / Part"
SKIP = "N/A"  # in a gate's date cell: the gate is off for that part
ACTION_COLS = ["Action", "Part", "Owner", "Due", "From", "Closed"]
LOG_COLS = ["Part", "Gate", "Due", "Passed", "Status"]
REVIEW_COLS = ["Part", "Review", "Date", "Outcome", "Presenter", "Notes"]
TEMPLATE_COLS = ["Template", "Designation", "Gate", "After", "Due before"]
LIBRARY_COLS = ["Code", "Gate", "Designation", "Description"]
DR_TYPES = ["DR1 Concept", "DR2 Intermediate", "DR3 Final"]
OUTCOMES = ["Pass", "Pass with actions", "Fail"]
GENERAL = "General"
NOT_STARTED, IN_PROGRESS, DONE_STATE, OFF = "Not started", "In progress", "Done", "Off"
STATES = [NOT_STARTED, IN_PROGRESS, DONE_STATE, OFF]
DEFAULT_TEMPLATE = "Standard"
DEFAULT_GATES = ["DR01 - Concept", "DR02 - Detailed Design", "DR03 - Final Design", "Manufacturing", "Testing"]
DEFAULT_CODES = ["DR01", "DR02", "DR03", "MFG", "TST"]
# Files made with the original default gates are shown with the new names, and rewritten with them on the next save.
OLD_DEFAULT_GATES = ["Concept", "Detailed Design", "Drawings", "Manufacturing", "Testing"]
FLAGS = ["OK", "At risk", "Blocked"]
DONE = "Done"
KEY_DATE_COLS = ["Date", "Event", "Type", "People", "Done", "Internal deadline", "End date", "Track", "Competition", "Notes"]
KEY_DATE_TYPES = ["Deadline", "Design Review", "Milestone", "Freeze", "Blackout", "Event"]
PERIOD_TYPES = ("Design Review", "Blackout", "Event")  # things that happen, rather than targets you can miss
TRACKS = ["Team", "Mech", "Elec"]
ROLLUP_COLS = ["Subteam", "Part", "Parent", "Owner", "Assignees", "Tags", "Template", "Current Gate", "Due", "Status", "Flag", "Notes"]
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
    current: str = ""  # as read from the file; only used to work out done gates in files the app didn't write
    flag: str = "OK"
    notes: str = ""
    dates: dict = field(default_factory=dict)  # gate -> due date | None
    parent: str = ""  # name of the part this is a sub-part of (one level only), or ""
    assignees: list = field(default_factory=list)  # people working on it besides the owner
    tags: list = field(default_factory=list)  # e.g. ["Development", "Aero-Testing"]
    passed: dict = field(default_factory=dict)  # done gates: gate -> (date done | None if unknown, due date when done)
    skipped: set = field(default_factory=set)  # gates that are off for this part
    template: str = ""  # gate template name; "" = the subteam's default
    started: set = field(default_factory=set)  # gates marked in progress

    @property
    def key(self):
        """What the Actions, Gate Log and Design Reviews sheets call this part: "Front Wing / Element 1" for a sub-part."""
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
    passed: date  # None when the gate was done before the app recorded it (or isn't done)
    status: str = DONE_STATE


@dataclass
class Review:
    part: str  # part key
    review: str  # one of DR_TYPES
    date: date = None
    outcome: str = ""
    presenter: str = ""
    notes: str = ""


@dataclass
class GateDef:
    """A gate in the library."""
    name: str
    code: str = ""
    designation: str = GENERAL  # General, or the subteam it belongs to
    description: str = ""


@dataclass
class Template:
    """A gate workflow: gates in order, what each comes after (gates can run in parallel), and key-date rules."""
    name: str
    designation: str = GENERAL
    gates: list = field(default_factory=list)
    after: dict = field(default_factory=dict)  # gate -> gates it comes after (always earlier in the list)
    due_before: dict = field(default_factory=dict)  # gate -> key date (event name) it must be due before


@dataclass
class Subteam:
    name: str
    path: Path
    parts: list
    writable: bool
    mtime: float
    actions: list = field(default_factory=list)
    templates: dict = field(default_factory=dict)  # this subteam's own templates, by name (file order)
    library: list = field(default_factory=list)  # this subteam's own gates (GateDef)
    reviews: list = field(default_factory=list)
    unmatched_log: list = field(default_factory=list)  # Gate Log rows for parts/gates that no longer exist
    shared: dict = field(default_factory=dict)  # the general templates (from the master); not saved here
    default_template: str = ""  # from the Settings sheet; "" = work it out
    columns: list = field(default_factory=list)  # gate columns as read from the Plan sheet

    @property
    def default(self):
        """The template new parts get, and parts with no template use."""
        if self.default_template in self.templates or self.default_template in self.shared:
            return self.default_template
        if self.templates:
            return next(iter(self.templates))
        if DEFAULT_TEMPLATE in self.shared:
            return DEFAULT_TEMPLATE
        return next(iter(self.shared), "")

    def available(self):
        """Templates parts here can use: this subteam's own, then the general ones."""
        return {**self.shared, **self.templates}

    @property
    def gates(self):
        """The gate columns of the Plan sheet: every gate of every template in use here, in template order."""
        used = [self.default] + list(self.templates) + [p.template for p in self.parts if p.template]
        out = []
        for name in dict.fromkeys(used):
            t = self.available().get(name)
            out += [g for g in (t.gates if t else []) if g not in out]
        return out or list(self.columns)

    # the default template's rules, for code (and tests) that think of a subteam as having one gate list
    @property
    def after(self):
        return template_of(self, None).after

    @property
    def due_before(self):
        return template_of(self, None).due_before

    @property
    def log(self):
        """Every gate that's done or in progress, part by part (what the Gate Log sheet holds), then unmatched rows."""
        rows = []
        for p in self.parts:
            for g in template_of(self, p).gates:
                if g in p.passed:
                    rows.append(GatePass(p.key, g, p.passed[g][1], p.passed[g][0], DONE_STATE))
                elif g in p.started:
                    rows.append(GatePass(p.key, g, p.dates.get(g), None, IN_PROGRESS))
        return rows + self.unmatched_log


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
    car: str = ""
    library: list = field(default_factory=list)  # general gates (GateDef)
    templates: dict = field(default_factory=dict)  # general templates, by name


def standard_template(gates=DEFAULT_GATES, name=DEFAULT_TEMPLATE, designation=GENERAL):
    return Template(name, designation, list(gates), linear(list(gates)), {})


def standard_library(gates=DEFAULT_GATES):
    codes = DEFAULT_CODES if list(gates) == DEFAULT_GATES else []
    out, taken = [], set()
    for i, g in enumerate(gates):
        code = codes[i] if codes else auto_code(g, taken)
        taken.add(code)
        out.append(GateDef(g, code, GENERAL))
    return out


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


def template_of(team, part):
    """The template a part follows (its own choice, else the subteam's default). Never None."""
    available = team.available()
    name = (part.template if part is not None and part.template in available else "") or team.default
    if name in available:
        return available[name]
    return Template(DEFAULT_TEMPLATE, GENERAL, list(team.columns), linear(list(team.columns)), {})


def cleared(part, gate):
    return gate in part.passed or gate in part.skipped


def ancestors(template, gate):
    seen, todo = [], list(template.after.get(gate, []))
    while todo:
        g = todo.pop()
        if g in template.gates and g not in seen:
            seen.append(g)
            todo += template.after.get(g, [])
    return [g for g in template.gates if g in seen]


def descendants(template, gate):
    return [g for g in template.gates if gate in ancestors(template, g)]


def open_gates(part, team):
    return [g for g in template_of(team, part).gates if not cleared(part, g)]


def active_gates(part, team):
    """Gates that are ready to work on: not done or off, and everything before them (all the way back) is cleared."""
    t = template_of(team, part)
    return [g for g in t.gates if not cleared(part, g) and all(cleared(part, a) for a in ancestors(t, g))]


def working_gates(part, team):
    """What the part is working on: the ready gates plus any marked in progress, in template order."""
    ready = active_gates(part, team)
    return [g for g in template_of(team, part).gates if g in ready or (g in part.started and not cleared(part, g))]


def gate_state(part, team, gate):
    """done / off / in progress / ready / waiting, or n/a when the gate isn't in the part's template."""
    if gate not in template_of(team, part).gates:
        return "n/a"
    if gate in part.passed:
        return "done"
    if gate in part.skipped:
        return "off"
    if gate in part.started:
        return "in progress"
    return "ready" if gate in active_gates(part, team) else "waiting"


def state_name(part, gate):
    """The state as the four-way toggle shows it."""
    if gate in part.passed:
        return DONE_STATE
    if gate in part.skipped:
        return OFF
    return IN_PROGRESS if gate in part.started else NOT_STARTED


def set_gate_state(part, gate, state, today=None):
    """Set one gate to Not started / In progress / Done / Off. Out-of-order states are allowed (they show as conflicts)."""
    part.started.discard(gate)
    part.skipped.discard(gate)
    if state != DONE_STATE:
        part.passed.pop(gate, None)
    if state == DONE_STATE and gate not in part.passed:
        part.passed[gate] = (today or date.today(), part.dates.get(gate))
    elif state == IN_PROGRESS:
        part.started.add(gate)
    elif state == OFF:
        part.skipped.add(gate)


def pass_gate(part, gate, today=None):
    set_gate_state(part, gate, DONE_STATE, today)


def unpass_gate(part, team, gate):
    """Un-pass a gate, and with it every done gate that comes after it."""
    for g in [gate] + descendants(template_of(team, part), gate):
        part.passed.pop(g, None)


def children(team, part):
    """The sub-parts of a top-level part."""
    return [p for p in team.parts if p.parent and p.parent == part.name] if not part.parent else []


def is_done(part, team):
    """Every gate of its template is done or off, or (for an assembly) every sub-assembly is done."""
    kids = children(team, part)
    return not open_gates(part, team) or (bool(kids) and all(is_done(k, team) for k in kids))


def current_gate(part, team, sep=" + ", rollup=True):
    """What the part is working on (its ready / in-progress gates), or Done. With rollup=False an assembly's own gates
    count, not its sub-assemblies (that's what the Plan sheet stores, since Done there means every gate is done)."""
    done = is_done(part, team) if rollup else not open_gates(part, team)
    return DONE if done else sep.join(working_gates(part, team))


def due_date(part, team):
    """The earliest due date among the gates it's working on."""
    if is_done(part, team):
        return None
    dates = [part.dates.get(g) for g in working_gates(part, team) if part.dates.get(g)]
    return min(dates) if dates else None


def status(part, team, today=None):
    return date_status(due_date(part, team), is_done(part, team), today)


def needs_attention(part, team, today=None):
    return status(part, team, today) == "Late" or part.flag != "OK" or bool(conflicts(part, team))


def progress(part, team):
    """(gates cleared, gates in its template)."""
    gates = template_of(team, part).gates
    return sum(cleared(part, g) for g in gates), len(gates)


def conflicts(part, team):
    """Gates marked done or in progress while a gate before them isn't done (or off)."""
    t = template_of(team, part)
    out = []
    for g in t.gates:
        if g in part.passed or (g in part.started and not cleared(part, g)):
            missing = [a for a in ancestors(t, g) if not cleared(part, a)]
            if missing:
                out.append(f"{g} is {'done' if g in part.passed else 'in progress'} but {', '.join(missing)} isn't done yet.")
    return out


def rule_warnings(part, team, key_dates):
    """Problems with a part's gates under its template's rules (warnings, not errors): out-of-order states, gates due
    before a gate they come after, after their 'due before' key date, or inside a blackout."""
    t = template_of(team, part)
    fmt = lambda d: d.strftime("%a %d %b")
    events = {k.event.casefold(): k for k in sorted((k for k in key_dates if k.date), key=lambda k: k.date, reverse=True)}
    blackouts = [k for k in key_dates if k.type == "Blackout" and k.date]
    out = conflicts(part, team)
    for g in open_gates(part, team):
        due = part.dates.get(g)
        if not due:
            continue
        for a in t.after.get(g, []):
            if not cleared(part, a) and part.dates.get(a) and due < part.dates[a]:
                out.append(f"{g} is due before {a}, which comes first.")
        kd = events.get(t.due_before.get(g, "").casefold())
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


def review_status(review, today=None):
    """Upcoming design reviews count down; once held they need an outcome."""
    today = today or date.today()
    if review.outcome:
        return "Held"
    if not review.date:
        return "No date"
    return "Needs outcome" if review.date < today else date_status(review.date, False, today)


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
    """(gates done on or before their due date, gates done that had a due date and a done date)."""
    timed = [g for g in team.log if g.status == DONE_STATE and g.due and g.passed]
    return sum(1 for g in timed if g.passed <= g.due), len(timed)


def record_part_change(team, old, new):
    """Keep Actions and Design Reviews in step with an edited part: renaming a part (or moving it under another part)
    renames it there, and renaming a parent does the same for its sub-parts. (The Gate Log is written from the parts.)"""
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
    for review in team.reviews:
        if review.part == old_key:
            review.part = new_key


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


RESERVED = [DONE, SKIP] + FIXED_COLS


def check_gates(names):
    """The problem with a proposed gate list, or None if it's fine."""
    if not names:
        return "A template needs at least one gate."
    if any(not n for n in names):
        return "A gate can't be blank."
    if any(c in n for n in names for c in ',"+'):
        return "Gate names can't contain commas, quotes or +."
    if any(n.casefold() == r.casefold() for n in names for r in RESERVED):
        return f"These names are reserved: {', '.join(RESERVED)}."
    if len({n.casefold() for n in names}) < len(names):
        return "Two gates have the same name."
    return None


def check_template_name(tracker, name, old=None):
    if not name:
        return "Give the template a name."
    names = {n.casefold() for n in all_templates(tracker)} - ({old.casefold()} if old else set())
    return f'There\'s already a template called "{name}".' if name.casefold() in names else None


def update_template(template, gates, after=None, due_before=None):
    """Change a template's gates in place. `gates` is a list of (old name or None for a new gate, new name), in order;
    `after` / `due_before` give the new rules by new name. Without them, renamed gates keep their rules and a new gate
    comes after the one before it. Rules only point back up the list, so the order stays valid."""
    rename = {old: new for old, new in gates if old}
    names = [new for _, new in gates]
    if after is None:
        after = {new: [rename[a] for a in template.after.get(old, []) if a in rename] if old else names[i - 1:i]
                 for i, (old, new) in enumerate(gates)}
    if due_before is None:
        due_before = {new: template.due_before.get(old, "") for old, new in gates if old}
    template.after = {g: [a for a in after.get(g, []) if a in names[:i]] for i, g in enumerate(names)}
    template.due_before = {g: due_before.get(g, "") for g in names if due_before.get(g)}
    template.gates = names


def remap_parts(team, template_name, gates):
    """After a template changed (`gates` = its (old or None, new) pairs): carry each affected part's dates and states
    across to the new gate names; gates that were removed lose theirs, new gates start empty."""
    for p in team.parts:
        if template_of(team, p).name != template_name:
            continue
        p.dates = {new: p.dates.get(old) if old else None for old, new in gates}
        p.passed = {new: p.passed[old] for old, new in gates if old in p.passed}
        p.skipped = {new for old, new in gates if old in p.skipped}
        p.started = {new for old, new in gates if old in p.started}


def change_gates(team, gates, after=None, due_before=None):
    """Change the gates of the subteam's default template (a subteam's own copy is made if the default is a general
    template). `gates` is (old name or None, new name) pairs, in order. Renamed gates keep their dates and states."""
    t = template_of(team, None)
    if t.name not in team.templates:
        t = Template(t.name if t.name not in team.shared else f"{team.name} gates", team.name, list(t.gates), dict(t.after),
                     dict(t.due_before))
        for p in team.parts:
            if not p.template or p.template == template_of(team, None).name:
                p.template = ""
        team.templates = {t.name: t, **team.templates}
        team.default_template = t.name
    update_template(t, gates, after, due_before)
    remap_parts(team, t.name, gates)
    ensure_library(team, t)


# ---- gate library


def auto_code(name, taken=()):
    """A short code for a gate: initials ("Mounting Design" -> "MD"), or the first letters of a one-word name."""
    words = re.findall(r"[A-Za-z0-9]+", name)
    base = ("".join(w[0] for w in words) if len(words) > 1 else (words[0][:3] if words else "G")).upper()[:5]
    code, n = base, 2
    taken = {t.casefold() for t in taken}
    while code.casefold() in taken:
        code, n = f"{base}{n}", n + 1
    return code


def library_for(tracker, designation=None):
    """(GateDef, owner) for the whole library: general gates (owner None = the master), then each subteam's."""
    out = [(g, None) for g in tracker.library]
    for t in tracker.subteams:
        out += [(g, t) for g in t.library]
    return [(g, o) for g, o in out if designation is None or g.designation in (GENERAL, designation)]


def ensure_library(owner, template, general_library=()):
    """Add any gate of `template` that isn't in the library yet to `owner`'s library (a Subteam, or a Tracker for general)."""
    known = {g.name.casefold() for g in list(owner.library) + list(general_library)}
    taken = {g.code for g in list(owner.library) + list(general_library)}
    for name in template.gates:
        if name.casefold() not in known:
            std = dict(zip(DEFAULT_GATES, DEFAULT_CODES)).get(name)
            code = std if std and std not in taken else auto_code(name, taken)
            taken.add(code)
            owner.library.append(GateDef(name, code, owner.name if isinstance(owner, Subteam) else GENERAL))
            known.add(name.casefold())


def all_templates(tracker):
    """(Template, owner) by name: general ones (owner None) and every subteam's own."""
    out = {n: (t, None) for n, t in tracker.templates.items()}
    for team in tracker.subteams:
        out.update({n: (t, team) for n, t in team.templates.items()})
    return out


def template_users(tracker, name):
    """How many parts follow this template."""
    return sum(1 for t in tracker.subteams for p in t.parts if template_of(t, p).name == name)


def gate_users(tracker, gate):
    """(names of templates that use the gate, how many parts follow those templates)."""
    names = [n for n, (t, _) in all_templates(tracker).items() if gate in t.gates]
    return names, sum(template_users(tracker, n) for n in names)


def code_of(tracker_or_team, gate):
    libs = list(getattr(tracker_or_team, "library", []))
    if isinstance(tracker_or_team, Tracker):
        libs += [g for t in tracker_or_team.subteams for g in t.library]
    return next((g.code for g in libs if g.name == gate and g.code), "")


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
    wb = load_workbook(master, data_only=True)
    ws = wb["Key Dates"]
    cols = _headers(ws)
    key_dates = []
    for row in ws.iter_rows(min_row=2, max_col=len(cols), values_only=True):
        cell = lambda name: row[cols.index(name)] if name in cols else None
        if _text(cell("Event")):
            key_dates.append(KeyDate(
                _date(cell("Date")), _text(cell("Event")), _text(cell("Type")) or "Deadline", split_names(_text(cell("People"))),
                _truthy(cell("Done")), _date(cell("Internal deadline")), _date(cell("End date")), _text(cell("Track")),
                _text(cell("Competition")), _text(cell("Notes"))))
    if TEMPLATE_SHEET in wb.sheetnames or LIBRARY_SHEET in wb.sheetnames:
        templates, library = _read_templates(wb, GENERAL), _read_library(wb, GENERAL)
    else:  # a master from before templates: the standard gates
        templates, library = {DEFAULT_TEMPLATE: standard_template()}, standard_library()
    tracker = Tracker(folder, key_dates, [], is_writable(master), master.stat().st_mtime,
                      _settings(wb).get("Car") or folder.name, library, templates)
    for p in sorted((folder / SUBTEAM_DIR).glob("*.xlsx")):
        if not p.name.startswith("~$"):
            team = load_subteam(p, templates, library)
            if team:
                tracker.subteams.append(team)
    return tracker


def load_subteam(path, shared=None, general_library=()):
    wb = load_workbook(path, data_only=True)
    if PLAN_SHEET not in wb.sheetnames:
        return None
    ws = wb[PLAN_SHEET]
    fixed, columns = _plan_layout(ws)
    renamed = dict(zip(OLD_DEFAULT_GATES, DEFAULT_GATES)) if columns == OLD_DEFAULT_GATES else {}
    columns = [renamed.get(g, g) for g in columns]
    team = Subteam(path.stem, path, [], is_writable(path), path.stat().st_mtime, shared=dict(shared or {}), columns=columns)
    if TEMPLATE_SHEET in wb.sheetnames:
        team.templates = _read_templates(wb, team.name)
    elif LEGACY_GATE_SHEET in wb.sheetnames:  # one rule set for the subteam, from before templates
        team.templates = {team.name: _legacy_template(wb, team.name, columns)}
    elif not (shared and template_matches(shared.get(DEFAULT_TEMPLATE), columns)):
        team.templates = {team.name: standard_template(columns, team.name, team.name)}  # its gates, in order
    team.library = _read_library(wb, team.name)
    team.default_template = _settings(wb).get("Default template", "")
    n = len(fixed)
    for row in ws.iter_rows(min_row=2, max_col=n + len(columns), values_only=True):
        cell = lambda col: _text(row[fixed.index(col)]) if col in fixed else ""
        if not cell("Part"):
            continue
        flag = cell("Flag") or "OK"
        cells = dict(zip(columns, row[n:]))
        team.parts.append(Part(
            name=cell("Part"), owner=cell("Owner"), current=cell("Current Gate"), flag=flag if flag in FLAGS else "OK",
            notes=cell("Notes"), dates={g: _date(v) for g, v in cells.items()}, parent=cell("Parent"),
            assignees=split_names(cell("Assignees")), tags=split_tags(cell("Tags")),
            skipped={g for g, v in cells.items() if _text(v).upper() in (SKIP, "NA", "SKIP", "OFF")},
            template=cell("Template"),
        ))
    team.actions = [Action(str(r[0]).strip(), _text(r[1]), _text(r[2]), _date(r[3]), _text(r[4]), _date(r[5]))
                    for r in _rows(wb, ACTION_SHEET, len(ACTION_COLS))]
    team.reviews = [Review(_text(r[0]), _text(r[1]), _date(r[2]), _text(r[3]), _text(r[4]), _text(r[5]))
                    for r in _rows(wb, REVIEW_SHEET, len(REVIEW_COLS)) if _text(r[1]) in DR_TYPES]
    by_key = {p.key: p for p in team.parts}
    for r in _rows(wb, LOG_SHEET, len(LOG_COLS)):
        part, gate = by_key.get(_text(r[0])), renamed.get(_text(r[1]), _text(r[1]))
        state = _text(r[4]) or DONE_STATE
        if part and gate in template_of(team, part).gates:
            if state == IN_PROGRESS:
                part.started.add(gate)
            else:
                part.passed[gate] = (_date(r[3]), _date(r[2]))
        else:
            team.unmatched_log.append(GatePass(_text(r[0]), gate, _date(r[2]), _date(r[3]), state))
    for p in team.parts:
        _passed_from_current(p, team, renamed)
    for t in team.templates.values():
        ensure_library(team, t, general_library)
    return team


def template_matches(template, columns):
    return template is not None and template.gates == list(columns)


def _legacy_template(wb, name, columns):
    after, due_before = {g: [] for g in columns}, {}
    for r in _rows(wb, LEGACY_GATE_SHEET, 3):
        g = _text(r[0])
        if g in columns:
            earlier = columns[:columns.index(g)]
            after[g] = [a.strip() for a in _text(r[1]).split(",") if a.strip() in earlier]
            if _text(r[2]):
                due_before[g] = _text(r[2])
    return Template(name, name, list(columns), after, due_before)


def _read_templates(wb, designation):
    out = {}
    for r in _rows(wb, TEMPLATE_SHEET, len(TEMPLATE_COLS)):
        name, gate = _text(r[0]), _text(r[2])
        if not gate:
            continue
        t = out.setdefault(name, Template(name, _text(r[1]) or designation))
        if gate not in t.gates:
            t.gates.append(gate)
            t.after[gate] = [a.strip() for a in _text(r[3]).split(",") if a.strip() in t.gates[:-1]]
            if _text(r[4]):
                t.due_before[gate] = _text(r[4])
    return out


def _read_library(wb, designation):
    return [GateDef(_text(r[1]), _text(r[0]), _text(r[2]) or designation, _text(r[3]))
            for r in _rows(wb, LIBRARY_SHEET, len(LIBRARY_COLS), key=1) if _text(r[1])]


def _settings(wb):
    if SETTINGS_SHEET not in wb.sheetnames:
        return {}
    return {_text(r[0]): _text(r[1]) for r in wb[SETTINGS_SHEET].iter_rows(min_row=1, max_col=2, values_only=True) if r and r[0]}


def _passed_from_current(part, team, renamed):
    """Gates done before the app logged them: everything a listed current gate comes after (or every gate, for Done).
    This is how files edited by hand, or made by older versions, say which gates are done."""
    t = template_of(team, part)
    listed = [renamed.get(g.strip(), g.strip()) for g in re.split(r"[,+]", part.current) if g.strip()]
    before = t.gates if DONE in listed else [a for g in listed if g in t.gates for a in ancestors(t, g)]
    for g in before:
        if g not in part.passed and g not in part.skipped:
            part.passed[g] = (None, part.dates.get(g))
            part.started.discard(g)


def _headers(ws):
    headers = [_text(c.value) for c in ws[1]]
    return headers[:headers.index("")] if "" in headers else headers


def _rows(wb, sheet, width, key=0):
    if sheet not in wb.sheetnames:
        return []
    return [list(r) + [None] * width for r in wb[sheet].iter_rows(min_row=2, max_col=width, values_only=True)
            if len(r) > key and r[key] not in (None, "")]


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


def _sheet(wb, name, cols, widths, note=None):
    """Get a side sheet, creating it with a header (and an explanatory note) if it doesn't exist."""
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    _header(ws, cols)
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    if note:
        ws.cell(1, len(cols) + 2, note).font = Font(italic=True, color="757575")
    ws.freeze_panes = "A2"
    return ws


def _ensure_side_sheets(wb):
    """Add the sheets a subteam file needs, and drop the pre-template Gates sheet."""
    _sheet(wb, ACTION_SHEET, ACTION_COLS, (48, 24, 18, 12, 12, 12))
    ws = _sheet(wb, LOG_SHEET, LOG_COLS, (34, 24, 12, 12, 12),
                "Written by the app: gates that are done or in progress. A blank Passed date: done before the app recorded it.")
    if ws.cell(1, 5).value != "Status":
        _header(ws, LOG_COLS)
    _sheet(wb, REVIEW_SHEET, REVIEW_COLS, (34, 18, 12, 18, 20, 50), "One row per assembly and review (DR1 Concept / DR2 Intermediate / DR3 Final).")
    _sheet(wb, TEMPLATE_SHEET, TEMPLATE_COLS, (26, 16, 30, 50, 28),
           "One row per gate, in order. After: gates (comma-separated, earlier in the template) that come first. "
           "Due before: a key date (its Event name in Master.xlsx).")
    _sheet(wb, LIBRARY_SHEET, LIBRARY_COLS, (10, 30, 16, 60))
    _sheet(wb, SETTINGS_SHEET, ["Setting", "Value"], (20, 30))
    if LEGACY_GATE_SHEET in wb.sheetnames:
        del wb[LEGACY_GATE_SHEET]


def _write_rows(ws, rows, width, date_cols):
    for row in ws.iter_rows(min_row=2, max_row=max(ws.max_row, 2), max_col=width):
        for c in row:
            c.value = None
    for r, values in enumerate(rows, 2):
        for i, v in enumerate(values, 1):
            c = ws.cell(r, i, v if v not in ("", None) else None)
            if i in date_cols:
                c.number_format = "yyyy-mm-dd"


def _write_settings(wb, values):
    ws = _sheet(wb, SETTINGS_SHEET, ["Setting", "Value"], (20, 30))
    _write_rows(ws, [[k, v] for k, v in values.items()], 2, set())


def _write_templates(wb, templates, designation):
    ws = _sheet(wb, TEMPLATE_SHEET, TEMPLATE_COLS, (26, 16, 30, 50, 28))
    _write_rows(ws, [[t.name, t.designation or designation, g, ", ".join(t.after.get(g, [])), t.due_before.get(g)]
                     for t in templates.values() for g in t.gates], len(TEMPLATE_COLS), set())


def _write_library(wb, library):
    ws = _sheet(wb, LIBRARY_SHEET, LIBRARY_COLS, (10, 30, 16, 60))
    _write_rows(ws, [[g.code, g.name, g.designation, g.description] for g in library], len(LIBRARY_COLS), set())


def _plan_sheet(wb, gates):
    """A fresh, empty Plan sheet for these gates, with headers, dropdowns and highlighting.
    Replaces the existing one, so changing the gates (or upgrading an older layout) keeps Excel's formatting right."""
    if PLAN_SHEET in wb.sheetnames:
        wb.remove(wb[PLAN_SHEET])
    ws = wb.create_sheet(PLAN_SHEET, 0)
    wb.active = 0
    n = len(FIXED_COLS)
    _header(ws, FIXED_COLS + gates, [HEADER_FILL] * n + [GATE_FILL] * len(gates))
    for letter, width in zip("ABCDEFGHI", (28, 20, 18, 26, 22, 22, 30, 10, 40)):
        ws.column_dimensions[letter].width = width
    col = lambda name: get_column_letter(FIXED_COLS.index(name) + 1)
    cur = col("Current Gate")
    first, last = get_column_letter(n + 1), get_column_letter(n + max(len(gates), 1))
    for i in range(len(gates)):
        letter = get_column_letter(n + 1 + i)
        ws.column_dimensions[letter].width = max(15, len(gates[i]) + 4)
        for r in range(2, FORMAT_ROWS + 2):
            ws[f"{letter}{r}"].number_format = "yyyy-mm-dd"
    area = f"{first}2:{last}{FORMAT_ROWS + 1}"
    # A gate listed in Current Gate: red when overdue, yellow otherwise. Gates that are off (N/A) and finished parts: grey.
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
    (no Parent, Assignees, Tags or Template) still load; the gates are the named columns after the last fixed one.
    A fixed column whose header was cleared is just ignored (and restored on the next save)."""
    headers = [_text(c.value) for c in ws[1]]
    n = max((i + 1 for i, h in enumerate(headers) if h in FIXED_COLS), default=0)
    fixed, gates = headers[:n], headers[n:]
    if fixed and fixed[0] == "":
        fixed[0] = "Part"  # column A is always the part name
    return fixed, gates[:gates.index("")] if "" in gates else gates


def create_subteam_file(path, gates=DEFAULT_GATES, template=None, default=""):
    """An empty subteam plan. Its parts follow the master's templates unless `template` (its own) is given."""
    wb = Workbook()
    wb.remove(wb.active)
    _plan_sheet(wb, list(template.gates if template else gates))
    _ensure_side_sheets(wb)
    if template:
        _write_templates(wb, {template.name: template}, template.designation)
        _write_library(wb, [GateDef(g, auto_code(g), template.designation) for g in template.gates])
    _write_settings(wb, {"Default template": default or (template.name if template else "")})
    _save_atomic(wb, Path(path))


def create_master(path, car="", library=None, templates=None, key_dates=()):
    wb = Workbook()
    ws = wb.active
    ws.title = "Key Dates"
    _key_date_sheet(ws)
    _write_rows(ws, [_key_date_row(k) for k in sorted(key_dates, key=lambda k: k.date or date.max)], len(KEY_DATE_COLS), {1, 6, 7})
    about = wb.create_sheet("Read Me")
    for line in [
        "Project Tracker: master file" + (f" for {car}" if car else ""),
        "",
        "Key Dates: competition deadlines, design reviews (gate meetings), milestones, freezes, blackouts and events.",
        "   Only the PM/captain edits this.",
        "   People = who's responsible (comma-separated). They see it under 'My work' in the app.",
        "   Done = Yes once a deadline is submitted or a milestone is met.",
        "   Internal deadline = the team's own date ahead of the official one; the app tracks against it.",
        "   End date = last day of a blackout or an event that spans several days.",
        "Gate Library / Gate Templates: the general gates and templates every sub-team can use (sub-teams keep their",
        "   own in their files). Edit them in the app's Gates page.",
        "All Parts / All Actions: rebuilt by the tracker app from the subteam files. Don't edit them here.",
        "",
        "Each subteam's plan is in Subteams/<name>.xlsx, and only that subteam's lead can edit it.",
        "In a subteam file, the gates are the columns after Notes, and each holds that gate's due date (or N/A when it's off).",
        "Template: which gate template the part follows (blank = the sub-team's default).",
        "Current Gate: the gate(s) the part is working on, comma-separated, or Done.",
    ]:
        about.append([line])
    about["A1"].font = Font(bold=True, size=14)
    about.column_dimensions["A"].width = 110
    templates = {DEFAULT_TEMPLATE: standard_template()} if templates is None else templates
    _write_library(wb, standard_library() if library is None else library)
    _write_templates(wb, templates, GENERAL)
    _write_settings(wb, {"Car": car or Path(path).parent.name})
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


def _key_date_row(k):
    return [k.date, k.event, k.type, ", ".join(k.people), "Yes" if k.done else None, k.internal, k.end, k.track, k.competition, k.notes]


def save_subteam(team):
    """Write the parts, actions, gate log, reviews, templates and library back to the file. Raises a SaveError subclass if it can't."""
    ensure_can_save(team.path, team.mtime)
    wb = load_workbook(team.path)
    _ensure_side_sheets(wb)
    gates = team.gates
    if _plan_layout(wb[PLAN_SHEET]) != (FIXED_COLS, gates):
        _plan_sheet(wb, gates)
    rows = []
    for p in team.parts:
        own = template_of(team, p).gates
        rows.append([p.name, p.parent, p.owner, ", ".join(p.assignees), ", ".join(p.tags), p.template,
                     current_gate(p, team, ", ", rollup=False), p.flag, p.notes] + [(SKIP if g in p.skipped else p.dates.get(g)) if g in own else None for g in gates])
    _write_rows(wb[PLAN_SHEET], rows, len(FIXED_COLS) + len(gates), set(range(len(FIXED_COLS) + 1, len(FIXED_COLS) + len(gates) + 1)))
    _write_rows(wb[ACTION_SHEET], [[a.text, a.part, a.owner, a.due, a.source, a.closed] for a in team.actions], len(ACTION_COLS), {4, 6})
    _write_rows(wb[LOG_SHEET], [[g.part, g.gate, g.due, g.passed, g.status] for g in team.log], len(LOG_COLS), {3, 4})
    _write_rows(wb[REVIEW_SHEET], [[r.part, r.review, r.date, r.outcome, r.presenter, r.notes] for r in team.reviews
                                   if r.date or r.outcome or r.presenter or r.notes], len(REVIEW_COLS), {3})
    _write_templates(wb, team.templates, team.name)
    _write_library(wb, team.library)
    _write_settings(wb, {"Default template": team.default_template})
    _save_atomic(wb, team.path)
    team.mtime = team.path.stat().st_mtime


def save_key_dates(tracker):
    path = tracker.folder / MASTER
    ensure_can_save(path, tracker.master_mtime)
    wb = load_workbook(path)
    ws = wb["Key Dates"]
    _key_date_sheet(ws)  # also upgrades a master made before the newer columns existed
    _write_rows(ws, [_key_date_row(k) for k in sorted(tracker.key_dates, key=lambda k: k.date or date.max)], len(KEY_DATE_COLS), {1, 6, 7})
    _save_atomic(wb, path)
    tracker.master_mtime = path.stat().st_mtime


def save_master_gates(tracker):
    """Write the general gate library and templates (and the car name) to the master."""
    path = tracker.folder / MASTER
    ensure_can_save(path, tracker.master_mtime)
    wb = load_workbook(path)
    _write_library(wb, tracker.library)
    _write_templates(wb, tracker.templates, GENERAL)
    _write_settings(wb, {**_settings(wb), "Car": tracker.car})
    _save_atomic(wb, path)
    tracker.master_mtime = path.stat().st_mtime


def write_rollup(tracker, today=None):
    """Regenerate Master.xlsx > All Parts and All Actions from every subteam file."""
    path = tracker.folder / MASTER
    ensure_can_save(path)  # regenerated from scratch, so a newer master isn't a conflict
    wb = load_workbook(path)
    stamp = f"Generated by the tracker app {datetime.now():%Y-%m-%d %H:%M}. Don't edit it here."
    sheets = [
        ("All Parts", ROLLUP_COLS, (24, 28, 20, 18, 26, 22, 22, 30, 11, 10, 9, 40), {9},
         [[t.name, p.name, p.parent, p.owner, ", ".join(p.assignees), ", ".join(p.tags), template_of(t, p).name,
           current_gate(p, t, ", "), due_date(p, t), status(p, t, today), p.flag, p.notes]
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
