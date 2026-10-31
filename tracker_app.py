"""Project Tracker: a Material Design front end over the SharePoint-synced Excel files.

    pip install -r requirements.txt
    python tracker_app.py                      # finds the synced folder, or asks for it
    python tracker_app.py --folder "C:/Users/me/OneDrive - Uni/Project Tracker"

The app never decides who can edit what. SharePoint does: files you can't edit
are synced read-only, and the app shows them as view-only.
"""

import argparse
import copy
import ctypes
import json
import os
import sys
import threading
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import flet as ft

import tracker_core as core

CONFIG = Path(os.environ.get("APPDATA", Path.home())) / "FSAETracker" / "config.json"
ASSETS = Path(__file__).resolve().parent / "assets"
LOOK_AHEAD_DAYS = 14
FIXED_VIEWS = ["overview", "mywork", "keydates"]
SUBTEAM_EMOJI = {  # matched against the start of the subteam's file name
    "aero": "🪽", "suspension": "🛞", "chassis": "🏎️", "powertrain": "🔋",
    "electrical": "⚡", "driverless": "🤖",
}

STATUS_COLOURS = {  # status -> (container colour, text colour)
    "Late": (ft.Colors.ERROR_CONTAINER, ft.Colors.ON_ERROR_CONTAINER),
    "Due soon": (ft.Colors.AMBER_100, ft.Colors.AMBER_900),
    "On track": (ft.Colors.GREEN_100, ft.Colors.GREEN_900),
    "Done": (ft.Colors.SURFACE_CONTAINER_HIGHEST, ft.Colors.ON_SURFACE_VARIANT),
    "Held": (ft.Colors.SURFACE_CONTAINER_HIGHEST, ft.Colors.ON_SURFACE_VARIANT),
    "No date": (ft.Colors.SURFACE_CONTAINER_HIGHEST, ft.Colors.ON_SURFACE_VARIANT),
    "At risk": (ft.Colors.AMBER_100, ft.Colors.AMBER_900),
    "Blocked": (ft.Colors.ERROR_CONTAINER, ft.Colors.ON_ERROR_CONTAINER),
    "You": (ft.Colors.PRIMARY_CONTAINER, ft.Colors.ON_PRIMARY_CONTAINER),
}
TYPE_ICONS = {"Deadline": ft.Icons.FLAG_OUTLINED, "Design Review": ft.Icons.GROUPS_OUTLINED, "Milestone": ft.Icons.EMOJI_EVENTS_OUTLINED}


def load_config():
    try:
        return json.loads(CONFIG.read_text())
    except (OSError, ValueError):
        return {}


def save_config(cfg):
    CONFIG.parent.mkdir(parents=True, exist_ok=True)
    CONFIG.write_text(json.dumps(cfg))


def fmt(d):
    return d.strftime("%a %d %b") if d else "—"


def countdown(d, today):
    if d is None:
        return "TBD"
    days = (d - today).days
    return "today" if days == 0 else (f"in {plural(days, 'day')}" if days > 0 else f"{plural(-days, 'day')} ago")


def pill(text, key=None):
    bg, fg = STATUS_COLOURS.get(key or text, STATUS_COLOURS["No date"])
    return ft.Container(
        ft.Text(text, size=12, weight=ft.FontWeight.W_600, color=fg),
        bgcolor=bg, border_radius=12, padding=ft.Padding.symmetric(horizontal=10, vertical=3),
    )


def pills(*items):
    """A row of pills; items are (text, key) or None."""
    return ft.Row([pill(*i) for i in items if i], spacing=6, tight=True)


def heading(text, sub=None):
    controls = [ft.Text(text, theme_style=ft.TextThemeStyle.HEADLINE_SMALL, weight=ft.FontWeight.W_600)]
    if sub:
        controls.append(ft.Text(sub, color=ft.Colors.ON_SURFACE_VARIANT))
    return ft.Column(controls, spacing=2)


def section(title, body, trailing=None):
    head = ft.Row([ft.Text(title, theme_style=ft.TextThemeStyle.TITLE_MEDIUM, weight=ft.FontWeight.W_600)] + ([trailing] if trailing else []),
                  alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
    return ft.Card(ft.Container(ft.Column([head, body], spacing=8), padding=16), variant=ft.CardVariant.OUTLINED)


def empty(text):
    return ft.Container(ft.Text(text, color=ft.Colors.ON_SURFACE_VARIANT, italic=True), padding=ft.Padding.symmetric(vertical=8))


def gate_dots(part, gates):
    """One dot per gate: filled = passed, ringed = current, faint = to come."""
    current = core.current_gate(part, gates)
    passed, _ = core.progress(part, gates)
    st = core.status(part, gates)
    dots = []
    for i, g in enumerate(gates):
        if i < passed:
            colour, border = ft.Colors.PRIMARY, None
        elif g == current:
            colour = STATUS_COLOURS.get(st, STATUS_COLOURS["No date"])[0]
            border = ft.Border.all(2, ft.Colors.PRIMARY)
        else:
            colour, border = ft.Colors.SURFACE_CONTAINER_HIGHEST, None
        due = part.dates.get(g)
        dots.append(ft.Container(width=14, height=14, border_radius=7, bgcolor=colour, border=border,
                                 tooltip=f"{g}: {fmt(due) if due else 'no date'}"))
    return ft.Row(dots, spacing=4)


def plural(n, word):
    return f"{n} {word}{'' if n == 1 else 's'}"


def lateness(entry):
    if not entry.due:
        return ""
    days = (entry.passed - entry.due).days
    return "on time" if days == 0 else f"{plural(abs(days), 'day')} {'late' if days > 0 else 'early'}"


class TrackerApp:
    def __init__(self, page, folder):
        self.page = page
        self.folder = Path(folder)
        self.tracker = None
        self.view = "overview"
        self.me = load_config().get("me", "")
        self.rail = ft.NavigationRail(
            label_type=ft.NavigationRailLabelType.ALL, min_width=96, group_alignment=-0.9, scrollable=True, on_change=self.on_nav,
            trailing=ft.Column([
                ft.IconButton(ft.Icons.REFRESH, tooltip="Reload files", on_click=lambda e: self.reload()),
                ft.IconButton(ft.Icons.DARK_MODE_OUTLINED, tooltip="Light / dark", on_click=self.toggle_theme),
            ]),
        )
        self.body = ft.Container(expand=True, padding=24)
        page.add(ft.Row([self.rail, ft.VerticalDivider(width=1), self.body], expand=True))
        self.reload()

    # ---- data

    def reload(self):
        try:
            self.tracker = core.load(self.folder)
        except Exception as e:  # noqa: BLE001 - shown to the user, who can fix the file
            self.toast(f"Couldn't read the tracker files: {e}", error=True)
            return
        self.rail.destinations = [
            ft.NavigationRailDestination(icon=ft.Icons.DASHBOARD_OUTLINED, selected_icon=ft.Icons.DASHBOARD, label="Overview"),
            ft.NavigationRailDestination(icon=ft.Icons.PERSON_OUTLINE, selected_icon=ft.Icons.PERSON, label="My work"),
            ft.NavigationRailDestination(icon=ft.Icons.EVENT_OUTLINED, selected_icon=ft.Icons.EVENT, label="Key dates"),
        ] + [
            ft.NavigationRailDestination(icon=ft.Text(emoji(t.name), size=22), label=_short(t.name))
            for t in self.tracker.subteams
        ]
        if self.view not in FIXED_VIEWS and self.view not in [t.name for t in self.tracker.subteams]:
            self.view = "overview"
        self.show(self.view)

    def team(self, name):
        return next(t for t in self.tracker.subteams if t.name == name)

    def known_names(self):
        names = {}
        for t in self.tracker.subteams:
            for n in [p.owner for p in t.parts] + [a.owner for a in t.actions]:
                if n:
                    names.setdefault(n.casefold(), n)
        for kd in self.tracker.key_dates:
            for n in kd.people:
                names.setdefault(n.casefold(), n)
        return sorted(names.values(), key=str.casefold)

    def publish_rollup(self):
        """Keep Master > All Parts / All Actions current whenever someone who can edit the master saves."""
        if self.tracker.master_writable:
            try:
                core.write_rollup(self.tracker)
            except core.SaveError:
                pass  # master open in Excel; it'll be refreshed on the next save

    # ---- navigation

    def on_nav(self, e):
        i = e.control.selected_index
        self.show(FIXED_VIEWS[i] if i < len(FIXED_VIEWS) else self.tracker.subteams[i - len(FIXED_VIEWS)].name)

    def show(self, view):
        self.view = view
        if view in FIXED_VIEWS:
            index = FIXED_VIEWS.index(view)
            content = {"overview": self.overview, "mywork": self.my_work, "keydates": self.key_dates_view}[view]()
        else:
            index = [t.name for t in self.tracker.subteams].index(view) + len(FIXED_VIEWS)
            content = self.subteam_view(self.team(view))
        self.rail.selected_index = index
        self.page.floating_action_button = content[1]
        self.body.content = content[0]
        self.page.update()

    def toggle_theme(self, e):
        dark = self.page.theme_mode != ft.ThemeMode.DARK
        self.page.theme_mode = ft.ThemeMode.DARK if dark else ft.ThemeMode.LIGHT
        cfg = load_config()
        cfg["dark"] = dark
        save_config(cfg)
        self.page.update()

    def toast(self, text, error=False):
        self.page.show_dialog(ft.SnackBar(
            ft.Text(text, color=ft.Colors.ON_ERROR_CONTAINER if error else None),
            bgcolor=ft.Colors.ERROR_CONTAINER if error else None, show_close_icon=True,
        ))

    # ---- shared list tiles

    def part_tile(self, team, part, today, show_team=True):
        index = team.parts.index(part)
        sub = f"{core.current_gate(part, team.gates)} due {fmt(core.due_date(part, team.gates))}"
        sub = (f"{team.name} · " if show_team else "") + sub + (f" · {part.owner}" if part.owner else "") + (f"\n{part.notes}" if part.notes else "")
        return ft.ListTile(
            leading=ft.Icon(ft.Icons.BLOCK if part.flag == "Blocked" else ft.Icons.PRECISION_MANUFACTURING_OUTLINED,
                            color=ft.Colors.ERROR if core.needs_attention(part, team.gates, today) else ft.Colors.PRIMARY),
            title=ft.Text(part.name), subtitle=ft.Text(sub),
            trailing=pills((part.flag, None) if part.flag != "OK" else None, (core.status(part, team.gates, today), None)),
            on_click=lambda e: self.edit_part(team, index),
        )

    def action_tile(self, team, action, today, show_team=True):
        index = team.actions.index(action)
        st = core.action_status(action, today)
        bits = ([team.name] if show_team else []) + [action.part or "General", action.owner or "no owner", f"due {fmt(action.due)}"]
        if action.source:
            bits.append(f"from {action.source}")
        return ft.ListTile(
            leading=ft.Checkbox(value=action.closed is not None, disabled=not team.writable,
                                tooltip="Mark done" if team.writable else None,
                                on_change=lambda e: self.close_action(team, index, e.control.value)),
            title=ft.Text(action.text, color=ft.Colors.ON_SURFACE_VARIANT if action.closed else None),
            subtitle=ft.Text(" · ".join(bits)),
            trailing=pills((st, None)),
            on_click=lambda e: self.edit_action(team, index),
        )

    def key_date_tile(self, kd, today):
        st = core.key_date_status(kd, today)
        sub = f"{fmt(kd.date)} · {kd.type}" + (f" · {', '.join(kd.people)}" if kd.people else "")
        mine = any(core.is_me(n, self.me) for n in kd.people)
        label = "Done" if kd.done else ("Held" if st == "Held" else countdown(kd.date, today))
        index = self.tracker.key_dates.index(kd)
        return ft.ListTile(
            leading=ft.Icon(TYPE_ICONS.get(kd.type, ft.Icons.EVENT),
                            color=ft.Colors.ERROR if st == "Late" else (ft.Colors.OUTLINE if st in ("Done", "Held") else ft.Colors.PRIMARY)),
            title=ft.Text(kd.event, color=ft.Colors.ON_SURFACE_VARIANT if st in ("Done", "Held") else None),
            subtitle=ft.Text(sub),
            trailing=pills(("You", "You") if mine else None, (label, st)),
            on_click=(lambda e: self.edit_key_date(index)) if self.tracker.master_writable else None,
        )

    # ---- overview

    def overview(self):
        today = date.today()
        soon = today + timedelta(days=LOOK_AHEAD_DAYS)
        cards, attention, upcoming = [], [], []
        for t in self.tracker.subteams:
            stats = {s: 0 for s in ("Late", "Due soon", "Done")}
            flagged = 0
            for p in t.parts:
                st = core.status(p, t.gates, today)
                if st in stats:
                    stats[st] += 1
                flagged += p.flag != "OK"
                if core.needs_attention(p, t.gates, today):
                    attention.append((core.due_date(p, t.gates) or date.max, self.part_tile(t, p, today)))
                due = core.due_date(p, t.gates)
                if due and today <= due <= soon:
                    upcoming.append((due, ft.Icons.PRECISION_MANUFACTURING_OUTLINED, f"{p.name}: {core.current_gate(p, t.gates)}", t.name, t.name))
            late_actions = 0
            for a in t.actions:
                st = core.action_status(a, today)
                if st == "Late":
                    late_actions += 1
                    attention.append((a.due, self.action_tile(t, a, today)))
                elif a.closed is None and a.due and today <= a.due <= soon:
                    upcoming.append((a.due, ft.Icons.TASK_ALT, a.text, f"Action · {t.name}" + (f" · {a.owner}" if a.owner else ""), t.name))
            passed = sum(core.progress(p, t.gates)[0] for p in t.parts)
            total = sum(core.progress(p, t.gates)[1] for p in t.parts) or 1
            chips = [(f"{stats['Late']} late", "Late") if stats["Late"] else None,
                     (f"{late_actions} action{'s' if late_actions > 1 else ''} overdue", "Late") if late_actions else None,
                     (f"{flagged} flagged", "At risk") if flagged else None,
                     (f"{stats['Due soon']} due soon", "Due soon") if stats["Due soon"] else None]
            chips = [pill(*c) for c in chips if c] or [pill("All on track", "On track") if t.parts else pill("No parts yet", "No date")]
            cards.append(ft.Container(
                ft.Card(ft.Container(ft.Column([
                    ft.Row([ft.Text(f"{emoji(t.name)}  {t.name}", weight=ft.FontWeight.W_600, size=16, expand=True),
                            ft.Icon(ft.Icons.EDIT_OUTLINED if t.writable else ft.Icons.LOCK_OUTLINE, size=16,
                                    color=ft.Colors.ON_SURFACE_VARIANT, tooltip="You can edit" if t.writable else "View only")]),
                    ft.Text(f"{plural(len(t.parts), 'part')} · {stats['Done']} done · {plural(sum(a.closed is None for a in t.actions), 'open action')}",
                            color=ft.Colors.ON_SURFACE_VARIANT, size=13),
                    ft.ProgressBar(value=passed / total, bar_height=6, border_radius=3),
                    ft.Row(chips, wrap=True, spacing=6),
                ], spacing=10), padding=16, ink=True, on_click=lambda e, n=t.name: self.show(n)), variant=ft.CardVariant.FILLED),
                col={"xs": 12, "sm": 6, "lg": 4},
            ))
        for kd in self.tracker.key_dates:
            if core.key_date_needs_attention(kd, today):
                attention.append((kd.date, self.key_date_tile(kd, today)))
            elif kd.date and today <= kd.date <= soon and not kd.done:
                upcoming.append((kd.date, TYPE_ICONS.get(kd.type, ft.Icons.EVENT), kd.event,
                                 kd.type + (f" · {', '.join(kd.people)}" if kd.people else ""), "keydates"))
        upcoming.sort(key=lambda u: u[0])
        attention.sort(key=lambda a: a[0])

        next_review = min((k for k in self.tracker.key_dates if k.type == "Design Review" and k.date and k.date >= today),
                          key=lambda k: k.date, default=None)
        sub = f"{today:%A %d %B %Y}"
        if next_review:
            sub += f"  ·  Next design review: {next_review.event}, {fmt(next_review.date)}"
        upcoming_rows = [
            ft.ListTile(leading=ft.Icon(icon, color=ft.Colors.PRIMARY), title=ft.Text(title),
                        subtitle=ft.Text(f"{fmt(d)} · {sub_}"), dense=True, on_click=lambda e, v=view: self.show(v))
            for d, icon, title, sub_, view in upcoming
        ]
        content = ft.Column([
            heading("Overview", sub),
            ft.ResponsiveRow(cards, spacing=12, run_spacing=12),
            ft.ResponsiveRow([
                ft.Container(section(f"Needs attention ({len(attention)})",
                                     ft.Column([a[1] for a in attention]) if attention else empty("Nothing late, flagged or due soon.")),
                             col={"xs": 12, "lg": 7}),
                ft.Container(section(f"Next {LOOK_AHEAD_DAYS} days",
                                     ft.Column(upcoming_rows) if upcoming_rows else empty("Nothing due.")),
                             col={"xs": 12, "lg": 5}),
            ], spacing=12, run_spacing=12),
        ], spacing=20, scroll=ft.ScrollMode.AUTO)
        return content, None

    # ---- my work

    def my_work(self):
        today = date.today()
        names = self.known_names()

        def pick(e):
            self.me = e.control.value or ""
            cfg = load_config()
            cfg["me"] = self.me
            save_config(cfg)
            self.show("mywork")

        who = ft.Dropdown(label="I am…", value=self.me if self.me in names else None, width=320, enable_filter=True,
                          editable=True, options=[ft.DropdownOption(key=n, text=n) for n in names], on_select=pick)
        header = ft.Row([heading("My work", "Your parts, actions and the key dates you're responsible for."), who],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN, vertical_alignment=ft.CrossAxisAlignment.START)
        if not self.me:
            return ft.Column([header, section("Get started", empty(
                "Pick your name above. Names come from part owners, action owners and key-date people."))], spacing=20), None

        parts = [(t, p) for t in self.tracker.subteams for p in t.parts
                 if core.is_me(p.owner, self.me) and core.current_gate(p, t.gates) != core.DONE]
        parts.sort(key=lambda tp: (not core.needs_attention(tp[1], tp[0].gates, today), core.due_date(tp[1], tp[0].gates) or date.max))
        actions = [(t, a) for t in self.tracker.subteams for a in t.actions if core.is_me(a.owner, self.me) and a.closed is None]
        actions.sort(key=lambda ta: ta[1].due or date.max)
        dates = [k for k in self.tracker.key_dates
                 if any(core.is_me(n, self.me) for n in k.people) and core.key_date_status(k, today) not in ("Done", "Held")]
        dates.sort(key=lambda k: k.date or date.max)
        late = sum(core.needs_attention(p, t.gates, today) for t, p in parts) + \
            sum(core.action_status(a, today) == "Late" for _, a in actions) + \
            sum(core.key_date_needs_attention(k, today) for k in dates)
        summary = pill(f"{late} need attention", "Late") if late else pill("Nothing late", "On track")
        return ft.Column([
            header,
            ft.Row([summary]),
            section(f"My parts ({len(parts)})", ft.Column([self.part_tile(t, p, today) for t, p in parts]) if parts else empty("No open parts.")),
            section(f"My actions ({len(actions)})", ft.Column([self.action_tile(t, a, today) for t, a in actions]) if actions else empty("No open actions.")),
            section(f"My key dates ({len(dates)})", ft.Column([self.key_date_tile(k, today) for k in dates]) if dates else empty("No upcoming key dates.")),
        ], spacing=20, scroll=ft.ScrollMode.AUTO), None

    # ---- subteam

    def subteam_view(self, team):
        today = date.today()
        bg, fg = (ft.Colors.PRIMARY_CONTAINER, ft.Colors.ON_PRIMARY_CONTAINER) if team.writable else STATUS_COLOURS["No date"]
        lock = ft.Container(
            ft.Row([ft.Icon(ft.Icons.EDIT_OUTLINED if team.writable else ft.Icons.LOCK_OUTLINE, size=16, color=fg),
                    ft.Text("You can edit" if team.writable else "View only", color=fg, weight=ft.FontWeight.W_600)],
                   spacing=6, tight=True),
            bgcolor=bg, border_radius=16, padding=ft.Padding.symmetric(horizontal=14, vertical=8),
        )
        sub = " → ".join(team.gates)
        hit, total = core.on_time(team)
        if total:
            sub += f"   ·   Gates passed on time: {hit} of {total}"
        header = ft.Row([heading(f"{emoji(team.name)}  {team.name}", sub), lock], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)

        table = ft.DataTable(
            columns=[ft.DataColumn(ft.Text(h)) for h in ("Part", "Owner", "Gate", "Due", "Status", "Notes")],
            rows=self.part_rows(team, "", today), show_checkbox_column=False, data_row_min_height=56,
            data_row_max_height=72, column_spacing=28,
        )

        def filter_parts(e):
            table.rows = self.part_rows(team, e.control.value, today)
            table.update()

        search = ft.TextField(hint_text="Filter by part, owner or note", prefix_icon=ft.Icons.SEARCH, dense=True, width=300,
                              on_change=filter_parts)
        parts_body = ft.Row([table], scroll=ft.ScrollMode.AUTO) if team.parts else empty(
            "No parts yet." + (" Press + to add one." if team.writable else ""))

        open_actions = sorted((a for a in team.actions if a.closed is None), key=lambda a: a.due or date.max)
        closed = sorted((a for a in team.actions if a.closed), key=lambda a: a.closed, reverse=True)[:5]
        action_tiles = [self.action_tile(team, a, today, show_team=False) for a in open_actions + closed]
        add_action = ft.TextButton("Add action", icon=ft.Icons.ADD_TASK, on_click=lambda e: self.edit_action(team, None)) if team.writable else None
        fab = ft.FloatingActionButton(icon=ft.Icons.ADD, content="Add part", on_click=lambda e: self.edit_part(team, None)) if team.writable else None
        return ft.Column([
            header,
            section("Parts", parts_body, search if team.parts else None),
            section(f"Actions ({len(open_actions)} open)", ft.Column(action_tiles) if action_tiles else empty(
                "No actions. Add follow-ups from design reviews here."), add_action),
        ], spacing=20, scroll=ft.ScrollMode.AUTO), fab

    def part_rows(self, team, query, today):
        order = {"Late": 0, "Due soon": 1, "On track": 2, "No date": 3, "Done": 4}
        q = (query or "").strip().casefold()
        rows = []
        for idx in sorted(range(len(team.parts)),
                          key=lambda i: (team.parts[i].flag == "OK", order[core.status(team.parts[i], team.gates, today)],
                                         core.due_date(team.parts[i], team.gates) or date.max)):
            p = team.parts[idx]
            if q and q not in f"{p.name} {p.owner} {p.notes}".casefold():
                continue
            st = core.status(p, team.gates, today)
            rows.append(ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(p.name, weight=ft.FontWeight.W_500)),
                    ft.DataCell(ft.Text(p.owner or "—")),
                    ft.DataCell(ft.Column([ft.Text(core.current_gate(p, team.gates), size=13), gate_dots(p, team.gates)],
                                          spacing=4, tight=True, alignment=ft.MainAxisAlignment.CENTER)),
                    ft.DataCell(ft.Text(fmt(core.due_date(p, team.gates)))),
                    ft.DataCell(pills((st, None), (p.flag, None) if p.flag != "OK" else None)),
                    ft.DataCell(ft.Text(p.notes, max_lines=2, overflow=ft.TextOverflow.ELLIPSIS, width=260)),
                ],
                on_select_change=lambda e, i=idx: self.edit_part(team, i),
            ))
        return rows

    def edit_part(self, team, index):
        part = team.parts[index] if index is not None else core.Part(current=team.gates[0] if team.gates else core.DONE)
        ro = not team.writable
        name = ft.TextField(label="Part", value=part.name, read_only=ro, autofocus=index is None)
        owner = ft.TextField(label="Owner", value=part.owner, read_only=ro)
        gate = ft.Dropdown(label="Current gate", value=core.current_gate(part, team.gates), disabled=ro, expand=True,
                           options=[ft.DropdownOption(key=g, text=g) for g in team.gates + [core.DONE]])
        flag = ft.SegmentedButton(selected=[part.flag], disabled=ro, show_selected_icon=False,
                                  segments=[ft.Segment(value=f, label=ft.Text(f)) for f in core.FLAGS])
        notes = ft.TextField(label="Notes", value=part.notes, multiline=True, min_lines=2, max_lines=4, read_only=ro)
        date_fields = {g: date_field(self.page, g, part.dates.get(g), ro) for g in team.gates}

        def advance(e):
            g = gate.value
            if g in team.gates:
                i = team.gates.index(g)
                gate.value = team.gates[i + 1] if i + 1 < len(team.gates) else core.DONE
                gate.update()

        def save(e):
            if not name.value.strip():
                name.error = "Give the part a name"
                name.update()
                return
            try:
                dates = {g: f.parse() for g, f in date_fields.items()}
            except ValueError as err:
                self.toast(str(err), error=True)
                return
            new = core.Part(name.value.strip(), owner.value.strip(), gate.value or "", (flag.selected or ["OK"])[0], notes.value.strip(), dates)

            def change():
                if index is None:
                    team.parts.append(new)
                else:
                    core.record_part_change(team, team.parts[index], new)
                    team.parts[index] = new
            self.commit(team, change)

        def delete(e):
            self.commit(team, lambda: team.parts.pop(index))

        actions = [ft.TextButton("Close" if ro else "Cancel", on_click=lambda e: self.page.pop_dialog())]
        if not ro:
            if index is not None:
                actions.insert(0, ft.TextButton("Delete", icon=ft.Icons.DELETE_OUTLINE, on_click=delete,
                                                style=ft.ButtonStyle(color=ft.Colors.ERROR)))
            actions.append(ft.FilledButton("Save", on_click=save))
        gate_row = [gate] + ([ft.FilledTonalButton("Gate passed", icon=ft.Icons.CHECK_CIRCLE, on_click=advance)] if not ro else [])
        history = [e for e in team.log if e.part == part.name] if index is not None else []
        history.sort(key=lambda e: team.gates.index(e.gate) if e.gate in team.gates else 99)
        part_actions = [a for a in team.actions if a.part == part.name and a.closed is None] if index is not None else []
        extra = []
        if history:
            extra += [ft.Text("Gate history", theme_style=ft.TextThemeStyle.TITLE_SMALL)] + [
                ft.Text(f"✓ {e.gate}: passed {fmt(e.passed)}" + (f" ({lateness(e)})" if e.due else ""), size=13,
                        color=ft.Colors.ERROR if e.due and e.passed > e.due else None)
                for e in history]
        if part_actions:
            extra += [ft.Text("Open actions", theme_style=ft.TextThemeStyle.TITLE_SMALL)] + [
                ft.Text(f"• {a.text} ({a.owner or 'no owner'}, due {fmt(a.due)})", size=13) for a in part_actions]
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(("Edit part" if index is not None else "New part") if not ro else part.name),
            content=ft.Container(ft.Column([
                name, owner,
                ft.Row(gate_row, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Text("Flag", size=12, color=ft.Colors.ON_SURFACE_VARIANT), flag,
                notes,
                *extra,
                ft.Text("Gate due dates", theme_style=ft.TextThemeStyle.TITLE_SMALL),
                *[f.control for f in date_fields.values()],
            ], tight=True, spacing=12, scroll=ft.ScrollMode.AUTO, horizontal_alignment=ft.CrossAxisAlignment.STRETCH), width=460),
            actions=actions,
        ))

    def edit_action(self, team, index):
        action = team.actions[index] if index is not None else core.Action()
        ro = not team.writable
        text = ft.TextField(label="Action", value=action.text, read_only=ro, autofocus=index is None, multiline=True, max_lines=3)
        part = ft.Dropdown(label="Part", value=action.part or None, disabled=ro,
                           options=[ft.DropdownOption(key=p.name, text=p.name) for p in team.parts])
        owner = ft.TextField(label="Owner", value=action.owner, read_only=ro)
        due = date_field(self.page, "Due", action.due, ro)
        source = ft.TextField(label="From (e.g. DR2)", value=action.source, read_only=ro)
        done = ft.Switch(label="Done", value=action.closed is not None, disabled=ro)

        def save(e):
            try:
                d = due.parse()
            except ValueError as err:
                self.toast(str(err), error=True)
                return
            if not text.value.strip():
                text.error = "Describe the action"
                text.update()
                return
            closed = (action.closed or date.today()) if done.value else None
            new = core.Action(text.value.strip(), part.value or "", owner.value.strip(), d, source.value.strip(), closed)
            self.commit(team, lambda: team.actions.__setitem__(index, new) if index is not None else team.actions.append(new))

        buttons = [ft.TextButton("Close" if ro else "Cancel", on_click=lambda e: self.page.pop_dialog())]
        if not ro:
            if index is not None:
                buttons.insert(0, ft.TextButton("Delete", icon=ft.Icons.DELETE_OUTLINE, style=ft.ButtonStyle(color=ft.Colors.ERROR),
                                                on_click=lambda e: self.commit(team, lambda: team.actions.pop(index))))
            buttons.append(ft.FilledButton("Save", on_click=save))
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(("Edit action" if index is not None else "New action") if not ro else "Action"),
            content=ft.Container(ft.Column([text, part, owner, due.control, source, done], tight=True, spacing=12,
                                           horizontal_alignment=ft.CrossAxisAlignment.STRETCH), width=460),
            actions=buttons,
        ))

    def close_action(self, team, index, done):
        def change():
            team.actions[index].closed = date.today() if done else None
        self.commit(team, change, dialog=False)

    def commit(self, team, change, dialog=True):
        """Apply a change to a subteam and save it; on failure undo it (and reload if someone else saved first)."""
        backup = copy.deepcopy((team.parts, team.actions, team.log))
        change()
        try:
            core.save_subteam(team)
        except core.SaveError as e:
            team.parts, team.actions, team.log = backup
            if dialog:
                self.page.pop_dialog()
            self.toast(str(e), error=True)
            if isinstance(e, core.ChangedOnDiskError):
                self.reload()
            else:
                self.show(self.view)
            return
        if dialog:
            self.page.pop_dialog()
        self.publish_rollup()
        self.toast(f"Saved to {team.path.name}")
        self.show(self.view)

    # ---- key dates

    def key_dates_view(self):
        today = date.today()
        can_edit = self.tracker.master_writable
        show_past = load_config().get("show_past", False)
        tiles = [self.key_date_tile(kd, today) for kd in sorted(self.tracker.key_dates, key=lambda k: k.date or date.max)
                 if show_past or not (kd.date and kd.date < today) or core.key_date_status(kd, today) == "Late"]

        def toggle_past(e):
            cfg = load_config()
            cfg["show_past"] = e.control.value
            save_config(cfg)
            self.show("keydates")

        sub = "Competition deadlines, design reviews and milestones, and who's responsible for each." + (
            "" if can_edit else "  View only: the PM/captain edits these.")
        content = ft.Column([
            heading("Key dates", sub),
            section("Dates", ft.Column(tiles) if tiles else empty("No upcoming dates."),
                    ft.Switch(label="Show past", value=show_past, on_change=toggle_past)),
        ], spacing=20, scroll=ft.ScrollMode.AUTO)
        fab = ft.FloatingActionButton(icon=ft.Icons.ADD, content="Add date", on_click=lambda e: self.edit_key_date(None)) if can_edit else None
        return content, fab

    def edit_key_date(self, index):
        kd = self.tracker.key_dates[index] if index is not None else core.KeyDate()
        event = ft.TextField(label="Event", value=kd.event, autofocus=True)
        when = date_field(self.page, "Date", kd.date, False)
        kind = ft.SegmentedButton(selected=[kd.type if kd.type in core.KEY_DATE_TYPES else "Deadline"], show_selected_icon=False,
                                  segments=[ft.Segment(value=t, label=ft.Text(t), icon=ft.Icon(TYPE_ICONS[t])) for t in core.KEY_DATE_TYPES])
        people = ft.TextField(label="People responsible", value=", ".join(kd.people), hint_text="e.g. Sam, Alex")
        done = ft.Switch(label="Done (submitted / met)", value=kd.done)

        def save(e):
            try:
                d = when.parse()
            except ValueError as err:
                self.toast(str(err), error=True)
                return
            if not event.value.strip() or d is None:
                self.toast("An event needs a name and a date.", error=True)
                return
            new = core.KeyDate(d, event.value.strip(), kind.selected[0], core.split_names(people.value), done.value)
            self.commit_key_dates(lambda: self.tracker.key_dates.__setitem__(index, new) if index is not None
                                  else self.tracker.key_dates.append(new))

        actions = [ft.TextButton("Cancel", on_click=lambda e: self.page.pop_dialog()), ft.FilledButton("Save", on_click=save)]
        if index is not None:
            actions.insert(0, ft.TextButton("Delete", icon=ft.Icons.DELETE_OUTLINE, style=ft.ButtonStyle(color=ft.Colors.ERROR),
                                            on_click=lambda e: self.commit_key_dates(lambda: self.tracker.key_dates.pop(index))))
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text("Edit date" if index is not None else "New date"),
            content=ft.Container(ft.Column([event, when.control, kind, people, done], tight=True, spacing=12,
                                           horizontal_alignment=ft.CrossAxisAlignment.STRETCH), width=520),
            actions=actions,
        ))

    def commit_key_dates(self, change):
        backup = copy.deepcopy(self.tracker.key_dates)
        change()
        try:
            core.save_key_dates(self.tracker)
        except core.SaveError as e:
            self.tracker.key_dates = backup
            self.page.pop_dialog()
            self.toast(str(e), error=True)
            if isinstance(e, core.ChangedOnDiskError):
                self.reload()
            return
        self.page.pop_dialog()
        self.toast("Key dates saved")
        self.show(self.view)


class date_field:
    """A text field (yyyy-mm-dd) with a calendar button."""

    def __init__(self, page, label, value, read_only):
        self.field = ft.TextField(label=label, value=value.isoformat() if value else "", hint_text="yyyy-mm-dd",
                                  read_only=read_only, expand=True, dense=True)
        picker = ft.DatePicker(value=datetime.combine(value or date.today(), datetime.min.time()),
                               first_date=datetime(2020, 1, 1), last_date=datetime(2035, 12, 31), on_change=self._picked)
        button = ft.IconButton(ft.Icons.CALENDAR_MONTH, tooltip="Pick a date", disabled=read_only,
                               on_click=lambda e: page.show_dialog(picker))
        self.control = ft.Row([self.field, button], spacing=4)

    def _picked(self, e):
        v = e.control.value
        self.field.value = (v.date() if isinstance(v, datetime) else v).isoformat()
        self.field.update()

    def parse(self):
        text = (self.field.value or "").strip()
        if not text:
            return None
        try:
            return date.fromisoformat(text)
        except ValueError:
            raise ValueError(f"{self.field.label}: '{text}' isn't a date. Use yyyy-mm-dd.") from None


def sharpen_window_icon(title, icon_path, timeout=20):
    """Flet loads the 16 px icon and Windows stretches it on high-DPI screens, which looks blurry.
    Once the window exists, hand Windows the .ico entries that match the screen's DPI instead."""
    if sys.platform != "win32":
        return
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    user32.FindWindowExW.restype = wintypes.HWND
    user32.FindWindowExW.argtypes = [wintypes.HWND, wintypes.HWND, wintypes.LPCWSTR, wintypes.LPCWSTR]
    user32.GetDpiForWindow.argtypes = [wintypes.HWND]
    user32.LoadImageW.restype = wintypes.HANDLE
    user32.LoadImageW.argtypes = [wintypes.HINSTANCE, wintypes.LPCWSTR, wintypes.UINT, ctypes.c_int, ctypes.c_int, wintypes.UINT]
    user32.SendMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    IMAGE_ICON, LR_LOADFROMFILE, WM_SETICON, ICON_SMALL, ICON_BIG = 1, 0x10, 0x80, 0, 1

    deadline = time.time() + timeout
    applied = 0
    while time.time() < deadline and applied < 3:  # re-apply a few times in case Flet sets its own icon late
        hwnd, found = None, False
        while hwnd := user32.FindWindowExW(None, hwnd, None, title):  # every app window with this title
            found = True
            scale = (user32.GetDpiForWindow(hwnd) or 96) / 96
            for kind, base in ((ICON_SMALL, 16), (ICON_BIG, 32)):
                px = round(base * scale)
                icon = user32.LoadImageW(None, str(icon_path), IMAGE_ICON, px, px, LR_LOADFROMFILE)
                if icon:
                    user32.SendMessageW(hwnd, WM_SETICON, kind, icon)
        applied += found
        time.sleep(0.5)


def emoji(name):
    key = name.lower()
    return next((e for k, e in SUBTEAM_EMOJI.items() if key.startswith(k)), "🔧")


def _short(name):
    return name if len(name) <= 12 else name.split()[0]


def setup_screen(page, on_open, message=None):
    path = ft.TextField(label="Synced 'Project Tracker' folder", expand=True, value=load_config().get("folder", ""))
    picker = ft.FilePicker()

    async def browse(e):
        chosen = await picker.get_directory_path(dialog_title="Choose the synced Project Tracker folder")
        if chosen:
            path.value = chosen
            path.update()

    def open_folder(e):
        folder = Path(path.value.strip().strip('"'))
        if not (folder / core.MASTER).exists():
            path.error = f"No {core.MASTER} in that folder"
            path.update()
            return
        on_open(folder)

    page.add(ft.Container(ft.Column([
        ft.Icon(ft.Icons.FOLDER_OPEN, size=48, color=ft.Colors.PRIMARY),
        ft.Text("Connect to the tracker", theme_style=ft.TextThemeStyle.HEADLINE_SMALL, weight=ft.FontWeight.W_600),
        ft.Text(message or "In SharePoint, open the Project Tracker library and click Sync (or 'Add shortcut to My files'). "
                           "Then choose that folder here; it's usually under 'OneDrive - <university>'.",
                color=ft.Colors.ON_SURFACE_VARIANT),
        ft.Row([path, ft.IconButton(ft.Icons.FOLDER_OPEN, tooltip="Browse", on_click=browse)]),
        ft.FilledButton("Open", icon=ft.Icons.ARROW_FORWARD, on_click=open_folder),
    ], spacing=16, horizontal_alignment=ft.CrossAxisAlignment.START), width=560, padding=32))


def main(page: ft.Page, folder=None, start=None):
    page.title = "Project Tracker"
    page.theme = ft.Theme(color_scheme_seed=ft.Colors.INDIGO, use_material3=True)
    page.dark_theme = ft.Theme(color_scheme_seed=ft.Colors.INDIGO, use_material3=True)
    cfg = load_config()
    page.theme_mode = ft.ThemeMode.DARK if cfg.get("dark") else ft.ThemeMode.LIGHT
    page.padding = 0
    page.window.icon = str(ASSETS / "icon.ico")
    threading.Thread(target=sharpen_window_icon, args=(page.title, ASSETS / "icon.ico"), daemon=True).start()
    page.window.width, page.window.height = 1280, 820
    page.window.min_width, page.window.min_height = 720, 520
    page.window.alignment = ft.Alignment.CENTER

    def open_tracker(path):
        cfg = load_config()
        cfg["folder"] = str(path)
        save_config(cfg)
        page.controls.clear()
        app = TrackerApp(page, path)
        if start:
            app.show(start)

    folder = folder or cfg.get("folder")
    if folder and (Path(folder) / core.MASTER).exists():
        open_tracker(Path(folder))
        return
    found = core.find_folder()
    if found:
        open_tracker(found)
        return
    setup_screen(page, open_tracker, "Couldn't find the tracker folder you used last time." if folder else None)


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Project Tracker")
    ap.add_argument("--folder", help="the synced Project Tracker folder (remembered after the first run)")
    ap.add_argument("--start", help="open on this view: overview, mywork, keydates, or a subteam name")
    args = ap.parse_args()
    ft.run(lambda page: main(page, args.folder, args.start), assets_dir=str(ASSETS))
