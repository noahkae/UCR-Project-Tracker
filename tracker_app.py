"""Project Tracker: a Material Design front end over the SharePoint-synced Excel files.

    pip install -r requirements.txt
    python tracker_app.py                      # finds the synced folder, or asks for it
    python tracker_app.py --folder "C:/Users/me/OneDrive - Uni/Project Tracker"

The app never decides who can edit what. SharePoint does: files you can't edit
are synced read-only, and the app shows them as view-only.
"""

import argparse
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
SUBTEAM_EMOJI = {  # matched against the start of the subteam's file name
    "aero": "🪽", "suspension": "🛞", "chassis": "🏎️", "powertrain": "🔋",
    "electrical": "⚡", "driverless": "🤖",
}

STATUS_COLOURS = {  # status -> (container colour, text colour)
    "Late": (ft.Colors.ERROR_CONTAINER, ft.Colors.ON_ERROR_CONTAINER),
    "Due soon": (ft.Colors.AMBER_100, ft.Colors.AMBER_900),
    "On track": (ft.Colors.GREEN_100, ft.Colors.GREEN_900),
    "Done": (ft.Colors.SURFACE_CONTAINER_HIGHEST, ft.Colors.ON_SURFACE_VARIANT),
    "No date": (ft.Colors.SURFACE_CONTAINER_HIGHEST, ft.Colors.ON_SURFACE_VARIANT),
    "At risk": (ft.Colors.AMBER_100, ft.Colors.AMBER_900),
    "Blocked": (ft.Colors.ERROR_CONTAINER, ft.Colors.ON_ERROR_CONTAINER),
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


def pill(text, key=None):
    bg, fg = STATUS_COLOURS.get(key or text, STATUS_COLOURS["No date"])
    return ft.Container(
        ft.Text(text, size=12, weight=ft.FontWeight.W_600, color=fg),
        bgcolor=bg, border_radius=12, padding=ft.Padding.symmetric(horizontal=10, vertical=3),
    )


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


class TrackerApp:
    def __init__(self, page, folder):
        self.page = page
        self.folder = Path(folder)
        self.tracker = None
        self.view = "overview"
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

    def reload(self, keep_view=True):
        try:
            self.tracker = core.load(self.folder)
        except Exception as e:  # noqa: BLE001 - shown to the user, who can fix the file
            self.toast(f"Couldn't read the tracker files: {e}", error=True)
            return
        names = [t.name for t in self.tracker.subteams]
        self.rail.destinations = [
            ft.NavigationRailDestination(icon=ft.Icons.DASHBOARD_OUTLINED, selected_icon=ft.Icons.DASHBOARD, label="Overview"),
            ft.NavigationRailDestination(icon=ft.Icons.EVENT_OUTLINED, selected_icon=ft.Icons.EVENT, label="Key dates"),
        ] + [
            ft.NavigationRailDestination(icon=ft.Text(emoji(t.name), size=22), label=_short(t.name))
            for t in self.tracker.subteams
        ]
        if not keep_view or (self.view not in ("overview", "keydates") and self.view not in names):
            self.view = "overview"
        self.show(self.view)

    def team(self, name):
        return next(t for t in self.tracker.subteams if t.name == name)

    def publish_rollup(self):
        """Keep Master > All Parts current whenever someone who can edit the master saves."""
        if self.tracker.master_writable:
            try:
                core.write_rollup(self.tracker)
            except core.SaveError:
                pass  # master open in Excel; it'll be refreshed on the next save

    # ---- navigation

    def on_nav(self, e):
        i = e.control.selected_index
        self.show(["overview", "keydates"][i] if i < 2 else self.tracker.subteams[i - 2].name)

    def show(self, view):
        self.view = view
        if view == "overview":
            self.rail.selected_index, content = 0, self.overview()
        elif view == "keydates":
            self.rail.selected_index, content = 1, self.key_dates_view()
        else:
            names = [t.name for t in self.tracker.subteams]
            self.rail.selected_index, content = names.index(view) + 2, self.subteam_view(self.team(view))
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

    # ---- overview

    def overview(self):
        today = date.today()
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
                    attention.append((t, p))
                due = core.due_date(p, t.gates)
                if due and today <= due <= today + timedelta(days=LOOK_AHEAD_DAYS):
                    upcoming.append((due, ft.Icons.ARROW_FORWARD, f"{p.name}: {core.current_gate(p, t.gates)}", t.name, t.name))
            passed = sum(core.progress(p, t.gates)[0] for p in t.parts)
            total = sum(core.progress(p, t.gates)[1] for p in t.parts) or 1
            chips = [pill(f"{stats['Late']} late", "Late") if stats["Late"] else None,
                     pill(f"{flagged} flagged", "At risk") if flagged else None,
                     pill(f"{stats['Due soon']} due soon", "Due soon") if stats["Due soon"] else None]
            chips = [c for c in chips if c] or [pill("All on track", "On track") if t.parts else pill("No parts yet", "No date")]
            cards.append(ft.Container(
                ft.Card(ft.Container(ft.Column([
                    ft.Row([ft.Text(f"{emoji(t.name)}  {t.name}", weight=ft.FontWeight.W_600, size=16, expand=True),
                            ft.Icon(ft.Icons.EDIT_OUTLINED if t.writable else ft.Icons.LOCK_OUTLINE, size=16,
                                    color=ft.Colors.ON_SURFACE_VARIANT, tooltip="You can edit" if t.writable else "View only")]),
                    ft.Text(f"{len(t.parts)} parts · {stats['Done']} done", color=ft.Colors.ON_SURFACE_VARIANT, size=13),
                    ft.ProgressBar(value=passed / total, bar_height=6, border_radius=3),
                    ft.Row(chips, wrap=True, spacing=6),
                ], spacing=10), padding=16, ink=True, on_click=lambda e, n=t.name: self.show(n)), variant=ft.CardVariant.FILLED),
                col={"xs": 12, "sm": 6, "lg": 4},
            ))
        for kd in self.tracker.key_dates:
            if kd.date and today <= kd.date <= today + timedelta(days=LOOK_AHEAD_DAYS):
                upcoming.append((kd.date, TYPE_ICONS.get(kd.type, ft.Icons.EVENT), kd.event, kd.type, "keydates"))
        upcoming.sort(key=lambda u: u[0])
        attention.sort(key=lambda tp: (core.due_date(tp[1], tp[0].gates) or date.max))

        next_review = min((k for k in self.tracker.key_dates if k.type == "Design Review" and k.date and k.date >= today),
                          key=lambda k: k.date, default=None)
        sub = f"{today:%A %d %B %Y}"
        if next_review:
            sub += f"  ·  Next design review: {next_review.event}, {fmt(next_review.date)}"

        attention_rows = [
            ft.ListTile(
                leading=ft.Icon(ft.Icons.BLOCK if p.flag == "Blocked" else ft.Icons.WARNING_AMBER, color=ft.Colors.ERROR),
                title=ft.Text(p.name), subtitle=ft.Text(f"{t.name} · {core.current_gate(p, t.gates)} due {fmt(core.due_date(p, t.gates))}"
                                                       + (f" · {p.owner}" if p.owner else "") + (f"\n{p.notes}" if p.notes else "")),
                trailing=ft.Row([c for c in (pill(p.flag) if p.flag != "OK" else None, pill(core.status(p, t.gates))) if c],
                                spacing=6, tight=True),
                on_click=lambda e, n=t.name: self.show(n),
            )
            for t, p in attention
        ]
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
                                     ft.Column(attention_rows) if attention_rows else empty("Nothing late or flagged.")),
                             col={"xs": 12, "lg": 7}),
                ft.Container(section(f"Next {LOOK_AHEAD_DAYS} days",
                                     ft.Column(upcoming_rows) if upcoming_rows else empty("Nothing due.")),
                             col={"xs": 12, "lg": 5}),
            ], spacing=12, run_spacing=12),
        ], spacing=20, scroll=ft.ScrollMode.AUTO)
        return content, None

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
        header = ft.Row([heading(f"{emoji(team.name)}  {team.name}", " → ".join(team.gates)), lock],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
        order = {"Late": 0, "Due soon": 1, "On track": 2, "No date": 3, "Done": 4}
        rows = []
        for idx in sorted(range(len(team.parts)),
                          key=lambda i: (team.parts[i].flag == "OK", order[core.status(team.parts[i], team.gates, today)],
                                         core.due_date(team.parts[i], team.gates) or date.max)):
            p = team.parts[idx]
            st = core.status(p, team.gates, today)
            rows.append(ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(p.name, weight=ft.FontWeight.W_500)),
                    ft.DataCell(ft.Text(p.owner or "—")),
                    ft.DataCell(ft.Column([ft.Text(core.current_gate(p, team.gates), size=13), gate_dots(p, team.gates)],
                                          spacing=4, tight=True, alignment=ft.MainAxisAlignment.CENTER)),
                    ft.DataCell(ft.Text(fmt(core.due_date(p, team.gates)))),
                    ft.DataCell(ft.Row([c for c in (pill(st), pill(p.flag) if p.flag != "OK" else None) if c], spacing=6)),
                    ft.DataCell(ft.Text(p.notes, max_lines=2, overflow=ft.TextOverflow.ELLIPSIS, width=260)),
                ],
                on_select_change=lambda e, i=idx: self.edit_part(team, i),
            ))
        table = ft.DataTable(
            columns=[ft.DataColumn(ft.Text(h)) for h in ("Part", "Owner", "Gate", "Due", "Status", "Notes")],
            rows=rows, show_checkbox_column=False, data_row_min_height=56, data_row_max_height=72, column_spacing=28,
        )
        body = ft.Row([table], scroll=ft.ScrollMode.AUTO) if rows else empty(
            "No parts yet." + (" Press + to add one." if team.writable else ""))
        fab = ft.FloatingActionButton(icon=ft.Icons.ADD, content="Add part", on_click=lambda e: self.edit_part(team, None)) if team.writable else None
        return ft.Column([header, section("Parts", body)], spacing=20, scroll=ft.ScrollMode.AUTO), fab

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
            self.commit(team, lambda: team.parts.__setitem__(index, new) if index is not None else team.parts.append(new))

        def delete(e):
            self.commit(team, lambda: team.parts.pop(index))

        actions = [ft.TextButton("Close" if ro else "Cancel", on_click=lambda e: self.page.pop_dialog())]
        if not ro:
            if index is not None:
                actions.insert(0, ft.TextButton("Delete", icon=ft.Icons.DELETE_OUTLINE, on_click=delete,
                                                style=ft.ButtonStyle(color=ft.Colors.ERROR)))
            actions.append(ft.FilledButton("Save", on_click=save))
        gate_row = [gate] + ([ft.FilledTonalButton("Gate passed", icon=ft.Icons.CHECK_CIRCLE, on_click=advance)] if not ro else [])
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(("Edit part" if index is not None else "New part") if not ro else part.name),
            content=ft.Container(ft.Column([
                name, owner,
                ft.Row(gate_row, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                ft.Text("Flag", size=12, color=ft.Colors.ON_SURFACE_VARIANT), flag,
                notes,
                ft.Text("Gate due dates", theme_style=ft.TextThemeStyle.TITLE_SMALL),
                *[f.control for f in date_fields.values()],
            ], tight=True, spacing=12, scroll=ft.ScrollMode.AUTO, horizontal_alignment=ft.CrossAxisAlignment.STRETCH), width=460),
            actions=actions,
        ))

    def commit(self, team, change):
        """Apply a change to a subteam and save it; on any failure reload the file so the screen matches it."""
        backup = list(team.parts)
        change()
        try:
            core.save_subteam(team)
        except core.SaveError as e:
            team.parts = backup
            self.page.pop_dialog()
            self.toast(str(e), error=True)
            if isinstance(e, core.ChangedOnDiskError):
                self.reload()
            return
        self.page.pop_dialog()
        self.publish_rollup()
        self.toast(f"Saved to {team.path.name}")
        self.show(team.name)

    # ---- key dates

    def key_dates_view(self):
        today = date.today()
        can_edit = self.tracker.master_writable
        show_past = load_config().get("show_past", False)
        items = sorted(enumerate(self.tracker.key_dates), key=lambda ik: ik[1].date or date.max)
        tiles = []
        for i, kd in items:
            past = kd.date is not None and kd.date < today
            if past and not show_past:
                continue
            days = (kd.date - today).days if kd.date else None
            when = "today" if days == 0 else ("TBD" if days is None else (f"in {days} days" if days > 0 else f"{-days} days ago"))
            tiles.append(ft.ListTile(
                leading=ft.Icon(TYPE_ICONS.get(kd.type, ft.Icons.EVENT), color=ft.Colors.OUTLINE if past else ft.Colors.PRIMARY),
                title=ft.Text(kd.event, color=ft.Colors.ON_SURFACE_VARIANT if past else None),
                subtitle=ft.Text(f"{fmt(kd.date)} · {kd.type}"),
                trailing=pill(when, "Due soon" if days is not None and 0 <= days <= core.DUE_SOON_DAYS else "No date"),
                on_click=(lambda e, i=i: self.edit_key_date(i)) if can_edit else None,
            ))

        def toggle_past(e):
            cfg = load_config()
            cfg["show_past"] = e.control.value
            save_config(cfg)
            self.show("keydates")

        sub = "Competition deadlines, design reviews and milestones." + ("" if can_edit else "  View only: the PM/captain edits these.")
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

        def save(e):
            try:
                d = when.parse()
            except ValueError as err:
                self.toast(str(err), error=True)
                return
            if not event.value.strip() or d is None:
                self.toast("An event needs a name and a date.", error=True)
                return
            new = core.KeyDate(d, event.value.strip(), kind.selected[0])
            self.commit_key_dates(lambda: self.tracker.key_dates.__setitem__(index, new) if index is not None
                                  else self.tracker.key_dates.append(new))

        actions = [ft.TextButton("Cancel", on_click=lambda e: self.page.pop_dialog()), ft.FilledButton("Save", on_click=save)]
        if index is not None:
            actions.insert(0, ft.TextButton("Delete", icon=ft.Icons.DELETE_OUTLINE, style=ft.ButtonStyle(color=ft.Colors.ERROR),
                                            on_click=lambda e: self.commit_key_dates(lambda: self.tracker.key_dates.pop(index))))
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text("Edit date" if index is not None else "New date"),
            content=ft.Container(ft.Column([event, when.control, kind], tight=True, spacing=12), width=460),
            actions=actions,
        ))

    def commit_key_dates(self, change):
        backup = list(self.tracker.key_dates)
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
        self.show("keydates")


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
    ap.add_argument("--start", help="open on this view: overview, keydates, or a subteam name")
    args = ap.parse_args()
    ft.run(lambda page: main(page, args.folder, args.start), assets_dir=str(ASSETS))
