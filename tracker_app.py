"""Project Tracker: a Material Design front end over the SharePoint-synced Excel files.

    pip install -r requirements.txt
    python tracker_app.py                      # finds the synced folder, or asks for it
    python tracker_app.py --folder "C:/Users/me/OneDrive - Uni/Project Tracker"

The app never decides who can edit what. SharePoint does: files you can't edit
are synced read-only, and the app shows them as view-only.
"""

import argparse
import contextlib
import copy
import ctypes
import sys
import threading
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import flet as ft

import tracker_core as core
from tracker_core import load_config, save_config

ASSETS = Path(__file__).resolve().parent / "assets"
LOOK_AHEAD_DAYS = 14
FIXED_VIEWS = ["overview", "mywork", "calendar", "keydates", "gates"]
RAIL_VIEWS = FIXED_VIEWS[:4]  # the rail's own destinations; Gates opens from the settings menu
SUBVIEWS = ["List", "Grid", "Design reviews"]
SUBTEAM_EMOJI = {  # matched against the start of the subteam's file name
    "aero": "🪽", "suspension": "🛞", "chassis": "🏎️", "powertrain": "🔋",
    "electrical": "⚡", "driverless": "🤖",
}

STATUS_COLOURS = {  # status -> (container colour, text colour)
    "Late": (ft.Colors.ERROR_CONTAINER, ft.Colors.ON_ERROR_CONTAINER),
    "Due soon": None,  # amber / green ones depend on light or dark: see STATUS_TONES
    "On track": None,
    "Done": (ft.Colors.SURFACE_CONTAINER_HIGHEST, ft.Colors.ON_SURFACE_VARIANT),
    "Held": (ft.Colors.SURFACE_CONTAINER_HIGHEST, ft.Colors.ON_SURFACE_VARIANT),
    "No date": (ft.Colors.SURFACE_CONTAINER_HIGHEST, ft.Colors.ON_SURFACE_VARIANT),
    "At risk": None,
    "Blocked": (ft.Colors.ERROR_CONTAINER, ft.Colors.ON_ERROR_CONTAINER),
    "You": (ft.Colors.PRIMARY_CONTAINER, ft.Colors.ON_PRIMARY_CONTAINER),
    "Now": None,  # a blackout or event that's on right now
}
STATUS_TONES = {
    "dark": {"amber": ("#3D3112", "#FFD77A"), "green": ("#16331F", "#8FDCA5")},
    "light": {"amber": (ft.Colors.AMBER_100, ft.Colors.AMBER_900), "green": (ft.Colors.GREEN_100, ft.Colors.GREEN_900)},
}


def use_mode(dark):
    tones = STATUS_TONES["dark" if dark else "light"]
    STATUS_COLOURS.update({"Due soon": tones["amber"], "At risk": tones["amber"], "Now": tones["amber"], "On track": tones["green"]})


CALENDAR_COLOURS = {  # calendar-only keys: gates still to come, and upcoming design reviews
    "Future": (ft.Colors.SECONDARY_CONTAINER, ft.Colors.ON_SECONDARY_CONTAINER),
    "Review": (ft.Colors.TERTIARY_CONTAINER, ft.Colors.ON_TERTIARY_CONTAINER),
    "Blackout": (ft.Colors.SURFACE_CONTAINER_HIGHEST, ft.Colors.ON_SURFACE_VARIANT),
}
# Team look, from the Pit Wall newsletter and the Gantt export: near-black, red, warm grey.
# Pure brand red (#FF0000) stays in the logo; text and icons use Google's softer reds (dark #F28B82, light #C5221F),
# per Material's advice to avoid saturated colours for text on dark surfaces. Both pass WCAG AA on their backgrounds.
FONT = None  # the platform default (Segoe UI on Windows); "Arial" matches the documents
PALETTE = {
    "dark": dict(
        primary="#F28B82", on_primary="#601410", primary_container="#5A0B0B", on_primary_container="#FFDAD5",
        secondary="#D3C8C7", on_secondary="#171717", secondary_container="#2E2B2B", on_secondary_container="#F2ECEB",
        tertiary="#D3C8C7", on_tertiary="#171717", tertiary_container="#4A2222", on_tertiary_container="#FFDAD5",
        error="#F28B82", on_error="#601410", error_container="#5A0B0B", on_error_container="#FFDAD5",
        surface="#171717", on_surface="#FFFFFF", on_surface_variant="#D3C8C7", outline="#5C5656", outline_variant="#2C2C2C",
        surface_container_lowest="#121212", surface_container_low="#1C1C1C", surface_container="#202020",
        surface_container_high="#262626", surface_container_highest="#2E2E2E", surface_tint="#171717",
        inverse_surface="#F2ECEB", on_inverse_surface="#171717",
    ),
    "light": dict(
        primary="#C5221F", on_primary="#FFFFFF", primary_container="#FCE8E6", on_primary_container="#410000",
        secondary="#5C5656", on_secondary="#FFFFFF", secondary_container="#ECE6E5", on_secondary_container="#171717",
        tertiary="#5C5656", on_tertiary="#FFFFFF", tertiary_container="#F6DCD8", on_tertiary_container="#410000",
        error="#C5221F", on_error="#FFFFFF", error_container="#FCE8E6", on_error_container="#410002",
        surface="#FFFFFF", on_surface="#171717", on_surface_variant="#5C5656", outline="#8C8483", outline_variant="#E3DCDB",
        surface_container_lowest="#FFFFFF", surface_container_low="#FAF8F8", surface_container="#F5F2F2",
        surface_container_high="#EFEBEB", surface_container_highest="#E8E3E2", surface_tint="#FFFFFF",
        inverse_surface="#171717", on_inverse_surface="#FFFFFF",
    ),
}


def team_theme(mode):
    """Material 3, coloured with the team palette instead of a generated one."""
    p = PALETTE[mode]
    return ft.Theme(
        use_material3=True, font_family=FONT, color_scheme=ft.ColorScheme(**p), scaffold_bgcolor=p["surface"],
        navigation_rail_theme=ft.NavigationRailTheme(bgcolor=p["surface"], indicator_color=p["primary_container"]),
        card_theme=ft.CardTheme(color=p["surface_container_low"], shadow_color=ft.Colors.TRANSPARENT),
        divider_color=p["outline_variant"],
    )


PART_COLUMNS = ("Part", "Owner", "Gate", "Due", "Status", "Notes")
COLUMN_GAP = 20
TABLE_CHROME = 200  # rail, page padding and card around the parts table
MIN_NOTES = 160  # narrower than this and the table keeps fixed widths and scrolls sideways
TYPE_ICONS = {"Deadline": ft.Icons.FLAG_OUTLINED, "Design Review": ft.Icons.GROUPS_OUTLINED, "Milestone": ft.Icons.EMOJI_EVENTS_OUTLINED,
              "Freeze": ft.Icons.AC_UNIT, "Blackout": ft.Icons.DO_NOT_DISTURB_ON_OUTLINED, "Event": ft.Icons.CELEBRATION_OUTLINED}


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
    controls = [ft.Text(text, theme_style=ft.TextThemeStyle.HEADLINE_MEDIUM, weight=ft.FontWeight.W_800)]
    if sub:
        controls.append(ft.Text(sub, color=ft.Colors.ON_SURFACE_VARIANT))
    return ft.Column(controls, spacing=2)


def section(title, body, trailing=None):
    head = ft.Row([ft.Text(title, theme_style=ft.TextThemeStyle.TITLE_MEDIUM, weight=ft.FontWeight.W_700, color=ft.Colors.PRIMARY)]
                  + ([trailing] if trailing else []),
                  alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
    return ft.Card(ft.Container(ft.Column([head, body], spacing=8), padding=16), variant=ft.CardVariant.OUTLINED)


def empty(text):
    return ft.Container(ft.Text(text, color=ft.Colors.ON_SURFACE_VARIANT, italic=True), padding=ft.Padding.symmetric(vertical=8))


def gate_dots(part, team):
    """One dot per gate of the part's template: filled = done, amber = in progress, ringed = ready to start,
    hollow grey = off, faint = waiting for an earlier gate."""
    dots = []
    for g in core.template_of(team, part).gates:
        state = core.gate_state(part, team, g)
        due = part.dates.get(g)
        border = None
        if state == "done":
            colour, tip = ft.Colors.ON_SURFACE, "done"
        elif state == "in progress":
            colour, border, tip = STATUS_COLOURS["Due soon"][1], ft.Border.all(2, ft.Colors.PRIMARY), f"in progress · due {fmt(due)}"
        elif state == "ready":
            colour = STATUS_COLOURS.get(core.date_status(due), STATUS_COLOURS["No date"])[0]
            border, tip = ft.Border.all(2, ft.Colors.PRIMARY), f"ready · due {fmt(due)}"
        elif state == "off":
            colour, border, tip = None, ft.Border.all(1.5, ft.Colors.OUTLINE), "off"
        else:
            colour, tip = ft.Colors.SURFACE_CONTAINER_HIGHEST, f"waiting · due {fmt(due)}"
        dots.append(ft.Container(width=14, height=14, border_radius=7, bgcolor=colour, border=border, tooltip=f"{g}: {tip}"))
    return ft.Row(dots, spacing=4)


GRID_STATE = {  # state -> (symbol, background key) for the grid view
    core.DONE_STATE: ("✓", "done"), core.IN_PROGRESS: ("◐", "progress"), core.OFF: ("–", "off"), core.NOT_STARTED: ("", "empty"),
}


def tag_chip(tag):
    return ft.Container(ft.Text(tag, size=11, color=ft.Colors.ON_SECONDARY_CONTAINER), bgcolor=ft.Colors.SECONDARY_CONTAINER,
                        border_radius=8, padding=ft.Padding.symmetric(horizontal=7, vertical=1))


def plural(n, word):
    return f"{n} {word}{'' if n == 1 else 's'}"


def lateness(passed, due):
    if not (passed and due):
        return ""
    days = (passed - due).days
    return "on time" if days == 0 else f"{plural(abs(days), 'day')} {'late' if days > 0 else 'early'}"


class TrackerApp:
    def __init__(self, page, folder, on_change_folder=None, on_open_folder=None):
        self.page = page
        self.folder = Path(folder)
        self.tracker = None
        self.view = "overview"
        self.me = load_config().get("me", "")
        self.cal_month = date.today().replace(day=1)
        self.expanded = set()  # parts whose sub-parts are shown
        self.part_table, self.part_query, self.part_widths, self.notes_width = None, "", [], None
        self.part_tag = None  # tag filter on the subteam page
        self.subview = load_config().get("subview", "List")
        self.on_open_folder = on_open_folder
        self.car_button = ft.TextButton("", on_click=lambda e: self.cars_dialog(), tooltip="Switch car, or start the next one",
                                        style=ft.ButtonStyle(padding=ft.Padding.symmetric(horizontal=4)))
        page.on_resize = self.on_resize
        self.rail = ft.NavigationRail(
            label_type=ft.NavigationRailLabelType.ALL, min_width=96, group_alignment=-0.9, scrollable=True, on_change=self.on_nav,
            leading=ft.Container(ft.Column([ft.Image(src="icon.png", width=36, height=36, border_radius=5), self.car_button],
                                           spacing=0, horizontal_alignment=ft.CrossAxisAlignment.CENTER), padding=ft.Padding.only(top=8)),
            trailing=ft.Container(ft.Column([
                ft.IconButton(ft.Icons.REFRESH, tooltip="Reload files", on_click=lambda e: self.reload()),
                self.settings_menu(on_change_folder),
            ], spacing=4, horizontal_alignment=ft.CrossAxisAlignment.CENTER), padding=ft.Padding.only(bottom=12)),
            pin_trailing_to_bottom=True,
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
            ft.NavigationRailDestination(icon=ft.Icons.CALENDAR_MONTH_OUTLINED, selected_icon=ft.Icons.CALENDAR_MONTH, label="Calendar"),
            ft.NavigationRailDestination(icon=ft.Icons.EVENT_OUTLINED, selected_icon=ft.Icons.EVENT, label="Key dates"),
        ] + [
            ft.NavigationRailDestination(icon=ft.Text(emoji(t.name), size=22), label=_short(t.name))
            for t in self.tracker.subteams
        ]
        self.car_button.content = ft.Text(self.tracker.car, size=12, weight=ft.FontWeight.W_700)
        if self.view not in FIXED_VIEWS and self.view not in [t.name for t in self.tracker.subteams]:
            self.view = "overview"
        self.show(self.view)

    def team(self, name):
        return next(t for t in self.tracker.subteams if t.name == name)

    def known_names(self):
        names = {}
        for t in self.tracker.subteams:
            for n in [n for p in t.parts for n in p.people()] + [a.owner for a in t.actions]:
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
                pass  # it'll be refreshed on the next save

    # ---- navigation

    def on_nav(self, e):
        i = e.control.selected_index
        self.show(RAIL_VIEWS[i] if i < len(RAIL_VIEWS) else self.tracker.subteams[i - len(RAIL_VIEWS)].name)

    def show(self, view):
        self.view = view
        if view in FIXED_VIEWS:
            index = RAIL_VIEWS.index(view) if view in RAIL_VIEWS else None  # Gates lives in the settings menu, not the rail
            content = {"overview": self.overview, "mywork": self.my_work, "calendar": self.calendar_view,
                       "keydates": self.key_dates_view, "gates": self.gates_view}[view]()
        else:
            index = [t.name for t in self.tracker.subteams].index(view) + len(RAIL_VIEWS)
            content = self.subteam_view(self.team(view))
        self.rail.selected_index = index
        self.page.floating_action_button = content[1]
        self.body.content = content[0]
        self.page.update()

    def on_resize(self, e):
        # the parts table switches between "Notes fills the row" and "fixed widths + scroll" at narrow sizes
        if self.view not in FIXED_VIEWS and self.part_widths and (self.table_notes_width() is None) != (self.notes_width is None):
            self.show(self.view)

    def settings_menu(self, on_change_folder):
        """Material settings menu at the foot of the rail: night mode, the tracker folder, and the Gates page."""
        # The switch only shows the state: a tap anywhere on the row closes the menu, then switches the theme
        # (an open menu keeps the old theme's colours, so toggling it in place leaves the menu unreadable).
        self.night_switch = ft.Switch(value=self.page.theme_mode == ft.ThemeMode.DARK, scale=0.8)
        night = ft.Container(self.night_switch, ignore_interactions=True)

        def item(icon, label, on_click, trailing=None, sub=None):
            text = ft.Column([ft.Text(label), ft.Text(sub, size=12, color=ft.Colors.ON_SURFACE_VARIANT, max_lines=1,
                                                      overflow=ft.TextOverflow.ELLIPSIS, width=220)] if sub else [ft.Text(label)],
                             spacing=0, tight=True)
            return ft.PopupMenuItem(content=ft.Row([ft.Icon(icon, color=ft.Colors.ON_SURFACE_VARIANT, size=20),
                                                    ft.Container(text, expand=True)] + ([trailing] if trailing else []),
                                                   spacing=12, width=280), on_click=on_click)

        items = [item(ft.Icons.DARK_MODE_OUTLINED, "Night mode", self.toggle_theme, trailing=night)]
        if on_change_folder:
            items.append(item(ft.Icons.FOLDER_OPEN_OUTLINED, "Tracker folder", lambda e: on_change_folder(), sub=str(self.folder)))
        items += [ft.PopupMenuItem(),  # divider
                  item(ft.Icons.ACCOUNT_TREE_OUTLINED, "Gates", lambda e: self.show("gates"), sub="Gate templates and library")]
        return ft.PopupMenuButton(icon=ft.Icons.SETTINGS_OUTLINED, tooltip="Settings", items=items,
                                  menu_position=ft.PopupMenuPosition.OVER, bgcolor=ft.Colors.SURFACE_CONTAINER)

    def toggle_theme(self, e):
        dark = self.page.theme_mode != ft.ThemeMode.DARK
        self.page.theme_mode = ft.ThemeMode.DARK if dark else ft.ThemeMode.LIGHT
        self.night_switch.value = dark
        use_mode(dark)
        cfg = load_config()
        cfg["dark"] = dark
        save_config(cfg)
        self.show(self.view)

    def toast(self, text, error=False):
        self.page.show_dialog(ft.SnackBar(
            ft.Text(text, color=ft.Colors.ON_ERROR_CONTAINER if error else None),
            bgcolor=ft.Colors.ERROR_CONTAINER if error else None, show_close_icon=True,
        ))

    # ---- shared list tiles

    def part_tile(self, team, part, today, show_team=True):
        index = team.parts.index(part)
        sub = f"{core.current_gate(part, team)} due {fmt(core.due_date(part, team))}"
        sub = (f"{team.name} · " if show_team else "") + sub + (f" · {part.owner}" if part.owner else "") + (f"\n{part.notes}" if part.notes else "")
        return ft.ListTile(
            leading=ft.Icon(ft.Icons.BLOCK if part.flag == "Blocked" else ft.Icons.PRECISION_MANUFACTURING_OUTLINED,
                            color=ft.Colors.ERROR if core.needs_attention(part, team, today) else ft.Colors.ON_SURFACE_VARIANT),
            title=ft.Text(part.key), subtitle=ft.Text(sub),
            trailing=pills((part.flag, None) if part.flag != "OK" else None, (core.status(part, team, today), None)),
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
        bits = [fmt(kd.date) + (f" – {fmt(kd.end)}" if kd.end and kd.end != kd.date else "")]
        if kd.internal:
            bits.append(f"internal {fmt(kd.internal)}")
        bits += [x for x in (kd.type, kd.track, kd.competition) if x] + ([", ".join(kd.people)] if kd.people else [])
        mine = any(core.is_me(n, self.me) for n in kd.people)
        target = kd.date if kd.type in core.PERIOD_TYPES else core.key_date_target(kd)
        label = "Done" if kd.done else (st if st in ("Held", "Now") else countdown(target, today))
        index = self.tracker.key_dates.index(kd)
        return ft.ListTile(
            leading=ft.Icon(TYPE_ICONS.get(kd.type, ft.Icons.EVENT),
                            color=ft.Colors.ERROR if st == "Late" else (ft.Colors.OUTLINE if st in ("Done", "Held") else ft.Colors.ON_SURFACE_VARIANT)),
            title=ft.Text(kd.event, color=ft.Colors.ON_SURFACE_VARIANT if st in ("Done", "Held") else None),
            subtitle=ft.Text(" · ".join(bits) + (f"\n{kd.notes}" if kd.notes else ""), max_lines=3, overflow=ft.TextOverflow.ELLIPSIS),
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
                st = core.status(p, t, today)
                if st in stats:
                    stats[st] += 1
                flagged += p.flag != "OK"
                if core.needs_attention(p, t, today):
                    attention.append((core.due_date(p, t) or date.max, self.part_tile(t, p, today)))
                due = core.due_date(p, t)
                if due and today <= due <= soon:
                    upcoming.append((due, ft.Icons.PRECISION_MANUFACTURING_OUTLINED, f"{p.key}: {core.current_gate(p, t)}", t.name, t.name))
            for r in t.reviews:
                if r.date and today <= r.date <= soon and not r.outcome:
                    upcoming.append((r.date, ft.Icons.GROUPS_OUTLINED, f"{r.part}: {r.review}", f"Design review · {t.name}", t.name))
            late_actions = 0
            for a in t.actions:
                st = core.action_status(a, today)
                if st == "Late":
                    late_actions += 1
                    attention.append((a.due, self.action_tile(t, a, today)))
                elif a.closed is None and a.due and today <= a.due <= soon:
                    upcoming.append((a.due, ft.Icons.TASK_ALT, a.text, f"Action · {t.name}" + (f" · {a.owner}" if a.owner else ""), t.name))
            passed = sum(core.progress(p, t)[0] for p in t.parts)
            total = sum(core.progress(p, t)[1] for p in t.parts) or 1
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
                    ft.ProgressBar(value=passed / total, bar_height=6, border_radius=3, color=ft.Colors.ON_SURFACE,
                                   bgcolor=ft.Colors.SURFACE_CONTAINER_HIGHEST),
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
            ft.ListTile(leading=ft.Icon(icon, color=ft.Colors.ON_SURFACE_VARIANT), title=ft.Text(title),
                        subtitle=ft.Text(f"{fmt(d)} · {sub_}"), dense=True, on_click=lambda e, v=view: self.show(v))
            for d, icon, title, sub_, view in upcoming
        ]
        new_team = ft.OutlinedButton("New sub-team", icon=ft.Icons.GROUP_ADD_OUTLINED, on_click=lambda e: self.new_subteam_dialog())
        content = ft.Column([
            ft.Row([heading(f"{self.tracker.car} overview", sub), new_team], alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                   vertical_alignment=ft.CrossAxisAlignment.START),
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
        header = ft.Row([heading("My work", "Parts you own or work on, your actions, and the key dates you're responsible for."), who],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN, vertical_alignment=ft.CrossAxisAlignment.START)
        if not self.me:
            return ft.Column([header, section("Get started", empty(
                "Pick your name above. Names come from part owners and assignees, action owners and key-date people."))], spacing=20), None

        parts = [(t, p) for t in self.tracker.subteams for p in t.parts
                 if any(core.is_me(n, self.me) for n in p.people()) and not core.is_done(p, t)]
        parts.sort(key=lambda tp: (not core.needs_attention(tp[1], tp[0], today), core.due_date(tp[1], tp[0]) or date.max))
        actions = [(t, a) for t in self.tracker.subteams for a in t.actions if core.is_me(a.owner, self.me) and a.closed is None]
        actions.sort(key=lambda ta: ta[1].due or date.max)
        dates = [k for k in self.tracker.key_dates
                 if any(core.is_me(n, self.me) for n in k.people) and core.key_date_status(k, today) not in ("Done", "Held")]
        dates.sort(key=lambda k: k.date or date.max)
        mine = {(t.name, p.key) for t in self.tracker.subteams for p in t.parts if any(core.is_me(n, self.me) for n in p.people())}
        reviews = sorted(((t, r) for t in self.tracker.subteams for r in t.reviews
                          if (t.name, r.part) in mine and core.review_status(r, today) != "Held"), key=lambda tr: tr[1].date or date.max)
        late = sum(core.needs_attention(p, t, today) for t, p in parts) + \
            sum(core.action_status(a, today) == "Late" for _, a in actions) + \
            sum(core.key_date_needs_attention(k, today) for k in dates)
        summary = pill(f"{late} need attention", "Late") if late else pill("Nothing late", "On track")
        return ft.Column([
            header,
            ft.Row([summary]),
            section(f"My parts ({len(parts)})", ft.Column([self.part_tile(t, p, today) for t, p in parts]) if parts else empty("No open parts.")),
            section(f"My actions ({len(actions)})", ft.Column([self.action_tile(t, a, today) for t, a in actions]) if actions else empty("No open actions.")),
            section(f"My key dates ({len(dates)})", ft.Column([self.key_date_tile(k, today) for k in dates]) if dates else empty("No upcoming key dates.")),
            section(f"My design reviews ({len(reviews)})", ft.Column([self.review_tile(t, r, today) for t, r in reviews]) if reviews
                    else empty("No design reviews scheduled for your parts.")),
        ], spacing=20, scroll=ft.ScrollMode.AUTO), None

    def review_tile(self, team, review, today):
        st = core.review_status(review, today)
        label = review.outcome or ("Needs outcome" if st == "Needs outcome" else countdown(review.date, today))
        return ft.ListTile(
            leading=ft.Icon(ft.Icons.GROUPS_OUTLINED, color=ft.Colors.ERROR if st == "Needs outcome" else ft.Colors.ON_SURFACE_VARIANT),
            title=ft.Text(f"{review.part}: {review.review}"),
            subtitle=ft.Text(" · ".join(x for x in (team.name, fmt(review.date), review.presenter and f"presented by {review.presenter}") if x)),
            trailing=pills((label, "Late" if st == "Needs outcome" else st)),
            on_click=lambda e: self.edit_review(team, review.part, review.review),
        )

    # ---- calendar

    def calendar_events(self, hidden):
        """date -> [(sort key, label, colour key, on_click)] for everything not toggled off."""
        today = date.today()
        events = {}

        def add(d, order, label, key, click):
            if d:
                events.setdefault(d, []).append((order, label, key, click))

        for t in self.tracker.subteams:
            if t.name in hidden:
                continue
            for r in t.reviews:
                if r.date:
                    add(r.date, 0, f"◆ {r.part}: {r.review}", "Held" if r.outcome else "Review",
                        lambda e, t=t, r=r: self.edit_review(t, r.part, r.review))
            for pi, p in enumerate(t.parts):
                for g in core.template_of(t, p).gates:
                    state = core.gate_state(p, t, g)
                    if state == "off":
                        continue
                    key = {"done": "Done", "waiting": "Future"}.get(state) or core.date_status(p.dates.get(g), False, today)
                    add(p.dates.get(g), 1, f"{emoji(t.name)} {p.key}: {g}", key, lambda e, t=t, i=pi: self.edit_part(t, i))
            if "Actions" not in hidden:
                for ai, a in enumerate(t.actions):
                    add(a.due, 2, f"{'☑' if a.closed else '☐'} {a.text}", core.action_status(a, today),
                        lambda e, t=t, i=ai: self.edit_action(t, i))
        if "Key dates" not in hidden:
            for ki, k in enumerate(self.tracker.key_dates):
                st = core.key_date_status(k, today)
                click = (lambda e, i=ki: self.edit_key_date(i)) if self.tracker.master_writable else None
                key = "Review" if k.type == "Design Review" and st != "Held" else ("Blackout" if k.type == "Blackout" else st)
                for day in core.key_date_days(k):
                    add(day, 0, f"★ {k.event}", key, click)
                if k.internal and k.internal != k.date and not k.done:
                    add(k.internal, 0, f"⏱ {k.event} (internal)", st, click)
        for day in events.values():
            day.sort(key=lambda ev: (ev[0], ev[1]))
        return events

    def calendar_view(self):
        today = date.today()
        hidden = set(load_config().get("calendar_hidden", []))
        events = self.calendar_events(hidden)
        month = self.cal_month

        def toggle(name):
            def handler(e):
                cfg = load_config()
                h = set(cfg.get("calendar_hidden", []))
                h.symmetric_difference_update({name})
                cfg["calendar_hidden"] = sorted(h)
                save_config(cfg)
                self.show("calendar")
            return handler

        def go(months):
            def handler(e):
                if months == 0:
                    self.cal_month = today.replace(day=1)
                else:
                    m = self.cal_month.month - 1 + months
                    self.cal_month = date(self.cal_month.year + m // 12, m % 12 + 1, 1)
                self.show("calendar")
            return handler

        chips = [ft.Chip(label=ft.Text(f"{emoji(t.name)} {_short(t.name)}"), selected=t.name not in hidden,
                         on_select=toggle(t.name), show_checkmark=False) for t in self.tracker.subteams]
        chips += [ft.Chip(label=ft.Text(f"{icon} {name}"), selected=name not in hidden, on_select=toggle(name), show_checkmark=False)
                  for icon, name in (("★", "Key dates"), ("☐", "Actions"))]

        start = month - timedelta(days=month.weekday())  # Monday on or before the 1st
        weeks = []
        for w in range(6):
            days = [start + timedelta(days=w * 7 + i) for i in range(7)]
            if w and days[0].month != month.month:
                break
            weeks.append(ft.Row([self.calendar_cell(d, month, today, events.get(d, [])) for d in days], spacing=4))
        header = ft.Row([
            heading("Calendar", "Gate due dates, actions and key dates. Click anything to open it."),
            ft.Row([ft.IconButton(ft.Icons.CHEVRON_LEFT, tooltip="Previous month", on_click=go(-1)),
                    ft.Text(f"{month:%B %Y}", theme_style=ft.TextThemeStyle.TITLE_LARGE, width=190, text_align=ft.TextAlign.CENTER),
                    ft.IconButton(ft.Icons.CHEVRON_RIGHT, tooltip="Next month", on_click=go(1)),
                    ft.OutlinedButton("Today", on_click=go(0))], tight=True),
        ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN, vertical_alignment=ft.CrossAxisAlignment.START)
        weekdays = ft.Row([ft.Container(ft.Text(d, weight=ft.FontWeight.W_600, color=ft.Colors.ON_SURFACE_VARIANT), expand=1,
                                        padding=ft.Padding.only(left=8)) for d in ("Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun")], spacing=4)
        return ft.Column([header, ft.Row(chips, wrap=True, spacing=8, run_spacing=8), weekdays, *weeks],
                         spacing=10, scroll=ft.ScrollMode.AUTO), None

    def calendar_cell(self, d, month, today, day_events, limit=3):
        is_today = d == today
        number = ft.Container(
            ft.Text(str(d.day), size=12, weight=ft.FontWeight.W_600 if is_today else None,
                    color=ft.Colors.ON_PRIMARY if is_today else (None if d.month == month.month else ft.Colors.OUTLINE)),
            bgcolor=ft.Colors.PRIMARY if is_today else None, border_radius=11, width=22, height=22, alignment=ft.Alignment.CENTER,
        )
        items = [self.calendar_event(ev) for ev in day_events[:limit]]
        if len(day_events) > limit:
            items.append(ft.Text(f"+{len(day_events) - limit} more", size=11, color=ft.Colors.PRIMARY, weight=ft.FontWeight.W_600))
        return ft.Container(
            ft.Column([number, *items], spacing=3, tight=True),
            expand=1, height=128, padding=6, border_radius=8,
            border=ft.Border.all(1, ft.Colors.OUTLINE_VARIANT),
            bgcolor=None if d.month == month.month else ft.Colors.SURFACE_CONTAINER_LOW,
            on_click=(lambda e: self.day_dialog(d, day_events)) if day_events else None, ink=bool(day_events),
        )

    def calendar_event(self, ev, dense=True):
        _, label, key, click = ev
        bg, fg = CALENDAR_COLOURS.get(key, STATUS_COLOURS.get(key, STATUS_COLOURS["No date"]))
        return ft.Container(
            ft.Text(label, size=11 if dense else 14, color=fg, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS, tooltip=label),
            bgcolor=bg, border_radius=4, padding=ft.Padding.symmetric(horizontal=6, vertical=2 if dense else 8),
            on_click=click, ink=click is not None,
        )

    def day_dialog(self, d, day_events):
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(f"{d:%A %d %B}"),
            content=ft.Container(ft.Column([self.calendar_event(ev, dense=False) for ev in day_events], tight=True, spacing=6,
                                           scroll=ft.ScrollMode.AUTO), width=460),
            actions=[ft.TextButton("Close", on_click=lambda e: self.page.pop_dialog())],
        ))

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
        hit, total = core.on_time(team)
        sub = f"Gates passed on time: {hit} of {total}" if total else None
        if self.subview not in SUBVIEWS:
            self.subview = "List"

        def pick_view(e):
            self.subview = e.control.selected[0]
            cfg = load_config()
            cfg["subview"] = self.subview
            save_config(cfg)
            self.show(self.view)

        views = ft.SegmentedButton(selected=[self.subview], show_selected_icon=False, on_change=pick_view,
                                   segments=[ft.Segment(value=v, label=ft.Text(v)) for v in SUBVIEWS])
        templates = ft.OutlinedButton("Gate templates", icon=ft.Icons.ACCOUNT_TREE_OUTLINED, on_click=lambda e: self.show("gates"))
        header = ft.Row([heading(f"{emoji(team.name)}  {team.name}", sub), ft.Row([views, templates, lock], spacing=12, wrap=True)],
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN, vertical_alignment=ft.CrossAxisAlignment.START)

        self.part_query = ""
        self.part_widths = self.fit_part_columns(team)
        self.notes_width = self.table_notes_width()
        table = self.part_table = ft.Column([self.part_header()] + self.part_rows(team, "", today), spacing=0)

        def filter_parts(e):
            self.part_query = e.control.value
            table.controls = [self.part_header()] + self.part_rows(team, self.part_query, today)
            table.update()

        search = ft.TextField(hint_text="Filter by part, person or note", prefix_icon=ft.Icons.SEARCH, dense=True, width=300,
                              on_change=filter_parts)
        tags = core.all_tags(team)
        if self.part_tag not in tags:
            self.part_tag = None
        tag_row = ft.Row(wrap=True, spacing=6, run_spacing=6)

        def pick_tag(tag):
            def handler(e):
                self.part_tag = None if tag == self.part_tag else tag
                render_tags()
                table.controls = [self.part_header()] + self.part_rows(team, self.part_query, today)
                tag_row.update()
                table.update()
            return handler

        def render_tags():
            tag_row.controls = [ft.Text("Tags:", size=13, color=ft.Colors.ON_SURFACE_VARIANT)] + [
                ft.Chip(label=ft.Text(t), selected=t == self.part_tag, show_checkmark=False, on_select=pick_tag(t)) for t in tags]
        render_tags()
        if not team.parts:
            parts_body = empty("No parts yet." + (" Press + to add one." if team.writable else ""))
        else:
            parts_body = table if self.notes_width is None else ft.Row([table], scroll=ft.ScrollMode.AUTO)
            if tags:
                parts_body = ft.Column([tag_row, parts_body], spacing=8)

        open_actions = sorted((a for a in team.actions if a.closed is None), key=lambda a: a.due or date.max)
        closed = sorted((a for a in team.actions if a.closed), key=lambda a: a.closed, reverse=True)[:5]
        action_tiles = [self.action_tile(team, a, today, show_team=False) for a in open_actions + closed]
        add_action = ft.TextButton("Add action", icon=ft.Icons.ADD_TASK, on_click=lambda e: self.edit_action(team, None)) if team.writable else None
        fab = ft.FloatingActionButton(icon=ft.Icons.ADD, content="Add part", on_click=lambda e: self.edit_part(team, None)) if team.writable else None
        if self.subview == "Grid":
            main = self.grid_view(team, today)
        elif self.subview == "Design reviews":
            main = [self.reviews_view(team, today)]
        else:
            main = [section("Parts", parts_body, search if team.parts else None)]
        return ft.Column([
            header,
            *main,
            section(f"Actions ({len(open_actions)} open)", ft.Column(action_tiles) if action_tiles else empty(
                "No actions. Add follow-ups from design reviews here."), add_action),
        ], spacing=20, scroll=ft.ScrollMode.AUTO), fab

    def fit_part_columns(self, team):
        """Widths for every column but Notes, fitted to the longest entry. Notes takes whatever is left."""
        def fit(texts, px, low, high, extra=0):
            return max(low, min(high, int(max([len(t) for t in texts] + [0]) * px + extra)))
        today = date.today()
        summaries = [plural(len(kids), "sub-part") + f" · {sum(core.needs_attention(c, team, today) for c in kids)} need attention"
                     for kids in (core.children(team, p) for p in team.parts) if kids]
        part = max(fit([p.name for p in team.parts], 7.6, 140, 340, 56), fit(summaries, 6.4, 0, 340, 44))
        owner = fit([p.owner for p in team.parts] + ["Owner"], 7.6, 70, 200, 8)
        gate = max(fit(team.gates + [core.DONE], 7.2, 110, 240, 8), 20 * len(team.gates))
        status = 170 if any(p.flag != "OK" for p in team.parts) else 100  # one pill, or a status and a flag
        return [part, owner, gate, 92, status]

    def table_notes_width(self):
        """None when the window is wide enough for Notes to fill the rest of the row; otherwise a fixed width (and the table scrolls)."""
        fixed = sum(self.part_widths) + COLUMN_GAP * len(self.part_widths) + 24
        return None if (self.page.width or 1280) >= TABLE_CHROME + fixed + MIN_NOTES else 240

    def table_row(self, cells, notes, on_click=None, header=False):
        boxes = [ft.Container(c, width=w) for c, w in zip(cells, self.part_widths)]
        boxes.append(ft.Container(notes, expand=self.notes_width is None, width=self.notes_width))
        return ft.Container(
            ft.Row(boxes, spacing=COLUMN_GAP, vertical_alignment=ft.CrossAxisAlignment.CENTER),
            padding=ft.Padding.symmetric(horizontal=12, vertical=14 if header else 8), on_click=on_click, ink=on_click is not None,
            border=ft.Border(bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT)),
        )

    def part_header(self):
        label = lambda h: ft.Text(h, weight=ft.FontWeight.W_600)
        return self.table_row([label(h) for h in PART_COLUMNS[:-1]], label(PART_COLUMNS[-1]), header=True)

    def part_rows(self, team, query, today):
        """Top-level parts, problems first; a part with sub-parts sorts by its worst one and expands to show them.
        While filtering, a parent is shown (expanded) if any of its sub-parts match."""
        order = {"Late": 0, "Due soon": 1, "On track": 2, "No date": 3, "Done": 4}
        rank = lambda p: (p.flag == "OK", order[core.status(p, team, today)], core.due_date(p, team) or date.max)
        q = (query or "").strip().casefold()
        tag = self.part_tag
        filtering = bool(q or tag)
        match = lambda p: ((not q or q in f"{p.name} {p.parent} {' '.join(p.people())} {' '.join(p.tags)} {p.notes}".casefold())
                           and (not tag or tag in p.tags))
        index = {id(p): i for i, p in enumerate(team.parts)}
        nested = {id(c) for p in team.parts for c in core.children(team, p)}
        families = [(p, sorted(core.children(team, p), key=rank)) for p in team.parts if id(p) not in nested]
        rows = []
        for top, kids in sorted(families, key=lambda f: min([rank(f[0])] + [rank(c) for c in f[1]])):
            shown = [c for c in kids if match(c)] if filtering else kids
            if filtering and not match(top) and not shown:
                continue
            expanded = bool(shown) and (filtering or top.key in self.expanded)
            rows.append(self.part_row(team, top, index[id(top)], today, kids, expanded))
            if expanded:
                rows += [self.part_row(team, c, index[id(c)], today) for c in shown]
        return rows

    def part_row(self, team, p, idx, today, kids=(), expanded=False):
        st = core.status(p, team, today)
        warnings = core.rule_warnings(p, team, self.tracker.key_dates)
        name = ft.Row([ft.Text(p.name, weight=None if p.parent else ft.FontWeight.W_600)]
                      + ([ft.Icon(ft.Icons.WARNING_AMBER_ROUNDED, size=16, color=STATUS_COLOURS["Due soon"][1], tooltip="\n".join(warnings))]
                         if warnings else []), spacing=4, tight=True)
        lines = [name]
        if p.parent:  # a sub-part, indented under its parent
            lead = ft.Row([ft.Container(width=28), ft.Icon(ft.Icons.SUBDIRECTORY_ARROW_RIGHT, size=16, color=ft.Colors.OUTLINE)],
                          spacing=4, tight=True)
        else:
            lead = ft.IconButton(ft.Icons.EXPAND_MORE if expanded else ft.Icons.CHEVRON_RIGHT, icon_size=20,
                                 tooltip="Hide sub-parts" if expanded else "Show sub-parts",
                                 on_click=lambda e: self.toggle_part(team, p)) if kids else ft.Container(width=40)
            if kids:
                late = sum(core.needs_attention(c, team, today) for c in kids)
                summary = plural(len(kids), "sub-part") + (f" · {late} need attention" if late else "")
                lines.append(ft.Text(summary, size=12, color=ft.Colors.ERROR if late else ft.Colors.ON_SURFACE_VARIANT))
        if p.tags:
            lines.append(ft.Row([tag_chip(t) for t in p.tags], spacing=4, run_spacing=4, wrap=True))
        owner = ft.Text(p.owner or "—")
        if p.assignees:
            owner = ft.Column([owner, ft.Text("+ " + ", ".join(p.assignees), size=12, color=ft.Colors.ON_SURFACE_VARIANT, max_lines=1,
                                              overflow=ft.TextOverflow.ELLIPSIS, tooltip=", ".join(p.assignees))], spacing=0, tight=True)
        return self.table_row(
            [
                # the name block takes the rest of the Part column, so long names and tag rows wrap instead of spilling over
                ft.Row([lead, ft.Column(lines, spacing=2, tight=True, expand=True)], spacing=4, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                owner,
                ft.Column([ft.Text(core.current_gate(p, team), size=13), gate_dots(p, team)], spacing=4, tight=True),
                ft.Text(fmt(core.due_date(p, team))),
                pills((st, None), (p.flag, None) if p.flag != "OK" else None),
            ],
            ft.Text(p.notes, max_lines=2, overflow=ft.TextOverflow.ELLIPSIS, tooltip=p.notes or None),
            on_click=lambda e, i=idx: self.edit_part(team, i),
        )

    def toggle_part(self, team, part):
        self.expanded ^= {part.key}
        self.part_table.controls = [self.part_header()] + self.part_rows(team, self.part_query, date.today())
        self.part_table.update()

    def edit_part(self, team, index, parent=""):
        part = team.parts[index] if index is not None else core.Part(parent=parent)
        ro = not team.writable
        today = date.today()
        kids = core.children(team, part) if index is not None else []
        name = ft.TextField(label="Part", value=part.name, read_only=ro, autofocus=index is None)
        parents = [p.name for i, p in enumerate(team.parts) if not p.parent and i != index]
        parent_field = ft.Dropdown(
            label="Sub-part of", value=part.parent or "", disabled=ro or bool(kids), expand=True,
            options=[ft.DropdownOption(key="", text="— (top-level part)")] + [ft.DropdownOption(key=n, text=n) for n in parents],
            tooltip=f"{part.name} has sub-parts, so it stays top-level" if kids else None)
        owner = ft.TextField(label="Owner (accountable)", value=part.owner, read_only=ro, expand=True)
        assignees = ft.TextField(label="Also working on it", value=", ".join(part.assignees), hint_text="e.g. Yi, Artur",
                                 read_only=ro, expand=True)
        tags = ft.TextField(label="Tags", value=", ".join(part.tags), hint_text="e.g. Development, Composites", read_only=ro)
        suggestions = ft.Row(wrap=True, spacing=6, run_spacing=6)
        flag = ft.SegmentedButton(selected=[part.flag], disabled=ro, show_selected_icon=False,
                                  segments=[ft.Segment(value=f, label=ft.Text(f)) for f in core.FLAGS])
        notes = ft.TextField(label="Notes", value=part.notes, multiline=True, min_lines=2, max_lines=4, read_only=ro)

        def render_suggestions():
            have = {t.casefold() for t in core.split_tags(tags.value)}
            suggestions.controls = [] if ro else [ft.Chip(label=ft.Text(t), leading=ft.Icon(ft.Icons.ADD, size=14), on_click=add_tag(t))
                                                  for t in core.all_tags(team) if t.casefold() not in have]

        def add_tag(t):
            def handler(e):
                tags.value = ", ".join(core.split_tags(tags.value) + [t])
                render_suggestions()
                tags.update()
                suggestions.update()
            return handler

        def tags_changed(e):
            render_suggestions()
            suggestions.update()

        tags.on_change = tags_changed
        render_suggestions()

        # gates: edited on a copy, so Cancel leaves the part untouched
        work = core.Part(part.name, dates=dict(part.dates), passed=dict(part.passed), skipped=set(part.skipped),
                         started=set(part.started), template=part.template, parent=part.parent)
        options = team.available()
        template = ft.Dropdown(
            label="Gate template", value=core.template_of(team, part).name, disabled=ro, expand=True,
            options=[ft.DropdownOption(key=n, text=f"{n}  ({t.designation})" + ("  · default" if n == team.default else ""))
                     for n, t in options.items()])
        due_fields, pass_fields = {}, {}

        def fields_for(g):
            if g not in due_fields:
                due_fields[g] = date_field(self.page, "Due", part.dates.get(g), ro, on_change=lambda: refresh())
                pass_fields[g] = date_field(self.page, "Done on", work.passed[g][0] if g in work.passed else None, ro)
            return due_fields[g], pass_fields[g]

        gate_rows = ft.Column(spacing=0, tight=True)
        warn_box = ft.Container(padding=10, border_radius=8, visible=False, bgcolor=STATUS_COLOURS["Due soon"][0])

        def edited_dates():
            out = dict(part.dates)
            for g, fld in due_fields.items():
                try:
                    out[g] = fld.parse()
                except ValueError:
                    pass
            return out

        def render_gates():
            work.dates = edited_dates()
            work.template = template.value if template.value != team.default else ""
            tpl = core.template_of(team, work)
            rows = []
            for g in tpl.gates:
                due_f, done_f = fields_for(g)
                state = core.gate_state(work, team, g)
                icon, colour = {"done": (ft.Icons.CHECK_CIRCLE, ft.Colors.ON_SURFACE), "in progress": (ft.Icons.TIMELAPSE, ft.Colors.PRIMARY),
                                "ready": (ft.Icons.RADIO_BUTTON_CHECKED, ft.Colors.PRIMARY), "off": (ft.Icons.DO_NOT_DISTURB_ON_OUTLINED, ft.Colors.OUTLINE),
                                "waiting": (ft.Icons.RADIO_BUTTON_UNCHECKED, ft.Colors.OUTLINE)}[state]
                waiting_on = [a for a in tpl.after.get(g, []) if not core.cleared(work, a)]
                sub = {"done": "done", "in progress": "in progress", "ready": "ready to start", "off": "off: doesn't apply to this part",
                       "waiting": "after " + ", ".join(waiting_on)}[state]
                if state == "done" and lateness(work.passed[g][0], work.passed[g][1]):
                    sub += f" ({lateness(work.passed[g][0], work.passed[g][1])})"
                if tpl.due_before.get(g):
                    sub += f" · due before {tpl.due_before[g]}"
                picker = ft.Dropdown(value=core.state_name(work, g), dense=True, width=190, disabled=ro, on_select=set_state(g),
                                     options=[ft.DropdownOption(key=x, text=x) for x in core.STATES])
                fields = [] if state == "off" else [ft.Container(due_f.control, expand=1)]
                if state == "done":
                    fields.append(ft.Container(done_f.control, expand=1))
                rows.append(ft.Container(ft.Column([
                    ft.Row([ft.Icon(icon, color=colour, size=20),
                            ft.Column([ft.Text(g, weight=ft.FontWeight.W_600 if state in ("ready", "in progress") else None),
                                       ft.Text(sub, size=12, color=ft.Colors.ON_SURFACE_VARIANT)], spacing=0, tight=True, expand=True),
                            picker], vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    *([ft.Row(fields, spacing=8)] if fields else []),
                ], spacing=6, tight=True), padding=ft.Padding.symmetric(vertical=8),
                    border=ft.Border(bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT))))
            gate_rows.controls = rows
            warnings = core.rule_warnings(work, team, self.tracker.key_dates)
            warn_box.content = ft.Column([ft.Text(f"⚠ {w}", size=13, color=STATUS_COLOURS["Due soon"][1]) for w in warnings], spacing=2)
            warn_box.visible = bool(warnings)

        def refresh():
            render_gates()
            gate_rows.update()
            warn_box.update()

        def set_state(g):
            def handler(e):
                core.set_gate_state(work, g, e.control.value, today)
                if g in work.passed and not pass_fields[g].field.value:
                    pass_fields[g].field.value = today.isoformat()
                refresh()
            return handler

        template.on_select = lambda e: refresh()
        render_gates()

        def save(e):
            try:
                dates = {**part.dates, **{g: fld.parse() for g, fld in due_fields.items()}}
                new_passed = {g: (pass_fields[g].parse() if g in pass_fields else when, due_then) for g, (when, due_then) in work.passed.items()}
            except ValueError as err:
                self.toast(str(err), error=True)
                return
            new = core.Part(name.value.strip(), owner.value.strip(), "", (flag.selected or ["OK"])[0], notes.value.strip(), dates,
                            parent_field.value or "", core.split_names(assignees.value), core.split_tags(tags.value), new_passed,
                            set(work.skipped), template.value if template.value != team.default else "", set(work.started))
            problem = core.check_part(team, new, index)
            if problem:
                name.error = problem
                name.update()
                return

            def change():
                if index is None:
                    team.parts.append(new)
                else:
                    core.record_part_change(team, team.parts[index], new)
                    team.parts[index] = new
            self.commit(team, change)

        def delete(e):
            if kids:
                self.toast(f"{part.name} has {plural(len(kids), 'sub-part')}. Delete or move them first.", error=True)
                return
            self.commit(team, lambda: team.parts.pop(index))

        def add_sub_part(e):
            self.page.pop_dialog()
            self.edit_part(team, None, parent=part.name)

        actions = [ft.TextButton("Close" if ro else "Cancel", on_click=lambda e: self.page.pop_dialog())]
        if not ro:
            if index is not None:
                actions.insert(0, ft.TextButton("Delete", icon=ft.Icons.DELETE_OUTLINE, on_click=delete,
                                                style=ft.ButtonStyle(color=ft.Colors.ERROR)))
            actions.append(ft.FilledButton("Save", on_click=save))
            if index is not None and not part.parent:
                actions.insert(1, ft.TextButton("Add sub-part", icon=ft.Icons.ADD, on_click=add_sub_part))
        part_actions = [a for a in team.actions if a.part == part.key and a.closed is None] if index is not None else []
        extra = []
        if part_actions:
            extra += [ft.Text("Open actions", theme_style=ft.TextThemeStyle.TITLE_SMALL)] + [
                ft.Text(f"• {a.text} ({a.owner or 'no owner'}, due {fmt(a.due)})", size=13) for a in part_actions]
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(("Edit part" if index is not None else ("New sub-part" if parent else "New part")) if not ro else part.key),
            content=ft.Container(ft.Column([
                name, ft.Row([parent_field]), ft.Row([owner, assignees]), tags, suggestions, ft.Row([template]),
                ft.Text("Flag", size=12, color=ft.Colors.ON_SURFACE_VARIANT), flag,
                notes,
                warn_box,
                ft.Text("Gates", theme_style=ft.TextThemeStyle.TITLE_SMALL),
                gate_rows,
                *extra,
            ], tight=True, spacing=12, scroll=ft.ScrollMode.AUTO, horizontal_alignment=ft.CrossAxisAlignment.STRETCH), width=620),
            actions=actions,
        ))

    def edit_action(self, team, index):
        action = team.actions[index] if index is not None else core.Action()
        ro = not team.writable
        text = ft.TextField(label="Action", value=action.text, read_only=ro, autofocus=index is None, multiline=True, max_lines=3)
        part = ft.Dropdown(label="Part", value=action.part or None, disabled=ro,
                           options=[ft.DropdownOption(key=p.key, text=p.key) for p in team.parts])
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

    def edit_template(self, owner, name=None):
        """Build or edit a gate template: its gates (from the library, or new ones), their order, which gates each comes
        after (so gates can run in parallel), and the key date each must be due before. `owner` is a Subteam for its
        own templates, or None for the general ones in the master."""
        tr = self.tracker
        designation = owner.name if owner else core.GENERAL
        store = owner.templates if owner else tr.templates
        tpl = store.get(name) if name else None
        ro = not (owner.writable if owner else tr.master_writable)
        events = [""] + sorted({k.event for k in tr.key_dates if k.type != "Design Review"}, key=str.casefold)
        library = sorted({g.name for g, _ in core.library_for(tr, designation)}, key=str.casefold)
        title = ft.TextField(label="Template name", value=tpl.name if tpl else "", read_only=ro, autofocus=tpl is None)
        column = ft.Column(spacing=0, tight=True)
        error = ft.Text(color=ft.Colors.ERROR, size=13, visible=False)

        def make(old, gate, due=""):
            options = events + ([due] if due and due not in events else [])
            field = ft.TextField(value=gate, dense=True, expand=True, hint_text="Gate name", read_only=ro, on_blur=lambda e: refresh())
            pick = ft.PopupMenuButton(icon=ft.Icons.LIST, tooltip="Pick from the gate library", disabled=ro,
                                      items=[ft.PopupMenuItem(content=g, on_click=choose(field, g)) for g in library])
            return {"old": old, "after": set(), "field": field, "pick": pick,
                    "due": ft.Dropdown(value=due, dense=True, width=230, label="Due before", disabled=ro,
                                       options=[ft.DropdownOption(key=v, text=v or "— (no key date)") for v in options])}

        def choose(field, g):
            def handler(e):
                field.value = g
                refresh()
            return handler

        rows = [make(g, g, tpl.due_before.get(g, "")) for g in tpl.gates] if tpl else []
        by_name = {r["old"]: r for r in rows}
        for r in rows:
            r["after"] = {id(by_name[a]) for a in tpl.after.get(r["old"], []) if a in by_name}

        def render():
            controls = []
            for i, r in enumerate(rows):
                earlier = rows[:i]
                r["after"] &= {id(x) for x in earlier}  # a gate can only come after gates above it
                chips = [ft.Chip(label=ft.Text(x["field"].value or "(unnamed)", size=12), selected=id(x) in r["after"],
                                 on_select=toggle(r, x), disabled=ro) for x in earlier]
                controls.append(ft.Container(ft.Column([
                    ft.Row([ft.Text(f"{i + 1}.", width=22, color=ft.Colors.ON_SURFACE_VARIANT), r["field"], r["pick"],
                            ft.IconButton(ft.Icons.ARROW_UPWARD, tooltip="Move up", disabled=ro or i == 0, on_click=move(i, -1)),
                            ft.IconButton(ft.Icons.ARROW_DOWNWARD, tooltip="Move down", disabled=ro or i == len(rows) - 1, on_click=move(i, 1)),
                            ft.IconButton(ft.Icons.DELETE_OUTLINE, tooltip="Remove", disabled=ro, on_click=remove(i))],
                           spacing=4, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                    ft.Container(ft.Row([r["due"], ft.Text("Comes after:", size=12, color=ft.Colors.ON_SURFACE_VARIANT)]
                                        + (chips or [ft.Text("nothing (can start any time)", size=12, italic=True,
                                                             color=ft.Colors.ON_SURFACE_VARIANT)]),
                                        wrap=True, spacing=6, run_spacing=4, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                                 padding=ft.Padding.only(left=26)),
                ], spacing=6, tight=True), padding=ft.Padding.symmetric(vertical=10),
                    border=ft.Border(bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT))))
            column.controls = controls

        def refresh():
            render()
            column.update()

        def toggle(r, x):
            def handler(e):
                r["after"] ^= {id(x)}
                refresh()
            return handler

        def move(i, step):
            def handler(e):
                rows[i], rows[i + step] = rows[i + step], rows[i]
                refresh()
            return handler

        def remove(i):
            def handler(e):
                rows.pop(i)
                refresh()
            return handler

        def add(e):
            r = make(None, "")
            r["after"] = {id(rows[-1])} if rows else set()
            rows.append(r)
            refresh()

        def reset(e):
            names = tpl.gates if tpl else []
            rows[:] = [make(g if g in names else None, g) for g in core.DEFAULT_GATES]
            for prev, r in zip(rows, rows[1:]):
                r["after"] = {id(prev)}
            refresh()

        def save(e):
            new_name = title.value.strip()
            names = [r["field"].value.strip() for r in rows]
            problem = core.check_template_name(tr, new_name, tpl.name if tpl else None) or core.check_gates(names)
            if problem:
                error.value, error.visible = problem, True
                error.update()
                return
            name_of = {id(r): n for r, n in zip(rows, names)}
            pairs = [(r["old"], n) for r, n in zip(rows, names)]
            after = {n: [name_of[id(x)] for x in rows if id(x) in r["after"]] for r, n in zip(rows, names)}
            due = {n: r["due"].value or "" for r, n in zip(rows, names)}
            old_name = tpl.name if tpl else None
            target = tpl or core.Template(new_name, designation)

            def apply(teams):
                core.update_template(target, pairs, after, due)
                if old_name and old_name != new_name:
                    del store[old_name]
                    for t in teams:
                        for p in t.parts:
                            if p.template == old_name:
                                p.template = new_name
                        if t.default_template == old_name:
                            t.default_template = new_name
                target.name = new_name
                store[new_name] = target
                for t in teams:
                    t.shared = dict(tr.templates)
                    core.remap_parts(t, new_name, pairs)
                core.ensure_library(owner or tr, target, () if owner is None else tr.library)

            if owner:
                self.commit(owner, lambda: apply([owner]))
            else:
                self.commit_master(lambda: apply(tr.subteams))

        def delete(e):
            users = core.template_users(tr, tpl.name)
            if users:
                self.toast(f"{plural(users, 'part')} follow {tpl.name}. Move them to another template first.", error=True)
                return
            if owner:
                self.commit(owner, lambda: store.pop(tpl.name))
            else:
                self.commit_master(lambda: store.pop(tpl.name))

        render()
        actions = [ft.TextButton("Close" if ro else "Cancel", on_click=lambda e: self.page.pop_dialog())]
        if not ro:
            if tpl:
                actions.insert(0, ft.TextButton("Delete template", icon=ft.Icons.DELETE_OUTLINE, on_click=delete,
                                                style=ft.ButtonStyle(color=ft.Colors.ERROR)))
            actions.append(ft.FilledButton("Save", on_click=save))
        note = ("A gate starts once every gate it comes after is done (or off), so gates with the same prerequisites run in "
                "parallel; doing them out of order shows as a conflict. 'Due before' ties a gate to a key date. Renaming a "
                "gate keeps its dates and history; removing one deletes them.")
        if not owner:
            note += " This is a general template: other sub-teams' files pick up changes when their leads next save."
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(f"{'Edit' if tpl else 'New'} template · {designation}"),
            content=ft.Container(ft.Column([
                title,
                ft.Text(note, size=13, color=ft.Colors.ON_SURFACE_VARIANT),
                column,
                ft.Row([ft.TextButton("Add gate", icon=ft.Icons.ADD, on_click=add, disabled=ro),
                        ft.TextButton("Reset to standard gates", icon=ft.Icons.RESTART_ALT, on_click=reset, disabled=ro,
                                      tooltip=" → ".join(core.DEFAULT_GATES))]),
                error,
            ], tight=True, spacing=12, scroll=ft.ScrollMode.AUTO), width=720),
            actions=actions,
        ))

    def commit_master(self, change):
        """Apply a change to the general library / templates and save the master; reload on failure."""
        change()
        try:
            core.save_master_gates(self.tracker)
        except core.SaveError as e:
            self.page.pop_dialog()
            self.toast(str(e), error=True)
            self.reload()
            return
        for team in self.tracker.subteams:  # bring the files you can edit up to date now; the rest update when saved
            if team.writable:
                with contextlib.suppress(core.SaveError):
                    core.save_subteam(team)
        self.page.pop_dialog()
        self.toast("Saved to Master.xlsx")
        self.reload()

    # ---- gate library and templates page

    def gates_view(self):
        tr = self.tracker
        templates = sorted(core.all_templates(tr).items(), key=lambda kv: (kv[1][1] is not None, kv[1][0].designation, kv[0].casefold()))
        rows = []
        for name, (t, owner) in templates:
            chain = " → ".join(core.code_of(tr, g) or g for g in t.gates)
            users = core.template_users(tr, name)
            rows.append(ft.ListTile(
                leading=ft.Icon(ft.Icons.ACCOUNT_TREE_OUTLINED, color=ft.Colors.ON_SURFACE_VARIANT),
                title=ft.Text(name), subtitle=ft.Text(f"{t.designation} · {chain}"),
                trailing=pill(plural(users, "item"), "No date"),
                on_click=lambda e, o=owner, n=name: self.edit_template(o, n)))
        editable = ([core.GENERAL] if tr.master_writable else []) + [t.name for t in tr.subteams if t.writable]
        new_template = ft.TextButton("New template", icon=ft.Icons.ADD, on_click=lambda e: self.pick_designation(
            "New template", lambda d: self.edit_template(None if d == core.GENERAL else self.team(d)))) if editable else None
        lib_rows = [self.library_row(g, owner) for g, owner in sorted(core.library_for(tr), key=lambda go: (go[0].designation != core.GENERAL,
                                                                                                          go[0].designation, go[0].code))]
        add_gate = ft.TextButton("Add gate", icon=ft.Icons.ADD, on_click=lambda e: self.pick_designation(
            "Add gate", lambda d: self.edit_gate_def(None if d == core.GENERAL else self.team(d), None))) if editable else None
        head = ft.Container(ft.Row([ft.Text(h, weight=ft.FontWeight.W_600, width=w) for h, w in
                                    (("Code", 70), ("Gate", 240), ("Designation", 150), ("Used in", 170))]
                                   + [ft.Text("Description", weight=ft.FontWeight.W_600, expand=True)], spacing=16),
                            padding=ft.Padding.symmetric(horizontal=12, vertical=10),
                            border=ft.Border(bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT)))
        return ft.Column([
            heading("Gates", "Gate templates are the workflows parts move through; they're built from the gates in the library. "
                             "General ones are kept in the master (PM); sub-team ones in that sub-team's file (its lead)."),
            section(f"Templates ({len(templates)})", ft.Column(rows) if rows else empty("No templates yet."), new_template),
            section(f"Gate library ({len(lib_rows)})", ft.Column([head] + lib_rows, spacing=0) if lib_rows else empty("No gates yet."), add_gate),
        ], spacing=20, scroll=ft.ScrollMode.AUTO), None

    def library_row(self, g, owner):
        names, items = core.gate_users(self.tracker, g.name)
        used = f"{plural(len(names), 'template')} · {plural(items, 'item')}"
        return ft.Container(ft.Row([
            ft.Text(g.code or "—", width=70, weight=ft.FontWeight.W_600), ft.Text(g.name, width=240), ft.Text(g.designation, width=150),
            ft.Text(used, width=170, tooltip=", ".join(names) or None, color=ft.Colors.ON_SURFACE_VARIANT),
            ft.Text(g.description, expand=True, max_lines=2, overflow=ft.TextOverflow.ELLIPSIS),
        ], spacing=16), padding=ft.Padding.symmetric(horizontal=12, vertical=10), ink=True,
            border=ft.Border(bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT)), on_click=lambda e: self.edit_gate_def(owner, g))

    def pick_designation(self, title, then):
        """Ask where a new template / gate goes: General (the master) or a sub-team you can edit."""
        choices = ([core.GENERAL] if self.tracker.master_writable else []) + [t.name for t in self.tracker.subteams if t.writable]
        if len(choices) == 1:
            then(choices[0])
            return
        dd = ft.Dropdown(label="Designation", value=choices[0], options=[ft.DropdownOption(key=c, text=c) for c in choices])

        def go(e):
            self.page.pop_dialog()
            then(dd.value)
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(title), content=ft.Container(ft.Column([
                ft.Text("General: every sub-team can use it (kept in the master). Or a sub-team's own.", size=13,
                        color=ft.Colors.ON_SURFACE_VARIANT), dd], tight=True, spacing=12), width=420),
            actions=[ft.TextButton("Cancel", on_click=lambda e: self.page.pop_dialog()), ft.FilledButton("Next", on_click=go)]))

    def edit_gate_def(self, owner, gate):
        tr = self.tracker
        ro = not (owner.writable if owner else tr.master_writable)
        lib = owner.library if owner else tr.library
        names, items = core.gate_users(tr, gate.name) if gate else ([], 0)
        taken = {g.code for g, _ in core.library_for(tr)} - ({gate.code} if gate else set())
        name = ft.TextField(label="Gate", value=gate.name if gate else "", read_only=ro or bool(names), autofocus=gate is None,
                            tooltip=f"Used in {', '.join(names)}; rename it in those templates instead" if names else None)
        code = ft.TextField(label="Code", value=gate.code if gate else "", read_only=ro, width=140, hint_text="e.g. CON")
        desc = ft.TextField(label="Description", value=gate.description if gate else "", read_only=ro, multiline=True, min_lines=2, max_lines=4)
        error = ft.Text(color=ft.Colors.ERROR, size=13, visible=False)

        def save(e):
            n, c = name.value.strip(), code.value.strip().upper() or core.auto_code(name.value, taken)
            others = {g.name.casefold() for g, _ in core.library_for(tr) if g is not gate}
            problem = core.check_gates([n]) or (f"There's already a gate called {n}." if n.casefold() in others else None) or (
                f"Code {c} is already used." if c in taken else None)
            if problem:
                error.value, error.visible = problem, True
                error.update()
                return

            def change():
                g = gate or core.GateDef(n, designation=owner.name if owner else core.GENERAL)
                g.name, g.code, g.description = n, c, desc.value.strip()
                if gate is None:
                    lib.append(g)
            if owner:
                self.commit(owner, change)
            else:
                self.commit_master(change)

        def delete(e):
            if owner:
                self.commit(owner, lambda: lib.remove(gate))
            else:
                self.commit_master(lambda: lib.remove(gate))

        actions = [ft.TextButton("Close" if ro else "Cancel", on_click=lambda e: self.page.pop_dialog())]
        if not ro:
            if gate and not names:
                actions.insert(0, ft.TextButton("Delete", icon=ft.Icons.DELETE_OUTLINE, on_click=delete, style=ft.ButtonStyle(color=ft.Colors.ERROR)))
            actions.append(ft.FilledButton("Save", on_click=save))
        used = f"Used in {plural(len(names), 'template')} ({', '.join(names)}) · {plural(items, 'item')}" if names else "Not used in any template yet."
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(f"{'Gate' if gate else 'New gate'} · {owner.name if owner else core.GENERAL}"),
            content=ft.Container(ft.Column([ft.Row([code, name]), desc, ft.Text(used, size=13, color=ft.Colors.ON_SURFACE_VARIANT), error],
                                           tight=True, spacing=12), width=560),
            actions=actions))

    # ---- grid view

    def grid_view(self, team, today):
        """One grid per template in use: a row per assembly / sub-assembly, a box per gate. Right-click (or click) a box
        to set Not started / In progress / Done / Off."""
        if not team.parts:
            return [section("Grid", empty("No parts yet." + (" Press + to add one." if team.writable else "")))]
        nested = {id(c) for p in team.parts for c in core.children(team, p)}
        ordered = [x for p in team.parts if id(p) not in nested for x in [p] + core.children(team, p)]
        groups = {}
        for p in ordered:
            groups.setdefault(core.template_of(team, p).name, []).append(p)
        out = []
        for name, parts in groups.items():
            tpl = core.template_of(team, parts[0])
            head = ft.Row([ft.Container(ft.Text("Part", weight=ft.FontWeight.W_600), width=280)] + [
                ft.Container(ft.Text(core.code_of(self.tracker, g) or g[:6], size=12, weight=ft.FontWeight.W_600, text_align=ft.TextAlign.CENTER),
                             width=64, tooltip=g, alignment=ft.Alignment.CENTER) for g in tpl.gates], spacing=6)
            rows = [head] + [self.grid_row(team, p, tpl, today) for p in parts]
            out.append(section(f"{name}  ·  {tpl.designation}", ft.Row([ft.Column(rows, spacing=4)], scroll=ft.ScrollMode.AUTO),
                               ft.Text("Right-click a box to change it", size=12, color=ft.Colors.ON_SURFACE_VARIANT)))
        return out

    def grid_row(self, team, p, tpl, today):
        idx = team.parts.index(p)
        kids = core.children(team, p)
        label = [ft.Text(p.name, weight=None if p.parent else ft.FontWeight.W_600, max_lines=1, overflow=ft.TextOverflow.ELLIPSIS)]
        if kids and core.is_done(p, team) and core.open_gates(p, team):
            label.append(pill("Done: all sub-assemblies", "Done"))
        name = ft.Container(ft.Row([ft.Container(width=20 if p.parent else 0)] + label, spacing=6, tight=True), width=280,
                            on_click=lambda e: self.edit_part(team, idx), ink=True, tooltip=p.key)
        clash = {c.split(" is ")[0] for c in core.conflicts(p, team)}
        return ft.Row([name] + [self.grid_cell(team, p, idx, g, g in clash, today) for g in tpl.gates], spacing=6,
                      vertical_alignment=ft.CrossAxisAlignment.CENTER)

    def grid_cell(self, team, p, idx, g, clash, today):
        state = core.state_name(p, g)
        symbol, kind = GRID_STATE[state]
        due = p.dates.get(g)
        late = state in (core.NOT_STARTED, core.IN_PROGRESS) and due and due < today
        bg = {"done": ft.Colors.ON_SURFACE, "progress": STATUS_COLOURS["Due soon"][0], "off": ft.Colors.SURFACE_CONTAINER_HIGHEST,
              "empty": None}[kind]
        fg = {"done": ft.Colors.SURFACE, "progress": STATUS_COLOURS["Due soon"][1]}.get(kind, ft.Colors.ON_SURFACE_VARIANT)
        ready = core.gate_state(p, team, g) == "ready"
        # thick red = late or out of order; light = ready to start (red is reserved for problems)
        border = ft.Colors.ERROR if (late or clash) else (ft.Colors.ON_SURFACE_VARIANT if ready else ft.Colors.OUTLINE_VARIANT)
        tip = f"{g}: {state}" + (f" · due {fmt(due)}" if due else "") + (
            f" · done {fmt(p.passed[g][0])}" if g in p.passed and p.passed[g][0] else "") + (" · LATE" if late else "") + (
            "\nConflict: an earlier gate isn't done" if clash else "")
        box = ft.Container(ft.Text(symbol + ("!" if clash else ""), color=fg, weight=ft.FontWeight.W_700, text_align=ft.TextAlign.CENTER),
                           width=64, height=30, bgcolor=bg, border_radius=6, border=ft.Border.all(2.5 if (late or clash) else (1.5 if ready else 1), border),
                           alignment=ft.Alignment.CENTER, tooltip=tip)
        if not team.writable:
            return box
        items = [ft.PopupMenuItem(content=x, checked=x == state, on_click=self.grid_set(team, idx, g, x)) for x in core.STATES]
        items.append(ft.PopupMenuItem(content="Dates and details…", on_click=lambda e: self.edit_part(team, idx)))
        return ft.ContextMenu(content=box, primary_items=items, secondary_items=items, primary_trigger=ft.ContextMenuTrigger.DOWN)

    def grid_set(self, team, idx, gate, state):
        def handler(e):
            self.commit(team, lambda: core.set_gate_state(team.parts[idx], gate, state, date.today()), dialog=False)
        return handler

    # ---- design reviews view

    def reviews_view(self, team, today):
        tops = [p for p in team.parts if not p.parent]
        if not tops:
            return section("Design reviews", empty("No assemblies yet."))
        head = ft.Container(ft.Row([ft.Text("Assembly", weight=ft.FontWeight.W_600, expand=2)] +
                                   [ft.Text(r, weight=ft.FontWeight.W_600, expand=1) for r in core.DR_TYPES], spacing=16),
                            padding=ft.Padding.symmetric(horizontal=12, vertical=10), border=ft.Border(bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT)))
        rows = [head]
        for p in tops:
            by_type = {r.review: r for r in team.reviews if r.part == p.key}
            cells = []
            for kind in core.DR_TYPES:
                r = by_type.get(kind)
                if r and (r.date or r.outcome):
                    st = core.review_status(r, today)
                    cells.append(ft.Container(ft.Column([ft.Text(fmt(r.date)), pill(r.outcome or ("Needs outcome" if st == "Needs outcome" else st),
                                                                                       "Late" if st == "Needs outcome" else ("Done" if r.outcome else st))],
                                                        spacing=4, tight=True), expand=1, on_click=self.open_review(team, p.key, kind), ink=True))
                else:
                    cells.append(ft.Container(ft.Text("+ set date" if team.writable else "—", color=ft.Colors.ON_SURFACE_VARIANT), expand=1,
                                              on_click=self.open_review(team, p.key, kind), ink=True))
            rows.append(ft.Container(ft.Row([ft.Column([ft.Text(p.name, weight=ft.FontWeight.W_600),
                                                         ft.Text(p.owner or "", size=12, color=ft.Colors.ON_SURFACE_VARIANT)],
                                                        spacing=0, tight=True, expand=2)] + cells, spacing=16,
                                            vertical_alignment=ft.CrossAxisAlignment.CENTER),
                                     padding=ft.Padding.symmetric(horizontal=12, vertical=8), border=ft.Border(bottom=ft.BorderSide(1, ft.Colors.OUTLINE_VARIANT))))
        return section("Design reviews", ft.Column(rows, spacing=0))

    def open_review(self, team, key, kind):
        return lambda e: self.edit_review(team, key, kind)

    def edit_review(self, team, key, kind):
        r = next((x for x in team.reviews if x.part == key and x.review == kind), None) or core.Review(key, kind)
        ro = not team.writable
        when = date_field(self.page, "Date", r.date, ro)
        outcome = ft.Dropdown(label="Outcome", value=r.outcome or "", disabled=ro,
                              options=[ft.DropdownOption(key="", text="— (not held yet)")] + [ft.DropdownOption(key=o, text=o) for o in core.OUTCOMES])
        presenter = ft.TextField(label="Presenter", value=r.presenter, read_only=ro)
        notes = ft.TextField(label="Notes", value=r.notes, multiline=True, min_lines=2, max_lines=4, read_only=ro)

        def save(e):
            try:
                d = when.parse()
            except ValueError as err:
                self.toast(str(err), error=True)
                return
            new = core.Review(key, kind, d, outcome.value or "", presenter.value.strip(), notes.value.strip())

            def change():
                team.reviews = [x for x in team.reviews if not (x.part == key and x.review == kind)] + [new]
            self.commit(team, change)

        actions = [ft.TextButton("Close" if ro else "Cancel", on_click=lambda e: self.page.pop_dialog())]
        if not ro:
            actions.append(ft.FilledButton("Save", on_click=save))
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(f"{kind} · {key}"),
            content=ft.Container(ft.Column([when.control, outcome, presenter, notes], tight=True, spacing=12,
                                           horizontal_alignment=ft.CrossAxisAlignment.STRETCH), width=460),
            actions=actions))

    # ---- cars and sub-teams

    def cars_dialog(self):
        cars = core.cars_near(self.folder)
        tiles = [ft.ListTile(leading=ft.Icon(ft.Icons.CHECK if path == self.folder else ft.Icons.DIRECTIONS_CAR_OUTLINED),
                             title=ft.Text(name, weight=ft.FontWeight.W_600 if path == self.folder else None), subtitle=ft.Text(str(path)),
                             on_click=self.open_car(path)) for name, path in cars]
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text("Cars"),
            content=ft.Container(ft.Column(tiles, tight=True, scroll=ft.ScrollMode.AUTO), width=560),
            actions=[ft.TextButton("New car…", icon=ft.Icons.ADD, on_click=lambda e: self.new_car_dialog()),
                     ft.TextButton("Close", on_click=lambda e: self.page.pop_dialog())]))

    def open_car(self, path):
        def handler(e):
            self.page.pop_dialog()
            if path != self.folder and self.on_open_folder:
                self.on_open_folder(path)
        return handler

    def new_car_dialog(self):
        self.page.pop_dialog()
        name = ft.TextField(label="Car", value=core.next_car_name(self.tracker.car), autofocus=True)
        structure = ft.Switch(label="Sub-teams' assembly structure (parts, sub-parts, templates, tags)", value=True)
        gates = ft.Switch(label="Gate library and templates", value=True)
        dates = ft.Switch(label="Key dates (not marked done)", value=True)
        shift = ft.Switch(label="…moved one year later", value=True)
        error = ft.Text(color=ft.Colors.ERROR, size=13, visible=False)

        def create(e):
            try:
                dest = core.new_car(self.tracker, name.value, structure.value, gates.value, dates.value, 364 if shift.value else 0)
            except (ValueError, OSError, core.SaveError) as err:
                error.value, error.visible = str(err), True
                error.update()
                return
            self.page.pop_dialog()
            if self.on_open_folder:
                self.on_open_folder(dest)
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text(f"Start the next car after {self.tracker.car}"),
            content=ft.Container(ft.Column([
                name, ft.Text(f"Creates a new tracker folder next to this one ({self.folder.parent}). Carry over:", size=13,
                              color=ft.Colors.ON_SURFACE_VARIANT),
                structure, gates, dates, ft.Container(shift, padding=ft.Padding.only(left=24)),
                ft.Text("Owners, dates and progress always start fresh.", size=12, color=ft.Colors.ON_SURFACE_VARIANT), error,
            ], tight=True, spacing=10), width=560),
            actions=[ft.TextButton("Cancel", on_click=lambda e: self.page.pop_dialog()), ft.FilledButton("Create", on_click=create)]))

    def new_subteam_dialog(self):
        tr = self.tracker
        available = core.all_templates(tr)
        name = ft.TextField(label="Sub-team name", autofocus=True, hint_text="e.g. Cooling")
        tpl = ft.Dropdown(label="Starts with gate template", value=core.DEFAULT_TEMPLATE if core.DEFAULT_TEMPLATE in available else next(iter(available), ""),
                          options=[ft.DropdownOption(key=n, text=f"{n}  ({t.designation})") for n, (t, _) in available.items()])
        error = ft.Text(color=ft.Colors.ERROR, size=13, visible=False)

        def create(e):
            try:
                core.create_subteam(tr, name.value, tpl.value or "")
            except (ValueError, OSError, core.SaveError) as err:
                error.value, error.visible = str(err) if not isinstance(err, OSError) else "Edit failed.", True
                error.update()
                return
            self.page.pop_dialog()
            self.reload()
            self.show(name.value.strip())
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text("New sub-team"),
            content=ft.Container(ft.Column([name, tpl, ft.Text(
                "Creates Subteams/<name>.xlsx. On SharePoint, give its lead Edit access to that file.", size=12,
                color=ft.Colors.ON_SURFACE_VARIANT), error], tight=True, spacing=12), width=460),
            actions=[ft.TextButton("Cancel", on_click=lambda e: self.page.pop_dialog()), ft.FilledButton("Create", on_click=create)]))

    def commit(self, team, change, dialog=True):
        """Apply a change to a subteam and save it; on failure undo it (and reload if someone else saved first)."""
        backup = copy.deepcopy((team.parts, team.actions, team.reviews, team.templates, team.library, team.default_template))
        change()
        try:
            core.save_subteam(team)
        except core.SaveError as e:
            team.parts, team.actions, team.reviews, team.templates, team.library, team.default_template = backup
            if dialog:
                self.page.pop_dialog()
            self.toast(str(e), error=True)
            self.reload()  # picks up someone else's save, or shows the file as View only
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

        sub = "Competition deadlines, design reviews, milestones, freezes, blackouts and events, and who's responsible." + (
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
        kind = ft.Dropdown(label="Type", value=kd.type if kd.type in core.KEY_DATE_TYPES else "Deadline", expand=True,
                           options=[ft.DropdownOption(key=t, text=t) for t in core.KEY_DATE_TYPES])
        track = ft.Dropdown(label="Track", value=kd.track or "", expand=True,
                            options=[ft.DropdownOption(key="", text="—")] + [ft.DropdownOption(key=t, text=t) for t in core.TRACKS])
        when = date_field(self.page, "Date (official)", kd.date, False)
        end = date_field(self.page, "End date (for a period)", kd.end, False)
        internal = date_field(self.page, "Internal deadline", kd.internal, False)
        competition = ft.TextField(label="Competition", value=kd.competition, expand=True)
        people = ft.TextField(label="People responsible", value=", ".join(kd.people), hint_text="e.g. Sam, Alex")
        notes = ft.TextField(label="Notes", value=kd.notes, multiline=True, min_lines=2, max_lines=5)
        done = ft.Switch(label="Done (submitted / met)", value=kd.done)

        def save(e):
            try:
                d, e_, i_ = when.parse(), end.parse(), internal.parse()
            except ValueError as err:
                self.toast(str(err), error=True)
                return
            if not event.value.strip() or d is None:
                self.toast("An event needs a name and a date.", error=True)
                return
            if e_ and e_ < d:
                self.toast("The end date is before the start date.", error=True)
                return
            new = core.KeyDate(d, event.value.strip(), kind.value or "Deadline", core.split_names(people.value), done.value, i_, e_,
                               track.value or "", competition.value.strip(), notes.value.strip())
            self.commit_key_dates(lambda: self.tracker.key_dates.__setitem__(index, new) if index is not None
                                  else self.tracker.key_dates.append(new))

        actions = [ft.TextButton("Cancel", on_click=lambda e: self.page.pop_dialog()), ft.FilledButton("Save", on_click=save)]
        if index is not None:
            actions.insert(0, ft.TextButton("Delete", icon=ft.Icons.DELETE_OUTLINE, style=ft.ButtonStyle(color=ft.Colors.ERROR),
                                            on_click=lambda e: self.commit_key_dates(lambda: self.tracker.key_dates.pop(index))))
        self.page.show_dialog(ft.AlertDialog(
            title=ft.Text("Edit date" if index is not None else "New date"),
            content=ft.Container(ft.Column([
                event, ft.Row([kind, track]),
                ft.Row([ft.Container(when.control, expand=1), ft.Container(end.control, expand=1)]),
                ft.Row([ft.Container(internal.control, expand=1), competition]),
                people, notes, done,
            ], tight=True, spacing=12, scroll=ft.ScrollMode.AUTO, horizontal_alignment=ft.CrossAxisAlignment.STRETCH), width=600),
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
            self.reload()  # picks up someone else's save, or shows the master as View only
            return
        self.page.pop_dialog()
        self.toast("Key dates saved")
        self.show(self.view)


class date_field:
    """A text field (yyyy-mm-dd) with a calendar button."""

    def __init__(self, page, label, value, read_only, on_change=None):
        self.on_change = on_change
        self.field = ft.TextField(label=label, value=value.isoformat() if value else "", hint_text="yyyy-mm-dd",
                                  read_only=read_only, expand=True, dense=True, on_blur=lambda e: on_change and on_change())
        picker = ft.DatePicker(value=datetime.combine(value or date.today(), datetime.min.time()),
                               first_date=datetime(2020, 1, 1), last_date=datetime(2035, 12, 31), on_change=self._picked)
        button = ft.IconButton(ft.Icons.CALENDAR_MONTH, tooltip="Pick a date", disabled=read_only,
                               on_click=lambda e: page.show_dialog(picker))
        self.control = ft.Row([self.field, button], spacing=4)

    def _picked(self, e):
        v = e.control.value
        self.field.value = (v.date() if isinstance(v, datetime) else v).isoformat()
        self.field.update()
        if self.on_change:
            self.on_change()

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


SHORTCUT_HELP = ("Each member adds the team's SharePoint files to OneDrive once: in SharePoint, open UCalgary-Racing → Documents "
                 "and click 'Add shortcut to My files'. It then shows up in File Explorer under "
                 f"'OneDrive - University of Calgary\\{core.SHAREPOINT_FOLDER}'.")


def setup_screen(page, on_open, on_cancel=None):
    """Choose the tracker folder: pre-filled with the one found in the team's default location."""
    page.controls.clear()
    page.floating_action_button = None
    current = load_config().get("folder", "")
    found = core.find_folder()
    default = core.default_location()
    if found:
        intro = "Found the tracker here. Click Open to use it, or Browse to pick a different folder."
    else:
        intro = f"Couldn't find the tracker in the usual place ({default}). {SHORTCUT_HELP} Then choose the folder below."
    path = ft.TextField(label="Tracker folder (the one holding Master.xlsx)", expand=True,
                        value=current if on_cancel and current else str(found or default))
    picker = ft.FilePicker()

    async def browse(e):
        start = Path(path.value.strip().strip('"'))
        start = start if start.is_dir() else (default if default.is_dir() else (core.onedrive_roots() or [Path.home()])[0])
        chosen = await picker.get_directory_path(dialog_title="Choose the tracker folder", initial_directory=str(start))
        if chosen:
            path.value = chosen
            path.error = None
            path.update()

    def open_folder(e):
        folder = core.tracker_in(path.value.strip().strip('"'))
        if not folder:
            path.error = f"No {core.MASTER} in that folder (or the folder directly inside it)"
            path.update()
            return
        on_open(folder)

    buttons = [ft.FilledButton("Open", icon=ft.Icons.ARROW_FORWARD, on_click=open_folder)]
    if on_cancel:
        buttons.append(ft.TextButton("Cancel", on_click=lambda e: on_cancel()))
    page.add(ft.Container(ft.Column([
        ft.Icon(ft.Icons.FOLDER_OPEN, size=48, color=ft.Colors.PRIMARY),
        ft.Text("Choose the tracker folder", theme_style=ft.TextThemeStyle.HEADLINE_SMALL, weight=ft.FontWeight.W_600),
        ft.Text(intro, color=ft.Colors.ON_SURFACE_VARIANT),
        ft.Row([path, ft.IconButton(ft.Icons.FOLDER_OPEN, tooltip="Browse", on_click=browse)]),
        ft.Row(buttons, spacing=12),
        ft.Text(SHORTCUT_HELP if found else "", size=12, color=ft.Colors.OUTLINE),
    ], spacing=16, horizontal_alignment=ft.CrossAxisAlignment.START), width=900, padding=32))
    page.update()


def main(page: ft.Page, folder=None, start=None):
    page.title = "Project Tracker"
    page.theme, page.dark_theme = team_theme("light"), team_theme("dark")
    cfg = load_config()
    page.theme_mode = ft.ThemeMode.DARK if cfg.get("dark", True) else ft.ThemeMode.LIGHT  # dark by default, like the documents
    use_mode(page.theme_mode == ft.ThemeMode.DARK)
    page.padding = 0
    page.window.icon = str(ASSETS / "icon.ico")
    threading.Thread(target=sharpen_window_icon, args=(page.title, ASSETS / "icon.ico"), daemon=True).start()
    page.window.width, page.window.height = 1280, 820
    page.window.min_width, page.window.min_height = 720, 520
    page.window.alignment = ft.Alignment.CENTER  # the size and place it restores to
    page.window.maximized = True

    def open_tracker(path):
        cfg = load_config()
        cfg["folder"] = str(path)
        save_config(cfg)
        page.controls.clear()
        app = TrackerApp(page, path, on_change_folder=lambda: setup_screen(page, open_tracker, on_cancel=lambda: open_tracker(path)),
                         on_open_folder=open_tracker)
        if start:
            app.show(start)

    folder = folder or cfg.get("folder")
    if folder and (Path(folder) / core.MASTER).exists():
        open_tracker(Path(folder))
    else:
        setup_screen(page, open_tracker)  # first run (or the remembered folder has gone): let the user choose


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Project Tracker")
    ap.add_argument("--folder", help="the synced Project Tracker folder (remembered after the first run)")
    ap.add_argument("--start", help="open on this view: overview, mywork, keydates, or a subteam name")
    args = ap.parse_args()
    ft.run(lambda page: main(page, args.folder, args.start), assets_dir=str(ASSETS))
