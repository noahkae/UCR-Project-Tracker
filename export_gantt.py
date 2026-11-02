"""Export the season as a Gantt chart PDF, in the Pit Wall newsletter style.

    py export_gantt.py                                   # the tracker folder the app last opened
    py export_gantt.py --subteams Aero Suspension        # only some subteams
    py export_gantt.py --out "UCR-04 Gantt.pdf" --car UCR-04 --no-open

Each part is a row. Its gates are bars that run from the due date of the gates they come after
to their own due date: white = passed, red = being worked on, outline = still to come. Gates that
run in parallel share the row in thinner lanes; skipped gates aren't drawn.
Diamonds show when a gate was actually passed (from the Gate Log): white if on time,
red if late, with a line back to the due date. Dashed lines are key dates, with design
reviews in red. The Milestone key dates form the timeline in the header.
"""

import argparse
import os
import sys
from datetime import date, timedelta
from pathlib import Path

import tracker_core as core

try:
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.dates as mdates
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_pdf import PdfPages
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch, Rectangle
except ImportError:
    sys.exit("The Gantt export needs matplotlib. Run:  py -m pip install -r requirements.txt")

# Pit Wall palette
BG, FG, RED, MUTED = "#171717", "#FFFFFF", "#FF0000", "#D3C8C7"
GRID, FAINT = "#2C2C2C", "#5C5656"

PAGE_W, PAGE_H = 11.0, 8.5  # landscape US Letter, inches
MARGIN = 0.4
LABEL_W = 2.4  # part-name column
ROW_H = 0.25
STRIP_H = 0.95  # key-date strip under the chart
BOTTOM = 0.75  # legend + footer
FIRST_TOP = 2.15  # header + milestone timeline on page 1
OTHER_TOP = 0.85
LOGO = Path(__file__).resolve().parent / "assets" / "logo.png"

plt.rcParams.update({
    "font.family": "sans-serif",
    "font.sans-serif": ["Arial", "Liberation Sans", "DejaVu Sans"],
    "pdf.fonttype": 42,  # embed real fonts so text stays selectable
    "text.color": FG,
})


def inch(x, y):
    """Figure coordinates from inches measured from the top-left corner."""
    return x / PAGE_W, 1 - y / PAGE_H


# --------------------------------------------------------------------------- data


def build_rows(tracker, only):
    rows = []
    for team in tracker.subteams:
        if only and team.name.casefold() not in only:
            continue
        rows.append(("team", team, None))
        nested = {id(c) for p in team.parts for c in core.children(team, p)}
        ordered = [x for p in team.parts if id(p) not in nested for x in [p] + core.children(team, p)]  # sub-parts under their parent
        rows += [("part", team, p) for p in ordered] or [("empty", team, None)]
    return rows


def paginate(rows):
    pages, page, capacity = [], [], None
    for row in rows:
        if capacity is None:
            capacity = rows_per_page(first=not pages)
        # don't leave a subteam heading alone at the bottom of a page
        if len(page) >= capacity or (row[0] == "team" and len(page) >= capacity - 1):
            pages.append(page)
            page, capacity = [], rows_per_page(first=False)
        page.append(row)
    if page or not pages:
        pages.append(page)
    return pages


def rows_per_page(first):
    top = FIRST_TOP if first else OTHER_TOP
    return int((PAGE_H - top - 0.35 - STRIP_H - BOTTOM) / ROW_H)


def gate_spans(part, team):
    """(gate, start, due, lane, lanes) for every gate with a due date that isn't skipped. A bar starts at the latest due
    date of the gates it comes after; bars that overlap in time get their own lane."""
    spans = []
    tpl = core.template_of(team, part)
    for g in tpl.gates:
        due = part.dates.get(g)
        if due is None or g in part.skipped:
            continue
        before = [part.dates[a] for a in tpl.after.get(g, []) if part.dates.get(a) and a not in part.skipped]
        prev = max(before) if before else None
        start = prev if prev and prev < due else due - timedelta(days=14 if prev is None else 3)
        spans.append([g, start, due])
    lanes = []  # end date of the last bar in each lane
    for span in sorted(spans, key=lambda sp: sp[1]):
        lane = next((i for i, end in enumerate(lanes) if end <= span[1]), len(lanes))
        if lane == len(lanes):
            lanes.append(span[2])
        else:
            lanes[lane] = span[2]
        span.append(lane)
    return [(g, start, due, lane, max(len(lanes), 1)) for g, start, due, lane in spans]


def date_range(tracker, rows, start, end):
    dates = [k.date for k in tracker.key_dates if k.date]
    for kind, team, part in rows:
        if kind == "part":
            dates += [sp[1] for sp in gate_spans(part, team)] + [d for d in part.dates.values() if d]
            dates += [p for p, _ in part.passed.values() if p]
    dates.append(date.today())
    lo = start or min(dates) - timedelta(days=7)
    hi = end or max(dates) + timedelta(days=10)
    return lo, hi


# --------------------------------------------------------------------------- drawing


def header_full(fig, title, subtitle_bold, subtitle_rest, milestones, today):
    if LOGO.exists():
        logo = plt.imread(str(LOGO))
        w = 2.0
        h = w * logo.shape[0] / logo.shape[1]
        ax = fig.add_axes([*inch(MARGIN, 0.3 + h)[:1], inch(0, 0.3 + h)[1], w / PAGE_W, h / PAGE_H])
        ax.imshow(logo)
        ax.axis("off")
    right = 1 - MARGIN / PAGE_W
    fig.text(right, inch(0, 0.32)[1], title, ha="right", va="top", fontsize=30, fontweight="bold", color=FG)
    rest = fig.text(right, inch(0, 0.95)[1], subtitle_rest, ha="right", va="top", fontsize=13, color=RED)
    fig.canvas.draw()
    rest_w = rest.get_window_extent().width / fig.dpi / PAGE_W
    fig.text(right - rest_w, inch(0, 0.95)[1], subtitle_bold, ha="right", va="top", fontsize=13, fontweight="bold", color=RED)
    if milestones:
        milestone_timeline(fig, milestones, today)


def milestone_timeline(fig, milestones, today, y=1.55, max_items=6):
    """The Pit Wall phase line: red up to now, ringed dot for the next milestone, white after."""
    ms = sorted(milestones, key=lambda k: k.date)
    nxt = next((i for i, k in enumerate(ms) if k.date >= today), len(ms))
    first = max(0, min(nxt - 2, len(ms) - max_items))
    ms = ms[first:first + max_items]
    nxt -= first
    x0, x1 = 1.3, PAGE_W - 1.3
    xs = [x0 + (x1 - x0) * i / max(len(ms) - 1, 1) for i in range(len(ms))] if len(ms) > 1 else [PAGE_W / 2]
    ax = fig.add_axes([0, 0, 1, 1], facecolor="none")
    ax.set_xlim(0, PAGE_W)
    ax.set_ylim(PAGE_H, 0)
    ax.axis("off")
    if len(xs) > 1:
        ax.plot([xs[0], xs[-1]], [y, y], color=FG, lw=1.5, solid_capstyle="butt", zorder=1)
        upto = xs[min(nxt, len(xs) - 1)] if nxt < len(xs) else xs[-1]
        ax.plot([xs[0], upto], [y, y], color=RED, lw=3, solid_capstyle="butt", zorder=2)
    for i, (k, x) in enumerate(zip(ms, xs)):
        current = i == nxt
        if i < nxt:
            ax.scatter([x], [y], s=60, color=RED, zorder=3)
        elif current:
            ax.scatter([x], [y], s=110, color=RED, edgecolors=FG, linewidths=2, zorder=3)
        else:
            ax.scatter([x], [y], s=60, color=FG, zorder=3)
        ax.text(x, y + 0.2, k.date.strftime("%b %Y"), ha="center", va="top", fontsize=9, fontweight="bold", color=RED if current else FG)
        ax.text(x, y + 0.37, k.event, ha="center", va="top", fontsize=8.5, color=FG)


def header_compact(fig, title, right_text):
    fig.text(MARGIN / PAGE_W, inch(0, 0.3)[1], title, ha="left", va="top", fontsize=20, fontweight="bold", color=RED)
    fig.text(1 - MARGIN / PAGE_W, inch(0, 0.36)[1], right_text, ha="right", va="top", fontsize=9, color=MUTED)


def chart(fig, rows, top, lo, hi, key_dates, today):
    left = MARGIN + LABEL_W
    width = PAGE_W - left - MARGIN
    n = max(len(rows), 1)
    height = n * ROW_H
    ax = fig.add_axes([left / PAGE_W, 1 - (top + height) / PAGE_H, width / PAGE_W, height / PAGE_H], facecolor=BG)
    ax.set_xlim(mdates.date2num(lo), mdates.date2num(hi))
    ax.set_ylim(n, 0)
    style_time_axis(ax, top=True)
    days_per_inch = (hi - lo).days / width

    for i, (kind, team, part) in enumerate(rows):
        y = i + 0.5
        if kind == "team":
            ax.text(-LABEL_W / width, y, team.name, transform=ax.get_yaxis_transform(), ha="left", va="center",
                    fontsize=10, fontweight="bold", color=RED, clip_on=False)
            ax.axhline(i + 1, color=FAINT, lw=0.6)
            continue
        if kind == "empty":
            ax.text(-LABEL_W / width, y, "no parts yet", transform=ax.get_yaxis_transform(), ha="left", va="center",
                    fontsize=8, style="italic", color=MUTED, clip_on=False)
            continue
        name = f"  └ {part.name}" if part.parent else part.name
        label = name if len(name) <= 30 else name[:29] + "…"
        ax.text(-LABEL_W / width, y, label, transform=ax.get_yaxis_transform(), ha="left", va="center", fontsize=8.5, color=FG, clip_on=False)
        if part.owner:
            ax.text(-0.012, y, part.owner, transform=ax.get_yaxis_transform(), ha="right", va="center", fontsize=7, color=MUTED, clip_on=False)
        ax.axhline(i + 1, color=GRID, lw=0.4)
        draw_part(ax, team, part, y, days_per_inch)

    for k in key_dates:
        if k.date and lo <= k.date <= hi:
            review = k.type == "Design Review"
            ax.axvline(mdates.date2num(k.date), color=RED if review else MUTED, lw=0.8, ls=(0, (3, 3)), alpha=0.75 if review else 0.4, zorder=1)
    if lo <= today <= hi:
        ax.axvline(mdates.date2num(today), color=RED, lw=1.6, zorder=3.5)  # over the bars, under their labels
    return ax, top + height


def draw_part(ax, team, part, y, days_per_inch):
    active = core.working_gates(part, team)
    spans = gate_spans(part, team)
    if not spans:
        ax.text(ax.get_xlim()[0], y, "  no dates yet", va="center", fontsize=7, style="italic", color=FAINT)
    for g, start, due, lane, lanes in spans:
        if g in part.passed:
            face, edge, text = FG, FG, BG
        elif g in active:
            face, edge, text = RED, RED, FG
        else:
            face, edge, text = "none", MUTED, MUTED
        h = 0.6 / lanes
        x0, x1 = mdates.date2num(start), mdates.date2num(due)
        ax.add_patch(Rectangle((x0, y - 0.3 + lane * h), x1 - x0, h, facecolor=face, edgecolor=edge, lw=0.7, zorder=3))
        room = (due - start).days / days_per_inch  # bar width in inches
        name = g if len(g) * 0.055 < room - 0.06 else (g.split()[0] if len(g.split()[0]) * 0.055 < room - 0.06 else "")
        if name:
            ax.text((x0 + x1) / 2, y - 0.3 + (lane + 0.5) * h, name, ha="center", va="center", fontsize=6.5 if lanes == 1 else 5,
                    color=text, zorder=4, clip_on=True)
    for e in team.log:
        if e.part != part.key or not e.passed:
            continue
        late = bool(e.due and e.passed > e.due)
        px, top = mdates.date2num(e.passed), y - 0.3  # sit on the bar's top edge so gate names stay readable
        if e.due and e.passed != e.due:
            ax.plot([mdates.date2num(e.due), px], [top, top], color=RED if late else FG, lw=1.1, ls="-" if late else ":", zorder=5)
        ax.scatter([px], [top], marker="D", s=20, color=RED if late else FG, edgecolors=FG if late else BG, linewidths=0.8, zorder=6)


def style_time_axis(ax, top):
    for s in ax.spines.values():
        s.set_visible(False)
    ax.set_yticks([])
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
    ax.xaxis.set_minor_locator(mdates.WeekdayLocator(byweekday=mdates.MO))
    ax.tick_params(axis="x", which="both", length=0, labelsize=8, labelcolor=MUTED, top=top, labeltop=top, bottom=not top, labelbottom=not top)
    ax.grid(axis="x", which="major", color=FAINT, lw=0.6)
    ax.grid(axis="x", which="minor", color=GRID, lw=0.4)
    ax.set_axisbelow(True)


def key_date_strip(fig, top, lo, hi, key_dates, today):
    left = MARGIN + LABEL_W
    width = PAGE_W - left - MARGIN
    ax = fig.add_axes([left / PAGE_W, 1 - (top + 0.3) / PAGE_H, width / PAGE_W, 0.3 / PAGE_H], facecolor=BG)
    ax.set_xlim(mdates.date2num(lo), mdates.date2num(hi))
    ax.set_ylim(0, 1)
    ax.axis("off")
    ax.text(-LABEL_W / width, 0.5, "Key dates", transform=ax.transAxes, ha="left", va="center", fontsize=10, fontweight="bold", color=RED)
    days_per_inch = (hi - lo).days / width
    by_day = {}
    for k in sorted((k for k in key_dates if k.date and lo <= k.date <= hi), key=lambda k: (k.date, k.type != "Design Review")):
        by_day.setdefault(k.date, []).append(k)
    last_x, level = None, 0
    for day, ks in by_day.items():
        review = any(k.type == "Design Review" for k in ks)
        x = mdates.date2num(day)
        ax.scatter([x], [0.75], marker="D" if review else "o", s=26, color=RED if review else MUTED, zorder=3)
        # stagger labels that would sit on top of the previous one; past a second level, leave the dot unlabelled
        close = last_x is not None and (x - last_x) / days_per_inch < 0.22
        if close and level >= 1:
            continue
        level = level + 1 if close else 0
        last_x = x
        name = ks[0].event if len(ks[0].event) <= 26 else ks[0].event[:25] + "…"
        ax.text(x, 0.45 - level * 0.6, name + (f" +{len(ks) - 1}" if len(ks) > 1 else ""), rotation=35, ha="right", va="top",
                fontsize=6.5, color=FG if review else MUTED, rotation_mode="anchor")
    if lo <= today <= hi:
        ax.text(mdates.date2num(today), 1.0, "Today", ha="center", va="bottom", fontsize=7, fontweight="bold", color=RED)


def legend(fig):
    handles = [
        Patch(facecolor=FG, edgecolor=FG, label="Gate passed"),
        Patch(facecolor=RED, edgecolor=RED, label="Current gate"),
        Patch(facecolor="none", edgecolor=MUTED, label="Gate to come"),
        Line2D([], [], marker="D", ls="none", color=FG, markeredgecolor=BG, label="Passed on time"),
        Line2D([], [], marker="D", ls="none", color=RED, markeredgecolor=FG, label="Passed late"),
        Line2D([], [], color=RED, ls=(0, (3, 3)), label="Design review"),
        Line2D([], [], color=MUTED, ls=(0, (3, 3)), label="Other key date"),
        Line2D([], [], color=RED, lw=1.6, label="Today"),
    ]
    fig.legend(handles=handles, loc="lower center", bbox_to_anchor=(0.5, 0.45 / PAGE_H), ncol=len(handles), frameon=False,
               fontsize=7.5, labelcolor=MUTED, handlelength=1.6, columnspacing=1.4)


def footer(fig, left_text, page, pages):
    y = 0.22 / PAGE_H
    fig.text(MARGIN / PAGE_W, y, left_text, ha="left", va="bottom", fontsize=8, color=MUTED)
    fig.text(1 - MARGIN / PAGE_W, y, f"{page} / {pages}", ha="right", va="bottom", fontsize=8, color=MUTED)


# --------------------------------------------------------------------------- main


def export(folder, out, only=None, car=None, start=None, end=None, today=None):
    today = today or date.today()
    tracker = core.load(folder)
    car = car or tracker.car  # the name in Master > Settings (or the folder name)
    only = {s.casefold() for s in only} if only else None
    rows = build_rows(tracker, only)
    pages = paginate(rows)
    lo, hi = date_range(tracker, rows, start, end)
    milestones = [k for k in tracker.key_dates if k.type == "Milestone" and k.date]
    n_parts = sum(1 for r in rows if r[0] == "part")
    n_teams = sum(1 for r in rows if r[0] == "team")
    foot = f"Generated by the UCR Project Tracker on {today:%d %B %Y} · {n_parts} parts across {n_teams} subteams"
    month = today.strftime("%B %Y")

    with PdfPages(out, metadata={"Title": f"{car} Season Gantt", "Creator": "UCR Project Tracker"}) as pdf:
        for number, page_rows in enumerate(pages, 1):
            fig = plt.figure(figsize=(PAGE_W, PAGE_H), facecolor=BG)
            if number == 1:
                header_full(fig, "Season Gantt", month, f" · {car} · Gantt" if car else " · Gantt", milestones, today)
                top = FIRST_TOP + 0.35
            else:
                header_compact(fig, "Season Gantt", f"{month} · {car}" if car else month)
                top = OTHER_TOP + 0.35
            if page_rows:
                _, bottom = chart(fig, page_rows, top, lo, hi, tracker.key_dates, today)
                key_date_strip(fig, bottom + 0.08, lo, hi, tracker.key_dates, today)
            else:
                fig.text(0.5, 0.5, "No parts to show yet.", ha="center", fontsize=14, color=MUTED)
            legend(fig)
            footer(fig, foot, number, len(pages))
            pdf.savefig(fig, facecolor=BG)
            plt.close(fig)
    return len(pages), n_parts


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--folder", help="the synced Project Tracker folder (default: the one the app last used)")
    ap.add_argument("--out", help="PDF to write (default: Downloads/Gantt <date>.pdf)")
    ap.add_argument("--subteams", nargs="+", help="only these subteams")
    ap.add_argument("--car", help="car name for the title (default: the one in the tracker's master file)")
    ap.add_argument("--from", dest="start", type=date.fromisoformat, help="first date shown, yyyy-mm-dd")
    ap.add_argument("--to", dest="end", type=date.fromisoformat, help="last date shown, yyyy-mm-dd")
    ap.add_argument("--no-open", action="store_true", help="don't open the PDF afterwards")
    args = ap.parse_args()

    folder = Path(args.folder) if args.folder else core.default_folder()
    if not folder or not (folder / core.MASTER).exists():
        sys.exit("Couldn't find the tracker folder. Pass it with --folder \"...\\Project Tracker\".")
    downloads = Path.home() / "Downloads"
    out = Path(args.out) if args.out else (downloads if downloads.is_dir() else Path.cwd()) / f"Gantt {date.today():%Y-%m-%d}.pdf"
    pages, parts = export(folder, out, args.subteams, args.car, args.start, args.end)
    print(f"Wrote {out} ({pages} page{'s' if pages != 1 else ''}, {parts} parts)")
    if not args.no_open and hasattr(os, "startfile"):
        os.startfile(out)


if __name__ == "__main__":
    main()
