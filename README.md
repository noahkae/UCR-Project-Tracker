# UCR Project Tracker

A small project-management tool for UCalgary Racing. Each subteam's plan is **one Excel file on SharePoint
that only its lead can edit**, and a **desktop app** (Material Design, built with Python + Flet) shows every
subteam's plan, deadlines and design reviews in one place.

* **SharePoint handles permissions.** Leads edit their own subteam file, the PM edits the master, and everyone else can view.
* **The data is just Excel.** If the app ever breaks, the plan is still readable and editable in Excel.
* **There's no server, database, login screen or `.exe`.**

**How to use the app:** the short illustrated guide [`docs/Project Tracker - How to use it.pdf`](docs/Project%20Tracker%20-%20How%20to%20use%20it.pdf)
(for members and leads; setup is covered below).

---

## Contents

1. [Install (everyone)](#1-install-everyone)
2. [One-time setup (PM / captain)](#2-one-time-setup-pm--captain)
3. [Connect the app to the tracker](#3-connect-the-app-to-the-tracker)
4. [Using the app](#4-using-the-app)
5. [How status is worked out](#5-how-status-is-worked-out)
6. [Editing in Excel instead](#6-editing-in-excel-instead)
7. [Customising](#7-customising)
8. [Troubleshooting](#8-troubleshooting)
9. [Files in this repo](#9-files-in-this-repo)

---

## 1. Install (everyone)

You only do this once per computer. It's Windows-first; Mac/Linux work too, see the note at the end of this section.

**1. Install Python 3.10 or newer** from <https://www.python.org/downloads/>. In the installer:
   * tick **"Add python.exe to PATH"**;
   * leave **"py launcher"** ticked, because double-clicking the app depends on it.

**2. Download `UCR Project Tracker.pyzw`** from this repo's **Releases** page (right-hand side on GitHub). It's a single
   ~75 KB file with the whole app in it. Put it somewhere permanent, e.g. `Documents`.

**3. Double-click it.** The first time, it asks to install the Python packages it needs:
   * [Flet](https://flet.dev) (the UI), [openpyxl](https://openpyxl.readthedocs.io) (the Excel files) and [matplotlib](https://matplotlib.org) (the Gantt export);
   * click **Yes** and wait a minute or two for the progress window to finish;
   * it then opens on its own.

   After that it opens straight away: centred, and with no terminal window behind it. The first start can take 10–20 seconds while Flet unpacks.

> **Python isn't bundled.** The `.pyzw` is Python's own single-file app format, so it needs Python installed (step 1)
> but is otherwise self-contained. It isn't an `.exe`, so Windows doesn't show "unknown publisher" warnings for it.

> **Tip:** right-click the file → *Send to → Desktop (create shortcut)*. In the shortcut's *Properties → Change Icon* you can
> pick `%LOCALAPPDATA%\FSAETracker\assets-…\icon.ico` (created on the first run) to give it the UCR icon.

**Running from the source code instead** (for whoever maintains the app):
* `git clone` the repo, then run `py -m pip install -r requirements.txt` in the folder.
* Double-click **`Project Tracker.pyw`**, or run `py tracker_app.py` to see error output in the terminal.
* Mac / Linux: `python3 -m pip install -r requirements.txt`, then `python3 tracker_app.py`.

---

## 2. One-time setup (PM / captain)

**1. Create the tracker files.** Open a terminal in the folder holding `UCR Project Tracker.pyzw` (File Explorer → click the address bar → type `cmd` → Enter) and run:

   ```
   py "UCR Project Tracker.pyzw" setup --out "Project Tracker"
   ```

   (From the source code, `py setup_tracker.py --out "Project Tracker"` does the same.)

   This creates:

   ```
   Project Tracker/
       Master.xlsx              key dates + the All Parts / All Actions rollups
       Subteams/
           Aero.xlsx
           Chassis & Controls.xlsx
           Driverless.xlsx
           Electrical.xlsx
           Powertrain & Battery Pack.xlsx
           Suspension.xlsx
   ```

   Options:
   * `--subteams Aero Suspension ...` picks your own subteam list.
   * `--gates "DR01 - Concept" "DR02 - Detailed Design" ...` picks your own gate list (in order).

   The script refuses to overwrite an existing tracker.

**2. Upload the tracker to SharePoint**, into **UCalgary-Racing → Documents → Project Management**. Either put
   `Master.xlsx` and `Subteams/` straight into *Project Management*, or upload the whole `Project Tracker/` folder into it.
   The app finds either layout.

**3. Set permissions** with *… → Manage access* on each file (or folder):

   | Who | Access |
   |---|---|
   | Each subteam lead | **Edit** on their own `Subteams/<name>.xlsx` |
   | PM / captain | **Edit** on `Master.xlsx` (and the subteam files if they help plan) |
   | Everyone else | **View** on the whole library |

**4. Add the first key dates** (competition deadlines, design reviews) in the app's *Key dates* tab, or directly in `Master.xlsx`.

---

## 3. Connect the app to the tracker

**Once per person: add the shortcut.** In SharePoint, open **UCalgary-Racing → Documents** and click **Add shortcut to My files**.
The team's documents then appear in File Explorer under OneDrive:

```
C:\Users\<you>\OneDrive - University of Calgary\UCalgary-Racing - Documents\Project Management
```

**First start: choose the folder.** The first time the app opens, it shows *Choose the tracker folder*:
* It looks in the location above automatically. Your own user name is filled in from Windows, so this works for everyone.
* If the tracker is there, the path is already filled in, so just click **Open**.
* If it isn't (e.g. the shortcut hasn't been added yet), the screen explains how to add it.
* Either way you can **Browse** to any other folder instead.

The app remembers your choice. To switch later, click the **folder** button at the bottom of the left rail; hovering over it
shows the current folder.

When the app loads, it checks each file by opening it for writing (nothing is written). Files you only have View access to
show as **View only**, so you can't change them by accident. If a save is still refused, the app says **Edit failed**,
undoes the change and reloads.

---

## 4. Using the app

### Everyone

* **Overview** shows:
  * one card per subteam (progress, open actions, plus how many parts are late, flagged or due soon);
  * **Needs attention**: late, *At risk* or *Blocked* parts; overdue actions; and key dates that are overdue or due within 7 days;
  * **Next 14 days**: gates, actions and key dates coming up;
  * the **next design review**, in the subtitle.
* **My work**: pick your name once (it's remembered) to see **the parts you own or work on, your open actions and the key
  dates you're responsible for**, with anything late at the top. Key dates you're named on carry a **You** tag everywhere in the app.
* **Subteam tabs** (left rail) show:
  * the subteam's parts, sorted with problems first. The dots show each gate: filled = passed, ringed = being worked on
    (several at once when gates run in parallel), hollow grey = skipped, faint = waiting. ⚠ marks a part whose dates break a
    gate rule (hover to see which). Under the name are its **tags**; under the owner, anyone else working on it.
    Filter by text (part, person, tag or note) or click a **tag** chip. Click a part to see its details;
  * its **actions**, open ones first;
  * how many gates it has passed **on time**, in the subtitle.
* **Calendar** is a month view of every gate due date, action and key date.
  * Use the chips at the top to show or hide each subteam, key dates and actions. Your choice is remembered.
  * Colours: red = late, amber = due soon, green = on track, grey = done, blue = a gate still to come, purple = an upcoming design review.
  * Click an item to open it, or click a busy day to see everything on it. Use ‹ › to change month and *Today* to jump back.
* **Key dates** lists competition deadlines, design reviews (gate meetings), milestones, freezes, blackouts and events, with the
  official date, the team's **internal deadline**, track, competition, notes and the people responsible. Overdue ones stay
  visible until marked done. Turn on *Show past* to see the rest. The calendar shows internal deadlines (⏱) and spans
  blackouts and multi-day events across their days.
* **↻** reloads the files (do this after a lead says they've updated). **☾** switches between light and dark mode.

### Subteam leads (your tab shows "You can edit")

* **Edit gates** (top of your subteam page) changes the gates your parts move through: rename, add, remove or reorder them,
  or **Reset to default** (DR01 - Concept → DR02 - Detailed Design → DR03 - Final Design → Manufacturing → Testing).
  * **Comes after** (chips under each gate): the gates that must be passed or skipped before this one starts. Gates with
    the same prerequisites **run in parallel**, e.g. Aero's Mould Design only needs CFD Freeze, while Mounting Design
    needs CFD Freeze *and* Mounting Concept. A gate can only come after gates above it in the list.
  * **Due before**: a key date (e.g. *Chassis Freeze*) the gate must be due before. Parts whose date for that gate is later
    get a ⚠ warning. Gates due inside a *Blackout* key date get one too, as do gates due before a gate they come after.
  * Renamed gates keep their due dates and history; removing a gate deletes them.
* **Add part** (bottom-right button): give it a name, an **owner** (the one accountable person), anyone else **working on it**,
  any **tags** (e.g. *Development*, *Aero-Testing*; click a suggested tag to add it), then a **due date for each gate**.
* **Sub-parts:** to keep big assemblies tidy, open a part (e.g. Front Wing) and click **Add sub-part**, or set *Sub-part of*
  when editing. Parts with sub-parts show collapsed with a summary ("8 sub-parts · 1 need attention"); click the arrow to
  expand. An assembly sorts by its worst sub-part, so problems still float to the top. Sub-parts go one level deep.
* **Click a part** to edit it:
  * each **gate** has its own row: **Passed** on a gate being worked on records today as its pass date (edit the date if it
    was passed earlier), **Undo** un-passes it (and anything after it), and **Skip** marks a gate that doesn't apply to this
    part (e.g. no mould for a mounting bracket). The app shows whether each pass was early or late;
  * **Flag** marks the part as *At risk* or *Blocked*. It then appears under *Needs attention*, so raise it at the next meeting;
  * **Notes** is a one-line reason, e.g. "waiting on supplier".
* **Actions** (follow-ups from design reviews):
  * *Add action* on your subteam page. Give each action one owner, a due date, the part it relates to, and where it came from (e.g. "DR2");
  * tick the box when it's done. Overdue actions show under *Needs attention*, and the owner sees them under *My work*.
* **Weekly habit:** before the leads meeting, pass finished gates, update dates and flags, and tick off finished actions. It takes about 5 minutes.

### PM / captain

* **Key dates:** add, edit or delete entries.
  * Use the *Design Review* type for gate meetings (e.g. "DR2: Aero + Suspension").
  * Fill in **People responsible** (comma-separated names, spelt the same way as in the subteam files, e.g. "Sam Lee, Alex Kim").
    They'll see the date in *My work*, and it shows under *Needs attention* from 7 days before.
  * Set an **Internal deadline** when the team works to an earlier date than the official one; status and reminders follow it.
  * Use **Freeze** for design freezes (gates can be tied to them), **Blackout** with an *End date* for periods with no gate
    dates (e.g. finals), and **Event** for things that just happen. Add the **Track**, **Competition** and any **Notes**.
  * Tick **Done** once a deadline is submitted or a milestone met. Until then, a past deadline shows as **late**.
    Design reviews don't need ticking; once the date has passed they count as held.
* **`Master.xlsx > All Parts` and `All Actions`** are rebuilt automatically whenever someone with edit access to the master saves anything.
  They let people without the app see everything in the browser. Don't edit those sheets by hand.
* **In the meeting**, go through *Needs attention* first, then *Next 14 days*. Record each review's follow-ups as actions straight away.

### Gantt chart export (for look-backs and reports)

This makes a PDF Gantt chart of the season, in the same style as the *Pit Wall* newsletter. In a terminal in the folder holding the app, run:

```
py "UCR Project Tracker.pyzw" gantt
```

(From the source code: `py export_gantt.py`. Both take the same options.)

It reads the tracker folder the app last used, writes `Gantt <date>.pdf` to your Downloads folder, and opens it. Options:

| Option | Does |
|---|---|
| `--subteams Aero Suspension` | Only these subteams |
| `--from 2026-09-01 --to 2027-06-30` | Only this date range |
| `--car UCR-05` | Car name in the title (default `UCR-04`) |
| `--out "somewhere\file.pdf"` | Choose the output file |
| `--folder "...\Project Tracker"` | Use a different tracker folder |
| `--no-open` | Don't open the PDF afterwards |

How to read it:
* **Header timeline:** the *Milestone* key dates. Red up to now, a ringed dot for the next milestone, white after that.
* **Rows:** one per part, grouped by subteam.
* **Bars:** each gate is a bar from the due date of the gates it comes after to its own due date. White = passed, red = being
  worked on, outline = still to come. Gates that run in parallel share the part's row in thinner lanes; skipped gates aren't drawn.
* **◆ diamonds:** when a gate was *actually* passed, from the gate history. White = on time or early; red = late, with a line back to the due date. This is the "in hindsight" part.
* **Dashed lines:** key dates, with design reviews in red. The **Key dates** strip underneath names them.
* **Solid red line:** today.

---

## 5. How status is worked out

Each gate of a part is **passed**, **skipped**, **being worked on** (everything it comes after is passed or skipped), or
**waiting**. The part's *current gate* is the gate(s) being worked on, and its status uses the earliest due date among them.

| Status | Meaning |
|---|---|
| **Late** | today is past the due date of a gate being worked on |
| **Due soon** | a gate being worked on is due within 7 days |
| **On track** | they're due later than that |
| **No date** | the gates being worked on have no due dates yet |
| **Done** | every gate is passed or skipped |

**Flags** are separate from status and set by the lead:
* **OK**
* **At risk**: might slip; discuss it.
* **Blocked**: can't progress without help.

**Actions and key dates** follow the same Late / Due soon / On track rules, based on their own due date.
* An action is *Done* once it's ticked.
* A deadline, milestone or freeze is *Done* once the PM ticks it; until then it's measured against the internal deadline
  if there is one.
* A design review, blackout or event shows as *Now* while it's on and *Held* once it's over.

**On time** (subteam subtitle) counts the gates passed on or before their due date, out of all gates passed that had a due date.

---

## 6. Editing in Excel instead

Leads can also edit their file directly in Excel (desktop or browser). Each subteam file has four sheets:

* **Plan**: the parts (below). This is the one you'd normally edit.
* **Actions**: `Action | Part | Owner | Due | From | Closed`. A row is done once *Closed* has a date.
* **Gate Log**: `Part | Gate | Due | Passed`. The app writes this when a gate is passed. A blank *Passed* means the gate was
  passed before the app recorded it.
* **Gates**: `Gate | After | Due before`. *After* lists (comma-separated) the gates that come first; blank = can start any time.
  *Due before* names a key date. Edit this with **Edit gates** in the app rather than by hand.

The master's **Key Dates** sheet is `Date | Event | Type | People | Done | Internal deadline | End date | Track | Competition | Notes`.

The **Plan** sheet looks like this:

| Part | Parent | Owner | Assignees | Tags | Current Gate | Flag | Notes | DR01 - Concept | DR02 - Detailed Design | DR03 - Final Design | Manufacturing | Testing |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Front Wing | | Sam | Yi, Alex | Car | DR02 - Detailed Design | OK | | 2026-10-11 | 2026-10-28 | 2026-11-10 | 2026-11-30 | 2026-12-20 |
| Element 1 | Front Wing | Alex | | Composites | DR01 - Concept | OK | | 2026-10-04 | 2026-10-25 | N/A | 2026-11-28 | 2026-12-18 |

* *Parent* is blank for a top-level part; for a sub-part it's the name of the part it belongs to. *Assignees* and *Tags* are comma-separated.
* *Current Gate* lists the gate(s) being worked on, comma-separated, or *Done*. The app writes it; if you change it by hand,
  the app counts every gate before the ones you list as passed.
* The gate columns hold due dates, or **N/A** for a gate the part skips. Excel highlights the gates being worked on and turns
  them red when overdue, and greys out skipped gates and finished parts.
* The **Actions** and **Gate Log** sheets refer to a sub-part by its full name, e.g. `Front Wing / Element 1`.
* Files made by older versions (fewer columns, no Gates sheet) still open; they're upgraded the next time the app saves them,
  and their gates run one after another until a lead changes that in **Edit gates**.

---

## 7. Customising

* **Change a subteam's gates:** the lead clicks **Edit gates** on the subteam page (see section 4).
  Files made with the original gates (*Concept*, *Detailed Design*, *Drawings*) show the DR01–DR03 names, and are
  updated to them the next time that subteam saves from the app.
* **Add a subteam:** copy any subteam file in `Subteams/`, rename it (e.g. `Cooling.xlsx`) and clear its rows. Then add its emoji to `SUBTEAM_EMOJI` in `tracker_app.py`; it shows 🔧 otherwise.
* **Remove a subteam:** delete or move its file out of `Subteams/`.
* **Change the "due soon" window:** `DUE_SOON_DAYS` in `tracker_core.py`.
* **Change the overview look-ahead:** `LOOK_AHEAD_DAYS` in `tracker_app.py`.

---

## 8. Troubleshooting

| Problem | Fix |
|---|---|
| Double-clicking the `.pyzw` opens an editor or asks what to open it with | Right-click → *Open with* → *Choose another app* → **Python** (the one with the rocket icon, `pyw.exe`) → *Always*. Or reinstall Python with "py launcher" ticked. |
| Installing the packages fails | You need internet access the first time. In a terminal, run `py -m pip install flet openpyxl matplotlib`, then double-click again. |
| Nothing happens on double-click | In a terminal, run `py "UCR Project Tracker.pyzw"` to see the error. |
| "… was changed by someone else" | Someone saved it at the same time. The app reloaded the file, so make your edit again. |
| "Edit failed." | You don't have Edit access to that file on SharePoint (it now shows as View only). Ask the PM. |
| A subteam is missing | Check its file is in `Subteams/` and has a sheet called **Plan**. |
| The app can't find the tracker | Add the SharePoint shortcut (section 3), wait for it to appear in File Explorer, then click the folder button in the app and **Browse** to it. |

---

## 9. Files in this repo

| File | What it is |
|---|---|
| `Project Tracker.pyw` | Double-click launcher when running from the source code (no terminal) |
| `launcher.py` | Shared start-up: offers to install missing packages, then runs the app, `gantt` or `setup` |
| `build_dist.py` | Builds the single file `dist/UCR Project Tracker.pyzw` |
| `tracker_app.py` | The app (Flet UI) |
| `tracker_core.py` | Reads and writes the Excel files, status rules, save safety. No UI code. |
| `setup_tracker.py` | Creates an empty tracker folder |
| `export_gantt.py` | Exports the season as a Pit Wall-style Gantt PDF |
| `requirements.txt` | Python dependencies |
| `assets/` | App icon (`icon.svg` source, `icon.ico` with a native render per size) and the full UCR logo (`logo.svg`, `logo.png`) used in the Gantt header |

The tracker data (`Project Tracker/`) is **not** in this repo. It lives on SharePoint and contains members' names.

### Releasing a new version

1. Run `py build_dist.py`. It writes `dist/UCR Project Tracker.pyzw`; the `dist/` folder isn't committed.
2. On GitHub: **Releases → Draft a new release**. Create a tag such as `v1.1` and drag the `.pyzw` file into the release.
3. Publish it, then tell members to download the new file and replace their old one. Their settings are kept.
