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
   py "UCR Project Tracker.pyzw" setup --out "UCR-04"
   ```

   (From the source code, `py setup_tracker.py --out "UCR-04"` does the same.)

   Each car gets its own tracker folder, named after the car. This creates:

   ```
   UCR-04/
       Master.xlsx              key dates, the general gate library and templates, the All Parts / All Actions rollups
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
   * `--gates "DR01 - Concept" "DR02 - Detailed Design" ...` picks the gates of the general *Standard* template (in order).
   * `--car UCR-04` sets the car name (default: the folder name).

   Later cars don't need this script: use **New car…** in the app (section 4).

   The script refuses to overwrite an existing tracker.

**2. Upload the tracker to SharePoint**, into **UCalgary-Racing → Documents → Project Management**, as its own folder
   (`Project Management/UCR-04/`). Next year's `UCR-05/` goes next to it; the app's car switcher lists every car folder there.

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

The app remembers your choice. To switch later, open **Settings** (⚙ at the bottom of the left rail) and click **Tracker
folder**; the current folder is shown under it.

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
* **Subteam tabs** (left rail) have three views (switch at the top right; your choice is remembered):
  * **List**: the parts, sorted with problems first. The dots show each gate of the part's template: filled = done,
    amber = in progress, ringed = ready to start (several at once when gates run in parallel), hollow grey = off,
    faint = waiting. ⚠ marks a part with a conflict or a date that breaks a gate rule (hover to see which). Under the name
    are its **tags**; under the owner, anyone else working on it. Filter by text or click a **tag** chip;
  * **Grid**: one grid per gate template, a row per assembly and sub-assembly, a box per gate (✓ done, ◐ in progress,
    – off, empty = not started). **Right-click (or click) a box** to set it. A thick red border means late or out of order
    (a gate done before the one it comes after: a *conflict*);
  * **Design reviews**: a row per assembly with **DR1 Concept / DR2 Intermediate / DR3 Final**. Click a cell to set the
    date, outcome and presenter. Upcoming reviews show in the owners' *My work*, the overview and the calendar;
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
* **Gates** (under ⚙ Settings, or *Gate templates* on a subteam page) lists every **gate template** (the workflows parts move through) and the **gate library** they're
  built from: each gate's code, name, designation (*General* or a sub-team), description, and **Used in** (how many
  templates and items use it). Click a template or gate to see or edit it.
* **The car** (e.g. *UCR-04*) shows under the logo. Click it to switch to another car's tracker, or start the next one.
* **↻** reloads the files (do this after a lead says they've updated). **⚙ Settings** (bottom of
  the rail) holds the **Night mode** switch, the **Tracker folder** and **Gates**.

### Subteam leads (your tab shows "You can edit")

* **Gate templates** (*Gates* page, or the button on your subteam page): each part follows one template. *General*
  templates (e.g. *Mechanical Standard*) are kept in the master and shared by every sub-team; your sub-team's own (e.g.
  *Aero Car Program*, *Aero Development*) are kept in your file. **New template** builds one, and clicking one edits it:
  * add gates by name or pick them from the **library** (☰); new gate names join the library with a short code;
  * **Comes after** (chips under each gate): the gates that must be done or off before this one starts. Gates with the same
    prerequisites **run in parallel**, e.g. Mould Design only needs CFD Freeze, while Mounting Design needs CFD Freeze
    *and* Mounting Concept. A gate can only come after gates above it. Doing a gate before the ones it comes after is
    allowed but shows as a **conflict**;
  * **Due before**: a key date (e.g. *Chassis Freeze*) the gate must be due before. Parts dated later get a ⚠ warning, as do
    gates due inside a *Blackout* or before a gate they come after;
  * renaming a gate keeps its dates and history; removing one deletes them. A template that parts still follow can't be deleted.
* **Gate library** (*Gates* page): **Add gate** with a code, name and description. A gate that's in use can't be renamed here
  (rename it in its templates) or deleted.
* **Add part** (bottom-right button): give it a name, an **owner** (the one accountable person), anyone else **working on it**,
  any **tags** (e.g. *Development*, *Aero-Testing*; click a suggested tag to add it), its **gate template** (your sub-team's
  default unless you pick another), then a **due date for each gate**.
* **Sub-parts:** to keep big assemblies tidy, open a part (e.g. Front Wing) and click **Add sub-part**, or set *Sub-part of*
  when editing. Parts with sub-parts show collapsed with a summary ("8 sub-parts · 1 need attention"); click the arrow to
  expand. An assembly sorts by its worst sub-part, so problems still float to the top. Sub-parts go one level deep.
  An assembly counts as **Done** once all its sub-assemblies are done.
* **Click a part** to edit it:
  * each **gate** has its own row with its due date and a state: **Not started**, **In progress**, **Done** (records today;
    edit the *Done on* date if it was earlier) or **Off** (doesn't apply to this part, e.g. no mould for a bracket). The app
    shows whether each gate was done early or late. The grid's right-click menu sets the same states;
  * **Flag** marks the part as *At risk* or *Blocked*. It then appears under *Needs attention*, so raise it at the next meeting;
  * **Notes** is a one-line reason, e.g. "waiting on supplier".
* **Actions** (follow-ups from design reviews):
  * *Add action* on your subteam page. Give each action one owner, a due date, the part it relates to, and where it came from (e.g. "DR2");
  * tick the box when it's done. Overdue actions show under *Needs attention*, and the owner sees them under *My work*.
* **Weekly habit:** before the leads meeting, mark finished gates done (the grid is quickest), update dates and flags, and
  tick off finished actions. It takes about 5 minutes.

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
* **General gate library and templates** (*Gates* page) are yours to edit; sub-teams use them alongside their own.
* **New sub-team** (Overview, top right): a name and a starting template. It creates `Subteams/<name>.xlsx`; on SharePoint,
  give its lead Edit access to that file.
* **New car…** (click the car name under the logo): starts the next car (e.g. *UCR-05*) as a new folder next to this one.
  Choose what carries over: each sub-team's **assembly structure** (parts, sub-parts, templates, tags), the **gate library and
  templates**, and the **key dates** (optionally moved a year later). Owners, dates and progress always start fresh. Set the
  new folder's SharePoint permissions as in section 2.
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

Each gate of a part's template is **done**, **off**, **in progress**, **ready** (everything it comes after is done or off) or
**waiting**. The part's *current gate* is the gate(s) in progress or ready, and its status uses the earliest due date among them.
An assembly is also done once all its sub-assemblies are done. A gate marked done or in progress while one it comes after
isn't done is a **conflict**: it shows in red in the grid and puts the part under *Needs attention*.

| Status | Meaning |
|---|---|
| **Late** | today is past the due date of a gate in progress or ready |
| **Due soon** | one of them is due within 7 days |
| **On track** | they're due later than that |
| **No date** | they have no due dates yet |
| **Done** | every gate is done or off (or, for an assembly, every sub-assembly is done) |

**Flags** are separate from status and set by the lead:
* **OK**
* **At risk**: might slip; discuss it.
* **Blocked**: can't progress without help.

**Actions and key dates** follow the same Late / Due soon / On track rules, based on their own due date.
* An action is *Done* once it's ticked.
* A deadline, milestone or freeze is *Done* once the PM ticks it; until then it's measured against the internal deadline
  if there is one.
* A design review, blackout or event shows as *Now* while it's on and *Held* once it's over.
* A DR1–DR3 review in the *Design reviews* view counts down to its date, then shows **Needs outcome** until one is entered.

**On time** (subteam subtitle) counts the gates passed on or before their due date, out of all gates passed that had a due date.

---

## 6. Editing in Excel instead

Leads can also edit their file directly in Excel (desktop or browser). Each subteam file has these sheets:

* **Plan**: the parts (below). This is the one you'd normally edit.
* **Actions**: `Action | Part | Owner | Due | From | Closed`. A row is done once *Closed* has a date.
* **Gate Log**: `Part | Gate | Due | Passed | Status`. The app writes the gates that are done or in progress. A blank *Passed*
  means the gate was done before the app recorded it.
* **Design Reviews**: `Part | Review | Date | Outcome | Presenter | Notes`, one row per assembly and review.
* **Gate Templates**: `Template | Designation | Gate | After | Due before`, one row per gate, in order. *After* lists
  (comma-separated) gates earlier in the template that come first; *Due before* names a key date.
* **Gate Library**: `Code | Gate | Designation | Description`, this sub-team's own gates.
* **Settings**: the sub-team's default template.

Edit templates and the library in the app rather than by hand. The master has the same *Gate Templates* and *Gate Library*
sheets for the general ones, a *Settings* sheet with the car name, and **Key Dates**:
`Date | Event | Type | People | Done | Internal deadline | End date | Track | Competition | Notes`.

The **Plan** sheet looks like this:

| Part | Parent | Owner | Assignees | Tags | Template | Current Gate | Flag | Notes | DR01 - Concept | DR02 - Detailed Design | DR03 - Final Design | Manufacturing | Testing |
|---|---|---|---|---|---|---|---|---|---|---|---|---|---|
| Front Wing | | Sam | Yi, Alex | Car | | DR02 - Detailed Design | OK | | 2026-10-11 | 2026-10-28 | 2026-11-10 | 2026-11-30 | 2026-12-20 |
| Element 1 | Front Wing | Alex | | Composites | | DR01 - Concept | OK | | 2026-10-04 | 2026-10-25 | N/A | 2026-11-28 | 2026-12-18 |

* *Parent* is blank for a top-level part; for a sub-part it's the name of the part it belongs to. *Assignees* and *Tags* are comma-separated.
* *Template* is the gate template the part follows; blank means the sub-team's default. The gate columns cover every gate of
  every template in use; a part only uses its own template's columns.
* *Current Gate* lists the gate(s) being worked on, comma-separated, or *Done*. The app writes it; if you change it by hand,
  the app counts every gate before the ones you list as passed.
* The gate columns hold due dates, or **N/A** for a gate that's off for the part. Excel highlights the current gates and
  turns them red when overdue, and greys out gates that are off and finished parts.
* The **Actions** and **Gate Log** sheets refer to a sub-part by its full name, e.g. `Front Wing / Element 1`.
* Files made by older versions still open, with nothing lost; they're upgraded the next time the app saves them. A file
  with its own gate list (or an older *Gates* sheet) becomes that sub-team's own template.

---

## 7. Customising

* **Change a subteam's gates:** edit its template on the *Gates* page (see section 4).
  Files made with the original gates (*Concept*, *Detailed Design*, *Drawings*) show the DR01–DR03 names, and are
  updated to them the next time that subteam saves from the app.
* **Add a subteam:** **New sub-team** on the Overview. To give it an emoji, add it to `SUBTEAM_EMOJI` in `tracker_app.py`;
  it shows 🔧 otherwise.
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
