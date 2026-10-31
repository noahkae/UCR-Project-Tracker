# UCR Project Tracker

A small project-management tool for UCalgary Racing. Each subteam's plan is **one Excel file on SharePoint
that only its lead can edit**, and a **desktop app** (Material Design, built with Python + Flet) shows every
subteam's plan, deadlines and design reviews in one place.

* **SharePoint handles permissions.** Leads edit their own subteam file, the PM edits the master, and everyone else can view.
* **The data is just Excel.** If the app ever breaks, the plan is still readable and editable in Excel.
* **There's no server, database, login screen or `.exe`.**

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
   * `--gates Concept "Detailed Design" ...` picks your own gate list (in order).

   The script refuses to overwrite an existing tracker.

**2. Upload `Project Tracker/` to SharePoint**, e.g. the team's Teams channel → *Files*, or a document library.

**3. Set permissions** with *… → Manage access* on each file (or folder):

   | Who | Access |
   |---|---|
   | Each subteam lead | **Edit** on their own `Subteams/<name>.xlsx` |
   | PM / captain | **Edit** on `Master.xlsx` (and the subteam files if they help plan) |
   | Everyone else | **View** on the whole library |

**4. Add the first key dates** (competition deadlines, design reviews) in the app's *Key dates* tab, or directly in `Master.xlsx`.

---

## 3. Connect the app to the tracker

1. In SharePoint, open the *Project Tracker* folder and click **Sync** (or **Add shortcut to My files**).
   OneDrive then keeps a local copy in File Explorer, usually under `OneDrive - University of Calgary`.
2. Start the app. It searches your OneDrive folders for `Master.xlsx`. If it can't find one, it asks you to
   **Browse** to the synced *Project Tracker* folder.
3. The app remembers the folder. To change it later, delete `%APPDATA%\FSAETracker\config.json`.

Files you can only view are synced **read-only**, so the app shows them as **View only** and you can't change them by accident.

---

## 4. Using the app

### Everyone

* **Overview** shows:
  * one card per subteam (progress, open actions, plus how many parts are late, flagged or due soon);
  * **Needs attention**: late, *At risk* or *Blocked* parts; overdue actions; and key dates that are overdue or due within 7 days;
  * **Next 14 days**: gates, actions and key dates coming up;
  * the **next design review**, in the subtitle.
* **My work**: pick your name once (it's remembered) to see **your parts, your open actions and the key dates you're
  responsible for**, with anything late at the top. Key dates you're named on carry a **You** tag everywhere in the app.
* **Subteam tabs** (left rail) show:
  * the subteam's parts, sorted with problems first. The dots show gate progress: filled = passed, ringed = current.
    Use the filter box to search by part, owner or note. Click a part to see its details and **gate history**;
  * its **actions**, open ones first;
  * how many gates it has passed **on time**, in the subtitle.
* **Calendar** is a month view of every gate due date, action and key date.
  * Use the chips at the top to show or hide each subteam, key dates and actions. Your choice is remembered.
  * Colours: red = late, amber = due soon, green = on track, grey = done, blue = a gate still to come, purple = an upcoming design review.
  * Click an item to open it, or click a busy day to see everything on it. Use ‹ › to change month and *Today* to jump back.
* **Key dates** lists competition deadlines, design reviews (gate meetings) and milestones, with the people responsible for each.
  Overdue ones stay visible until marked done. Turn on *Show past* to see the rest.
* **↻** reloads the files (do this after a lead says they've updated). **☾** switches between light and dark mode.

### Subteam leads (your tab shows "You can edit")

* **Add part** (bottom-right button): give it a name and owner, then fill in a **due date for each gate**.
* **Click a part** to edit it:
  * **Gate passed** moves the part to its next gate (after the last gate it becomes *Done*). When you save, the app
    records the date the gate was passed in the part's **gate history** and works out whether it was early or late.
    Moving a part back un-records those gates;
  * **Flag** marks the part as *At risk* or *Blocked*. It then appears under *Needs attention*, so raise it at the next meeting;
  * **Notes** is a one-line reason, e.g. "waiting on supplier".
* **Actions** (follow-ups from design reviews):
  * *Add action* on your subteam page. Give each action one owner, a due date, the part it relates to, and where it came from (e.g. "DR2");
  * tick the box when it's done. Overdue actions show under *Needs attention*, and the owner sees them under *My work*.
* **Weekly habit:** before the leads meeting, update each part's current gate, dates and flag, and tick off finished actions. It takes about 5 minutes.

### PM / captain

* **Key dates:** add, edit or delete entries.
  * Use the *Design Review* type for gate meetings (e.g. "DR2: Aero + Suspension").
  * Fill in **People responsible** (comma-separated names, spelt the same way as in the subteam files, e.g. "Sam Lee, Alex Kim").
    They'll see the date in *My work*, and it shows under *Needs attention* from 7 days before.
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
* **Bars:** each gate is a bar from the previous gate's due date to its own due date. White = passed, red = current gate, outline = still to come.
* **◆ diamonds:** when a gate was *actually* passed, from the gate history. White = on time or early; red = late, with a line back to the due date. This is the "in hindsight" part.
* **Dashed lines:** key dates, with design reviews in red. The **Key dates** strip underneath names them.
* **Solid red line:** today.

---

## 5. How status is worked out

Each part has a **current gate**, the gate it's working towards. Every gate to its left counts as passed.

| Status | Meaning |
|---|---|
| **Late** | today is past the current gate's due date |
| **Due soon** | the current gate is due within 7 days |
| **On track** | the current gate is due later than that |
| **No date** | the current gate has no due date yet |
| **Done** | the current gate is set to *Done* |

**Flags** are separate from status and set by the lead:
* **OK**
* **At risk**: might slip; discuss it.
* **Blocked**: can't progress without help.

**Actions and key dates** follow the same Late / Due soon / On track rules, based on their own due date.
* An action is *Done* once it's ticked.
* A deadline or milestone is *Done* once the PM ticks it.
* A design review that has happened shows as *Held*.

**On time** (subteam subtitle) counts the gates passed on or before their due date, out of all gates passed that had a due date.

---

## 6. Editing in Excel instead

Leads can also edit their file directly in Excel (desktop or browser). Each subteam file has three sheets:

* **Plan**: the parts (below). This is the one you'd normally edit.
* **Actions**: `Action | Part | Owner | Due | From | Closed`. A row is done once *Closed* has a date.
* **Gate Log**: `Part | Gate | Due | Passed`. The app fills this in when a gate is passed, so there's no need to edit it.
  Changing *Current Gate* directly in Excel does **not** add to the log; use the app for that.

The master's **Key Dates** sheet is `Date | Event | Type | People | Done`.

The **Plan** sheet looks like this:

| Part | Owner | Current Gate | Flag | Notes | Concept | Detailed Design | Drawings | Manufacturing | Testing |
|---|---|---|---|---|---|---|---|---|---|
| Front Wing | Sam | Detailed Design | OK | | 2026-10-11 | 2026-10-28 | 2026-11-10 | 2026-11-30 | 2026-12-20 |

* *Current Gate* and *Flag* are dropdowns.
* The gate columns hold due dates. Excel greys out passed gates, highlights the current one, and turns it red when overdue.
* **Close the file in Excel before saving from the app.** The app tells you if it's still open.

---

## 7. Customising

* **Rename or add gates for a subteam:** edit the header cells after *Notes* in that subteam's file. The app reads the gates from the headers, in order.
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
| "… is open in Excel" | Close the file in Excel and save again. |
| "… was changed by someone else" | Someone saved it at the same time. The app reloaded the file, so make your edit again. |
| "… is view-only for you" | You don't have Edit access on SharePoint. Ask the PM. |
| A subteam is missing | Check its file is in `Subteams/` and has a sheet called **Plan**. |
| The app can't find the tracker | Make sure the SharePoint folder is synced (step 3), then browse to it. |

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
