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

**2. Get the app.** Either:
   * `git clone <this repo's URL>`; or
   * on the repo page, **Code → Download ZIP**, then unzip it somewhere permanent (e.g. `Documents\UCR-Project-Tracker`).

**3. Install the dependencies.** Open the folder in File Explorer, click the address bar, type `cmd` and press Enter. Then run:

   ```
   py -m pip install -r requirements.txt
   ```

   This installs [Flet](https://flet.dev) (the UI) and [openpyxl](https://openpyxl.readthedocs.io) (reads and writes the Excel files).

**4. Start it** by double-clicking **`Project Tracker.pyw`**. The window opens in the middle of the screen with no terminal behind it.
   The first launch can take 10–20 seconds while Flet unpacks.

> **Mac / Linux:** run `python3 -m pip install -r requirements.txt` and then `python3 tracker_app.py`.
> The `.pyw` double-click is Windows-only.

> **Tip:** right-click `Project Tracker.pyw` → *Send to → Desktop (create shortcut)* for a desktop icon.
> In the shortcut's *Properties → Change Icon* you can point it at `assets\icon.ico`.

---

## 2. One-time setup (PM / captain)

**1. Create the tracker files.** From the app folder, run:

   ```
   py setup_tracker.py --out "Project Tracker"
   ```

   This creates:

   ```
   Project Tracker/
       Master.xlsx              key dates + the All Parts rollup
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
  * one card per subteam (progress, plus how many parts are late, flagged or due soon);
  * **Needs attention**: every late, *At risk* or *Blocked* part;
  * **Next 14 days**: gates and key dates coming up;
  * the **next design review**, in the subtitle.
* **Subteam tabs** (left rail) list that subteam's parts, sorted with problems first. The dots show gate
  progress: filled = passed, ringed = current. Click a part to see its details.
* **Key dates** lists competition deadlines, design reviews (gate meetings) and milestones. Turn on *Show past* to see older ones.
* **↻** reloads the files (do this after a lead says they've updated). **☾** switches between light and dark mode.

### Subteam leads (your tab shows "You can edit")

* **Add part** (bottom-right button): give it a name and owner, then fill in a **due date for each gate**.
* **Click a part** to edit it:
  * **Gate passed** moves the part to its next gate (after the last gate it becomes *Done*);
  * **Flag** marks the part as *At risk* or *Blocked*. It then appears under *Needs attention*, so raise it at the next meeting;
  * **Notes** is a one-line reason, e.g. "waiting on supplier".
* **Weekly habit:** before the leads meeting, update each part's current gate, dates and flag. It takes about 5 minutes.

### PM / captain

* **Key dates:** add, edit or delete entries. Use the *Design Review* type for gate meetings (e.g. "DR2: Aero + Suspension").
* **`Master.xlsx > All Parts`** is rebuilt automatically whenever someone with edit access to the master saves anything.
  It lets people without the app see everything in the browser. Don't edit that sheet by hand.
* **In the meeting**, go through *Needs attention* first, then *Next 14 days*.

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

---

## 6. Editing in Excel instead

Leads can also edit their file directly in Excel (desktop or browser). Each file has one sheet, **Plan**:

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
| Double-clicking the `.pyw` opens an editor | Right-click → *Open with* → *Choose another app* → **Python** (the one with the rocket icon, `pyw.exe`) → *Always*. Or reinstall Python with "py launcher" ticked. |
| A message says a package is missing | Run `py -m pip install -r requirements.txt` in the app folder. |
| Nothing happens on double-click | Run `py tracker_app.py` from a terminal in the app folder to see the error. |
| "… is open in Excel" | Close the file in Excel and save again. |
| "… was changed by someone else" | Someone saved it at the same time. The app reloaded the file, so make your edit again. |
| "… is view-only for you" | You don't have Edit access on SharePoint. Ask the PM. |
| A subteam is missing | Check its file is in `Subteams/` and has a sheet called **Plan**. |
| The app can't find the tracker | Make sure the SharePoint folder is synced (step 3), then browse to it. |

---

## 9. Files in this repo

| File | What it is |
|---|---|
| `Project Tracker.pyw` | Double-click launcher (no terminal) |
| `tracker_app.py` | The app (Flet UI) |
| `tracker_core.py` | Reads and writes the Excel files, status rules, save safety. No UI code. |
| `setup_tracker.py` | Creates an empty tracker folder |
| `requirements.txt` | Python dependencies |
| `assets/` | App icon (`icon.svg` source, `icon.ico` with a native render per size) |

The tracker data (`Project Tracker/`) is **not** in this repo. It lives on SharePoint and contains members' names.
