"""Start the tracker: make sure its Python packages are installed (offering to install them), then run it.

Used by both `Project Tracker.pyw` and the single-file `UCR Project Tracker.pyzw`:

    UCR Project Tracker.pyzw                  double-click: open the app
    py "UCR Project Tracker.pyzw" gantt ...   Gantt export (same options as export_gantt.py)
    py "UCR Project Tracker.pyzw" setup ...   create a tracker folder (same options as setup_tracker.py)

Python itself is never bundled; the packages are installed with pip the first time.
"""

import importlib.util
import os
import subprocess
import sys
import tempfile
import threading
import traceback
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ARCHIVE = HERE if zipfile.is_zipfile(HERE) else None  # set when running from the .pyzw
TITLE = "UCR Project Tracker"
NO_WINDOW = getattr(subprocess, "CREATE_NO_WINDOW", 0)


def read_file(name):
    """A file shipped with the app, whether it's a folder or a single .pyzw."""
    if ARCHIVE:
        with zipfile.ZipFile(ARCHIVE) as z:
            return z.read(name)
    return (HERE / name).read_bytes()


def assets_dir():
    """Real files for the icon and logo (Windows and matplotlib can't read them from inside a zip)."""
    if not ARCHIVE:
        return HERE / "assets"
    with zipfile.ZipFile(ARCHIVE) as z:
        entries = [i for i in z.infolist() if i.filename.startswith("assets/") and not i.is_dir()]
        version = format(sum(i.CRC for i in entries) & 0xFFFFFFFF, "08x")
        target = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "FSAETracker" / f"assets-{version}"
        for info in entries:
            path = target / Path(info.filename).name
            if not path.exists() or path.stat().st_size != info.file_size:
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_bytes(z.read(info))
    return target


def message(text, title=TITLE, icon=0x40, buttons=0):
    """A Windows message box (we usually have no console). Returns the button pressed."""
    if sys.platform == "win32":
        return __import__("ctypes").windll.user32.MessageBoxW(None, text, title, icon | buttons)
    print(text)
    return 6


def ensure_packages(modules):
    """Offer to pip-install anything missing, then restart so the new packages are picked up."""
    missing = [m for m in modules if importlib.util.find_spec(m) is None]
    if not missing:
        return
    text = (f"The tracker needs some Python packages that aren't installed yet ({', '.join(missing)}).\n\n"
            "Install them now? It takes a minute or two and only happens once.")
    if interactive():
        ok = input(f"{text} [Y/n] ").strip().lower() in ("", "y", "yes")
    else:
        ok = message(text, icon=0x20, buttons=0x4) == 6  # question icon, Yes/No; 6 = Yes
    if not ok:
        sys.exit(1)
    reqs = [line.split("#")[0].strip() for line in read_file("requirements.txt").decode().splitlines()]
    with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False) as f:
        f.write("\n".join(r for r in reqs if r))
    python = Path(sys.executable).with_name("python.exe") if sys.platform == "win32" else Path(sys.executable)
    cmd = [str(python if python.exists() else sys.executable), "-m", "pip", "install", "--disable-pip-version-check", "-r", f.name]
    result = run_with_progress(cmd, "Installing the tracker's Python packages…\nThis takes a minute or two and only happens once.")
    os.unlink(f.name)
    if result is None or result.returncode != 0:
        detail = (result.stdout + result.stderr)[-1200:] if result else "pip could not be started."
        message(f"Installing the packages failed:\n\n{detail}\n\nTry running this in a terminal:\n    py -m pip install flet openpyxl matplotlib",
                icon=0x10)
        sys.exit(1)
    restart()


def restart():
    """Start again in a new process: a brand-new user site-packages folder isn't on sys.path until Python restarts."""
    subprocess.Popen([sys.executable, *sys.argv])
    sys.exit(0)


def run_with_progress(cmd, text):
    """Run a command with a small 'please wait' window (tkinter is part of standard Python)."""
    if interactive():
        print(text)
        return subprocess.run(cmd, capture_output=True, text=True)
    result = {}
    worker = threading.Thread(target=lambda: result.update(p=subprocess.run(cmd, capture_output=True, text=True, creationflags=NO_WINDOW)),
                              daemon=True)
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        worker.start()
        worker.join()
        return result.get("p")
    root = tk.Tk()
    root.title(TITLE)
    root.resizable(False, False)
    try:
        root.iconbitmap(str(assets_dir() / "icon.ico"))
    except Exception:  # noqa: BLE001 - the icon is cosmetic
        pass
    tk.Label(root, text=text, padx=24, pady=16, justify="left").pack()
    bar = ttk.Progressbar(root, mode="indeterminate", length=320)
    bar.pack(padx=24, pady=(0, 20))
    bar.start(12)
    root.eval("tk::PlaceWindow . center")
    worker.start()

    def poll():
        if worker.is_alive():
            root.after(200, poll)
        else:
            root.destroy()
    root.after(200, poll)
    root.mainloop()
    worker.join()
    return result.get("p")


def interactive():
    return sys.stdin is not None and sys.stdin.isatty()


def main():
    command, args = (sys.argv[1], sys.argv[2:]) if len(sys.argv) > 1 and sys.argv[1] in ("gantt", "setup") else (None, sys.argv[1:])
    if command == "gantt":
        ensure_packages(["openpyxl", "matplotlib"])
        import export_gantt
        export_gantt.LOGO = assets_dir() / "logo.png"
        sys.argv = [f"{sys.argv[0]} gantt", *args]
        export_gantt.main()
    elif command == "setup":
        ensure_packages(["openpyxl"])
        import setup_tracker
        sys.argv = [f"{sys.argv[0]} setup", *args]
        setup_tracker.main()
    else:
        ensure_packages(["flet", "openpyxl"])
        import flet as ft
        import tracker_app
        tracker_app.ASSETS = assets_dir()
        try:
            ft.run(tracker_app.main, assets_dir=str(tracker_app.ASSETS))
        except Exception:  # noqa: BLE001 - no terminal to print to, so show the error instead of failing silently
            message(traceback.format_exc()[-1500:], "Project Tracker crashed", icon=0x10)


if __name__ == "__main__":
    main()
