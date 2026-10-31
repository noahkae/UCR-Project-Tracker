"""Double-click to open the Project Tracker (.pyw runs without a terminal window)."""

import ctypes
import traceback

try:
    import flet as ft

    import tracker_app
except ImportError as e:
    ctypes.windll.user32.MessageBoxW(
        None,
        f"The tracker needs a couple of Python packages first ({e.name} is missing).\n\n"
        "Open a terminal in this folder and run:\n    pip install -r requirements.txt",
        "Project Tracker", 0x30,
    )
    raise SystemExit(1)

try:
    ft.run(tracker_app.main, assets_dir=str(tracker_app.ASSETS))
except Exception:
    # No terminal to print to, so show the error instead of failing silently.
    ctypes.windll.user32.MessageBoxW(None, traceback.format_exc()[-1500:], "Project Tracker crashed", 0x10)
