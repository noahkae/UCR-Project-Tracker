"""Build the single file to share: dist/UCR Project Tracker.pyzw

    py build_dist.py

A .pyzw is Python's own single-file app format (a zip with a __main__.py). On Windows,
double-clicking it runs it with pythonw, so no terminal opens. Python is NOT bundled:
people install Python once, and the first launch offers to install the packages.
"""

import shutil
import tempfile
import zipapp
from pathlib import Path

ROOT = Path(__file__).resolve().parent
FILES = [
    "launcher.py", "tracker_app.py", "tracker_core.py", "export_gantt.py", "setup_tracker.py",
    "requirements.txt", "assets/icon.ico", "assets/logo.png",
    *[f"assets/{p.name}" for p in sorted((Path(__file__).resolve().parent / "assets").glob("team-*.svg"))],
]
OUT = ROOT / "dist" / "UCR Project Tracker.pyzw"


def main():
    with tempfile.TemporaryDirectory() as tmp:
        stage = Path(tmp) / "app"
        for name in FILES:
            (stage / name).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(ROOT / name, stage / name)
        (stage / "__main__.py").write_text("import launcher\n\nlauncher.main()\n")
        OUT.parent.mkdir(exist_ok=True)
        zipapp.create_archive(stage, OUT, compressed=True)
    print(f"Built {OUT} ({OUT.stat().st_size / 1024:.0f} KB)")


if __name__ == "__main__":
    main()
