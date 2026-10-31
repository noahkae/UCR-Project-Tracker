"""Create an empty 'Project Tracker' folder: Master.xlsx plus one plan file per subteam.

    python setup_tracker.py --out "Project Tracker"
    python setup_tracker.py --out "Project Tracker" --subteams Aero Suspension Electrical

Upload the folder to a SharePoint document library, then give each subteam lead
Edit access to their own file, the PM/captain Edit access to Master.xlsx, and
everyone else View access.
"""

import argparse
from pathlib import Path

import tracker_core as core

SUBTEAMS = ["Aero", "Suspension", "Chassis & Controls", "Powertrain & Battery Pack", "Electrical", "Driverless"]


def main():
    ap = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    ap.add_argument("--out", required=True, help="folder to create")
    ap.add_argument("--subteams", nargs="+", default=SUBTEAMS)
    ap.add_argument("--gates", nargs="+", default=core.DEFAULT_GATES, help="gate names, in order")
    args = ap.parse_args()

    out = Path(args.out)
    if (out / core.MASTER).exists():
        raise SystemExit(f"{out / core.MASTER} already exists; nothing was changed.")
    (out / core.SUBTEAM_DIR).mkdir(parents=True, exist_ok=True)
    core.create_master(out / core.MASTER)
    for name in args.subteams:
        path = out / core.SUBTEAM_DIR / f"{name}.xlsx"
        if not path.exists():
            core.create_subteam_file(path, args.gates)
    print(f"Created {out.resolve()} with {len(args.subteams)} subteam files. Gates: {' > '.join(args.gates)}")


if __name__ == "__main__":
    main()
