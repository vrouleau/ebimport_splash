#!/usr/bin/env python3
"""Export template structure from a SPLASH .mdb to JSON for the Lenex path.

Usage:
    python export_template_json.py --mdb template.mdb --out template_struct.json
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from load_to_mdb import MDB, TemplateIndex


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--mdb", required=True, type=Path)
    ap.add_argument("--out", type=Path, default=Path("template_struct.json"))
    args = ap.parse_args()

    db = MDB(args.mdb, dry_run=True)
    template = TemplateIndex(db)

    export = {"styles": {}, "events": []}

    for uid, style in template.styles_by_uid.items():
        if uid < 500:
            continue
        export["styles"][str(uid)] = {
            "distance": style.distance,
            "relay_count": style.relay_count,
            "name": style.name,
        }

    for key, ev_list in template.events_by_uid_gender.items():
        for ev in ev_list:
            if ev.uniqueid < 500:
                continue
            export["events"].append({
                "uid": ev.uniqueid,
                "gender": ev.gender,
                "eid": ev.swim_event_id,
                "enum": ev.event_number,
                "round": ev.round,
                "masters": ev.masters,
                "session": ev.session_id,
                "agegroups": [
                    {"id": a.agegroup_id, "min": a.amin, "max": a.amax, "g": a.gender}
                    for a in ev.agegroups
                ],
            })

    with open(args.out, "w") as f:
        json.dump(export, f, indent=2)

    print(f"Exported {len(export['styles'])} styles, {len(export['events'])} events -> {args.out}")
    db.close()


if __name__ == "__main__":
    main()
