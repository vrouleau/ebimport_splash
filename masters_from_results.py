"""
Generate a Masters-final Lenex from a SPLASH results export.

Reads the exported .lxf, finds Masters-age athletes with results in
prelim events, and generates a new Lenex file that registers them
in the corresponding Masters final events (with their SWIMTIME as
entry time). Import this into SPLASH to populate the Masters finals.

Usage:
    python masters_from_results.py --lxf export_result.lxf --out masters_import.lxf
"""
from __future__ import annotations

import argparse
import datetime as dt
import sys
import zipfile
from pathlib import Path
from xml.dom import minidom
from xml.etree import ElementTree as ET


def age_at(birthdate_str: str, ref: dt.date) -> int | None:
    if not birthdate_str:
        return None
    try:
        bd = dt.date.fromisoformat(birthdate_str)
    except ValueError:
        return None
    years = ref.year - bd.year
    if (ref.month, ref.day) < (bd.month, bd.day):
        years -= 1
    return years


def find_bracket(age: int, agegroups: list[dict]) -> dict | None:
    """Find the 5-year Masters bracket for this age."""
    for ag in agegroups:
        amin = ag["amin"]
        amax = ag["amax"]
        if amin is None:
            continue
        hi = 999 if (amax is None or amax < 0) else amax
        if amin <= age <= hi:
            return ag
    return None


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--lxf", required=True, type=Path, help="SPLASH results export (.lxf)")
    ap.add_argument("--out", required=True, type=Path, help="Output Lenex for Masters import")
    args = ap.parse_args()

    # Parse input
    if args.lxf.suffix.lower() == ".lxf":
        with zipfile.ZipFile(args.lxf) as z:
            names = z.namelist()
            xml = z.read(names[0])
    else:
        xml = args.lxf.read_bytes()

    root = ET.fromstring(xml)
    meet = root.find(".//MEET")
    age_date_el = meet.find("AGEDATE")
    age_date = dt.date.fromisoformat(age_date_el.get("value")) if age_date_el is not None else dt.date(2026, 12, 31)

    # Build event structures
    prelim_events = {}   # eventid -> {styleid, gender, masters_agids[]}
    masters_finals = {}  # (styleid, gender) -> {eventid, agegroups[]}

    for ev in meet.findall(".//EVENT"):
        eid = ev.get("eventid")
        rnd = ev.get("round")
        etype = ev.get("type", "")
        ss = ev.find("SWIMSTYLE")
        styleid = ss.get("swimstyleid")
        gender = ev.get("gender")

        agegroups = []
        for ag in ev.findall(".//AGEGROUP"):
            amin = int(ag.get("agemin")) if ag.get("agemin") else None
            amax = int(ag.get("agemax")) if ag.get("agemax") else None
            agegroups.append({"id": ag.get("agegroupid"), "amin": amin, "amax": amax})

        if rnd == "PRE":
            masters_ags = [a for a in agegroups if a["amin"] is not None and 25 <= a["amin"] < 100]
            if masters_ags:
                prelim_events[eid] = {"styleid": styleid, "gender": gender, "agegroups": masters_ags}

        if rnd == "TIM" and etype == "MASTERS":
            masters_finals[(styleid, gender)] = {"eventid": eid, "agegroups": agegroups}

    print(f"  Prelim events with Masters brackets: {len(prelim_events)}")
    print(f"  Masters final events: {len(masters_finals)}")

    # Find athletes with results in prelim events who are Masters age
    clubs = meet.find("CLUBS")
    transfers = []  # {club, athlete, final_eventid, final_agid, swimtime}

    for club in clubs.findall("CLUB"):
        for ath in club.findall(".//ATHLETE"):
            birthdate = ath.get("birthdate")
            athlete_age = age_at(birthdate, age_date)
            if athlete_age is None or athlete_age < 25:
                continue

            for result in ath.findall(".//RESULT"):
                eid = result.get("eventid")
                if eid not in prelim_events:
                    continue
                swimtime = result.get("swimtime", "00:00:00.00")
                if swimtime == "00:00:00.00":
                    continue

                info = prelim_events[eid]
                final = masters_finals.get((info["styleid"], info["gender"]))
                if final is None:
                    continue

                # Find the right agegroup in the final
                bracket = find_bracket(athlete_age, final["agegroups"])
                if bracket is None:
                    continue

                transfers.append({
                    "club_name": club.get("name"),
                    "club_code": club.get("code", club.get("name", "")[:10]),
                    "athlete": ath,
                    "final_eventid": final["eventid"],
                    "final_agid": bracket["id"],
                    "swimtime": swimtime,
                    "prelim_eid": eid,
                })

    print(f"  Masters transfers: {len(transfers)}")
    if not transfers:
        print("  No Masters results to transfer.")
        sys.exit(0)

    # Generate output Lenex
    out_root = ET.Element("LENEX", {"version": "3.0"})
    out_meets = ET.SubElement(out_root, "MEETS")
    out_meet = ET.SubElement(out_meets, "MEET", {
        "name": meet.get("name", ""),
        "city": meet.get("city", ""),
        "nation": meet.get("nation", "CAN"),
        "course": meet.get("course", "LCM"),
    })
    ET.SubElement(out_meet, "AGEDATE", {"value": age_date.isoformat(), "type": "CAN.FNQ"})

    # Group transfers by club
    by_club: dict[str, list] = {}
    for t in transfers:
        by_club.setdefault(t["club_name"], []).append(t)

    out_clubs = ET.SubElement(out_meet, "CLUBS")
    for club_name, club_transfers in sorted(by_club.items()):
        club_code = club_transfers[0]["club_code"]
        out_club = ET.SubElement(out_clubs, "CLUB", {
            "name": club_name, "code": club_code, "nation": "CAN",
        })
        out_aths = ET.SubElement(out_club, "ATHLETES")

        # Group by athlete
        by_ath: dict[str, list] = {}
        for t in club_transfers:
            aid = t["athlete"].get("athleteid")
            by_ath.setdefault(aid, []).append(t)

        for aid, ath_transfers in by_ath.items():
            ath = ath_transfers[0]["athlete"]
            attrs = {
                "athleteid": aid,
                "firstname": ath.get("firstname", ""),
                "lastname": ath.get("lastname", ""),
                "gender": ath.get("gender", "M"),
                "birthdate": ath.get("birthdate", ""),
            }
            if ath.get("license"):
                attrs["license"] = ath.get("license")
            out_ath = ET.SubElement(out_aths, "ATHLETE", attrs)
            out_entries = ET.SubElement(out_ath, "ENTRIES")

            for t in ath_transfers:
                ET.SubElement(out_entries, "ENTRY", {
                    "eventid": t["final_eventid"],
                    "agegroupid": t["final_agid"],
                    "entrytime": t["swimtime"],
                    "entrycourse": meet.get("course", "LCM"),
                })

    # Write output
    xml_str = minidom.parseString(
        ET.tostring(out_root, encoding="unicode")
    ).toprettyxml(indent="  ")
    xml_str = "\n".join(l for l in xml_str.splitlines() if l.strip())

    out_path = args.out
    if out_path.suffix.lower() == ".lxf":
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(out_path.stem + ".lef", xml_str)
    else:
        out_path.write_text(xml_str, encoding="utf-8")

    print(f"  Written: {out_path}")
    print(f"\n  Import this file into SPLASH, then delete the Masters")
    print(f"  athletes from the prelim events manually.")


if __name__ == "__main__":
    main()
