"""
Simulate meet results: generate a Lenex results file from a seeded MDB.

For each SWIMRESULT in the MDB, generates a SWIMTIME that is ±5% of
the athlete's ENTRYTIME (or a random time if NT). 5% of athletes get
a random DQ.

Usage:
    python simulate_results.py --mdb meet.mdb --out results.lxf
"""
from __future__ import annotations

import argparse
import random
import sys
import zipfile
from pathlib import Path
from xml.dom import minidom
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).parent))
from access_parser import AccessParser


def ms_to_lenex(ms: int) -> str:
    h = ms // 3_600_000
    rem = ms % 3_600_000
    m = rem // 60_000
    rem = rem % 60_000
    s = rem // 1000
    cs = (rem % 1000) // 10
    return f"{h:02d}:{m:02d}:{s:02d}.{cs:02d}"


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mdb", required=True, type=Path)
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--dq-rate", type=float, default=0.05, help="DQ probability (default 0.05)")
    ap.add_argument("--seed", type=int, default=None, help="Random seed")
    args = ap.parse_args()

    if args.seed is not None:
        random.seed(args.seed)

    db = AccessParser(str(args.mdb))
    sr = db.parse_table("SWIMRESULT")
    ath_t = db.parse_table("ATHLETE")
    ev_t = db.parse_table("SWIMEVENT")
    cl_t = db.parse_table("CLUB")
    ss_t = db.parse_table("SWIMSTYLE")

    # Build lookups
    athletes = {}
    for i in range(len(ath_t["ATHLETEID"])):
        aid = int(ath_t["ATHLETEID"][i])
        athletes[aid] = {
            "first": ath_t["FIRSTNAME"][i] or "",
            "last": ath_t["LASTNAME"][i] or "",
            "gender": int(ath_t["GENDER"][i]) if ath_t["GENDER"][i] else 0,
            "birthdate": ath_t["BIRTHDATE"][i],
            "license": ath_t["LICENSE"][i],
            "club_id": int(ath_t["CLUBID"][i]) if ath_t["CLUBID"][i] else 0,
        }

    clubs = {}
    for i in range(len(cl_t["CLUBID"])):
        clubs[int(cl_t["CLUBID"][i])] = {
            "name": cl_t["NAME"][i] or "",
            "code": cl_t["CODE"][i] or "",
        }

    events = {}
    for i in range(len(ev_t["SWIMEVENTID"])):
        eid = int(ev_t["SWIMEVENTID"][i])
        events[eid] = {
            "number": ev_t["EVENTNUMBER"][i],
            "gender": int(ev_t["GENDER"][i]) if ev_t["GENDER"][i] else 0,
            "round": int(ev_t["ROUND"][i]) if ev_t["ROUND"][i] else 0,
            "style_id": int(ev_t["SWIMSTYLEID"][i]) if ev_t["SWIMSTYLEID"][i] else 0,
            "session_id": int(ev_t["SWIMSESSIONID"][i]) if ev_t["SWIMSESSIONID"][i] else 0,
        }

    styles = {}
    for i in range(len(ss_t["SWIMSTYLEID"])):
        sid = int(ss_t["SWIMSTYLEID"][i])
        styles[sid] = {
            "uid": int(ss_t["UNIQUEID"][i]) if ss_t["UNIQUEID"][i] else 0,
            "name": ss_t["NAME"][i] or "",
            "distance": int(ss_t["DISTANCE"][i]) if ss_t["DISTANCE"][i] else 0,
        }

    # Generate results
    INT_MAX = 2147483647
    results = []  # (athlete_id, event_id, swimtime_ms, status)
    for i in range(len(sr["SWIMRESULTID"])):
        aid = int(sr["ATHLETEID"][i]) if sr["ATHLETEID"][i] else 0
        eid = int(sr["SWIMEVENTID"][i]) if sr["SWIMEVENTID"][i] else 0
        entry_time = sr["ENTRYTIME"][i]

        if aid == 0 or eid == 0:
            continue

        # DQ?
        if random.random() < args.dq_rate:
            results.append((aid, eid, 0, "DSQ"))
            continue

        # Generate swim time
        if entry_time and int(entry_time) > 0 and int(entry_time) < INT_MAX:
            base = int(entry_time)
            variation = base * 0.05
            swim_ms = max(1000, int(base + random.uniform(-variation, variation)))
        else:
            # NT athlete — generate a random slow time based on event distance
            ev = events.get(eid, {})
            style = styles.get(ev.get("style_id", 0), {})
            dist = style.get("distance", 100)
            # Rough: 15s per 25m for a slow swimmer
            base_ms = (dist / 25) * 15000
            swim_ms = int(base_ms + random.uniform(0, base_ms * 0.3))

        results.append((aid, eid, swim_ms, "OK"))

    print(f"  {len(results)} results generated ({sum(1 for _,_,_,s in results if s=='DSQ')} DQ)")

    # Build Lenex results XML
    root = ET.Element("LENEX", {"version": "3.0"})
    meets = ET.SubElement(root, "MEETS")
    meet = ET.SubElement(meets, "MEET", {
        "name": "Simulated Results",
        "nation": "CAN",
        "course": "LCM",
    })

    # Add sessions/events structure (required for SPLASH import)
    sessions_xml = ET.SubElement(meet, "SESSIONS")
    # Group events by session
    ses_t = db.parse_table("SWIMSESSION")
    session_map = {}
    for i in range(len(ses_t["SWIMSESSIONID"])):
        sid = int(ses_t["SWIMSESSIONID"][i])
        session_map[sid] = {
            "number": int(ses_t["SESSIONNUMBER"][i]) if ses_t["SESSIONNUMBER"][i] else 1,
            "name": ses_t["NAME"][i] or "",
        }

    events_by_session = {}
    for i in range(len(ev_t["SWIMEVENTID"])):
        eid = int(ev_t["SWIMEVENTID"][i])
        sid = int(ev_t["SWIMSESSIONID"][i]) if ev_t["SWIMSESSIONID"][i] else 0
        events_by_session.setdefault(sid, []).append(eid)

    for sid, eids in sorted(events_by_session.items()):
        ses_info = session_map.get(sid, {"number": 1, "name": ""})
        ses_xml = ET.SubElement(sessions_xml, "SESSION", {
            "number": str(ses_info["number"]),
            "course": "LCM",
        })
        evts_xml = ET.SubElement(ses_xml, "EVENTS")
        for eid in eids:
            ev = events.get(eid)
            if not ev:
                continue
            style = styles.get(ev["style_id"], {})
            gender_str = {1: "M", 2: "F", 3: "X"}.get(ev["gender"], "X")
            round_str = {1: "TIM", 2: "PRE", 9: "FIN"}.get(ev["round"], "TIM")
            ev_xml = ET.SubElement(evts_xml, "EVENT", {
                "eventid": str(eid),
                "number": str(ev["number"] or 0),
                "gender": gender_str,
                "round": round_str,
            })
            ET.SubElement(ev_xml, "SWIMSTYLE", {
                "distance": str(style.get("distance", 0)),
                "relaycount": "1",
                "stroke": "UNKNOWN",
                "name": style.get("name", ""),
            })

    # Group results by club -> athlete
    by_club = {}  # club_id -> {athlete_id -> [(event_id, swim_ms, status)]}
    for aid, eid, swim_ms, status in results:
        ath = athletes.get(aid)
        if not ath:
            continue
        cid = ath["club_id"]
        by_club.setdefault(cid, {}).setdefault(aid, []).append((eid, swim_ms, status))

    clubs_xml = ET.SubElement(meet, "CLUBS")
    for cid in sorted(by_club.keys()):
        club = clubs.get(cid, {"name": "Unknown", "code": "UNK"})
        club_xml = ET.SubElement(clubs_xml, "CLUB", {
            "name": club["name"], "code": club["code"], "nation": "CAN",
        })
        aths_xml = ET.SubElement(club_xml, "ATHLETES")

        for aid in sorted(by_club[cid].keys()):
            ath = athletes[aid]
            gender_str = {1: "M", 2: "F"}.get(ath["gender"], "M")
            attrs = {
                "athleteid": str(aid),
                "firstname": str(ath["first"]),
                "lastname": str(ath["last"]),
                "gender": gender_str,
            }
            if ath["birthdate"]:
                bd = ath["birthdate"]
                if hasattr(bd, "strftime"):
                    attrs["birthdate"] = bd.strftime("%Y-%m-%d")
                else:
                    attrs["birthdate"] = str(bd)[:10]
            if ath["license"]:
                attrs["license"] = str(ath["license"])

            ath_xml = ET.SubElement(aths_xml, "ATHLETE", attrs)
            results_xml = ET.SubElement(ath_xml, "RESULTS")

            for eid, swim_ms, status in by_club[cid][aid]:
                res_attrs = {"eventid": str(eid)}
                if status == "DSQ":
                    res_attrs["status"] = "DSQ"
                    res_attrs["swimtime"] = "00:00:00.00"
                else:
                    res_attrs["swimtime"] = ms_to_lenex(swim_ms)
                ET.SubElement(results_xml, "RESULT", res_attrs)

    # Write
    xml_str = minidom.parseString(
        ET.tostring(root, encoding="unicode")
    ).toprettyxml(indent="  ")
    xml_str = "\n".join(l for l in xml_str.splitlines() if l.strip())

    if args.out.suffix.lower() == ".lxf":
        with zipfile.ZipFile(args.out, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("results.lef", xml_str)
    else:
        args.out.write_text(xml_str, encoding="utf-8")

    print(f"  Written: {args.out}")


if __name__ == "__main__":
    main()
