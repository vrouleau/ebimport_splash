#!/usr/bin/env python3
"""Generate a Lenex 3.0 (.lef/.lxf) inscription file from an xlsx +
a meet structure (.lxf exported from SPLASH).

Usage:
    python load_to_lenex.py --xlsx CPLC2026FINAL.xlsx --meet splash_results_meet.lxf --out meet.lxf
"""
from __future__ import annotations

import argparse
import datetime as dt
import json
import sys
import zipfile
from pathlib import Path
from xml.dom import minidom
from xml.etree import ElementTree as ET

sys.path.insert(0, str(Path(__file__).parent))
from core import (
    read_attendees, IssueCollector, TemplateIndex, MDB,
    pick_agegroup_for_individual, pick_agegroup_for_relay,
    norm_key, age_at, EventKey, Inscription,
    GENDER_MALE, GENDER_FEMALE, GENDER_MIXED,
    TemplateEvent, TemplateAgeGroup, TemplateStyle,
)
from common import aggregate, run_sanity_checks, run_validation, run_cross_row_checks
from meet_parser import parse_meet_lxf, ParsedMeet
from collections import defaultdict
import re


class MeetLxfTemplate:
    """Adapter: wraps ParsedMeet to provide the same interface as TemplateIndex/TemplateJSON."""

    def __init__(self, meet: ParsedMeet):
        self._meet = meet
        self.styles_by_uid: dict[int, TemplateStyle] = {}
        self.events_by_uid_gender: dict[tuple, list[TemplateEvent]] = defaultdict(list)

        for ev in meet.all_events:
            uid = ev.swimstyleid
            if uid not in self.styles_by_uid:
                self.styles_by_uid[uid] = TemplateStyle(
                    swim_style_id=0, uniqueid=uid,
                    distance=ev.distance, relay_count=ev.relaycount,
                    name=ev.style_name)

            ags = [TemplateAgeGroup(
                agegroup_id=ag.agegroupid, amin=ag.agemin, amax=ag.agemax, gender=ev.gender_int
            ) for ag in ev.agegroups]

            tev = TemplateEvent(
                swim_event_id=ev.eventid, swim_style_id=0,
                uniqueid=uid, gender=ev.gender_int,
                event_number=ev.number,
                round=2 if ev.is_prelim else (1 if ev.round == "TIM" else 9),
                masters=ev.is_masters,
                session_id=None,
                agegroups=ags)
            self.events_by_uid_gender[(uid, ev.gender_int)].append(tev)

    def find_event(self, uid: int, gender: int, masters: bool = False) -> TemplateEvent | None:
        evs = self.events_by_uid_gender.get((uid, gender), [])
        # Prefer prelim (round=2) matching masters flag
        for e in evs:
            if e.masters == masters and e.round == 2:
                return e
        # For non-masters: fall back to any prelim (shared prelim marked as MASTERS)
        if not masters:
            for e in evs:
                if e.round == 2:
                    return e
        # Exact masters match (timed final for masters-only events)
        for e in evs:
            if e.masters == masters:
                return e
        return evs[0] if evs else None

    def find_prelim_for_dual_entry(self, uid: int, gender: int) -> TemplateEvent | None:
        for e in self.events_by_uid_gender.get((uid, gender), []):
            if e.masters or e.round != 2:
                continue
            for a in e.agegroups:
                if a.amin is not None and 25 <= a.amin < 100:
                    return e
        return None

# Lenex constants
MEET_NAME   = "Championnats canadiens"
MEET_CITY   = "Québec"
MEET_NATION = "CAN"
MEET_COURSE = "LCM"


def ms_to_lenex(ms: int | None) -> str:
    """Convert milliseconds to Lenex time format HH:MM:SS.hh (hundredths)."""
    if ms is None or ms <= 0:
        return ""
    h = ms // 3_600_000
    rem = ms % 3_600_000
    m = rem // 60_000
    rem = rem % 60_000
    s = rem // 1000
    cs = (rem % 1000) // 10
    return f"{h:02d}:{m:02d}:{s:02d}.{cs:02d}"


def lenex_gender(g: int) -> str:
    return {GENDER_MALE: "M", GENDER_FEMALE: "F", GENDER_MIXED: "X"}.get(g, "A")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--meet", type=Path,
                    help="Meet structure .lxf (exported from SPLASH)")
    ap.add_argument("--template", type=Path,
                    help="(deprecated) Template JSON")
    ap.add_argument("--mdb", type=Path,
                    help="(deprecated) Template .mdb")
    ap.add_argument("--out", required=True, type=Path)
    ap.add_argument("--zip", action="store_true")
    args = ap.parse_args()

    if not args.xlsx.exists():
        sys.exit(f"xlsx not found: {args.xlsx}")

    # Parse xlsx
    issues = IssueCollector()
    inscriptions, name_to_dob = read_attendees(args.xlsx, issues)
    print(f"  {len(inscriptions)} race inscriptions")

    # Open template — prefer meet .lxf
    db = None
    template = None
    if args.meet and args.meet.exists():
        meet = parse_meet_lxf(args.meet)
        template = MeetLxfTemplate(meet)
    elif args.template and args.template.exists():
        sys.exit("--template is deprecated. Use --meet with a SPLASH meet export .lxf")
    elif args.mdb and args.mdb.exists():
        db = MDB(args.mdb, dry_run=True)
        template = TemplateIndex(db)
    elif not args.meet:
        # No meet file — parse-only validation (no event matching)
        pass
    else:
        sys.exit("Provide --meet (SPLASH meet export .lxf)")

    # Sanity checks (only for MDB-based template)
    import core
    if db:
        sanity_errors = run_sanity_checks(template)
        if sanity_errors:
            for e in sanity_errors:
                print(f"  FATAL: {e}")
            sys.exit(2)

    AGE_DATE = core.AGE_DATE or dt.date(2026, 12, 31)
    core.AGE_DATE = AGE_DATE

    # Aggregate
    data = aggregate(inscriptions, issues, name_to_dob=name_to_dob)
    clubs = data.clubs
    athletes = data.athletes
    name_to_key = data.name_to_key
    events_in_xlsx = data.events_in_xlsx
    ind_entries = data.ind_entries
    relay_squads = data.relay_squads

    # Validation (requires template)
    if template:
        fatal = run_validation(events_in_xlsx, template)
        if fatal:
            print("\n  FATAL: template/xlsx mismatch")
            for f in fatal:
                print(f"  - {f}")
            if db: db.close()
            sys.exit(2)

    # Cross-row checks (requires template)
    if template:
        run_cross_row_checks(data, template, issues)

    # If no template, stop here (parse-only validation)
    if not template:
        print(f"\n  {len(clubs)} clubs, {len(athletes)} athletes")
        print(f"  {len(ind_entries)} individual entries")
        print(f"  {sum(len(s) for s in relay_squads.values())} relay squads")
        print(f"\n  (No meet .lxf — skipped event matching)")
        issues.report("Issues found while validating", full=True)
        if db: db.close()
        sys.exit(0)

    # Dedup individual entries (keep best time)
    best_by: dict[tuple, int | None] = {}
    for akey, ekey, ms in ind_entries:
        cur = best_by.get((akey, ekey))
        if cur is None or (ms is not None and (cur is None or ms < cur)):
            best_by[(akey, ekey)] = ms

    print(f"  {len(clubs)} clubs, {len(athletes)} athletes")
    print(f"  {len(best_by)} individual entries")
    print(f"  {sum(len(s) for s in relay_squads.values())} relay squads")

    # --- Build Lenex XML ---
    root = ET.Element("LENEX", {
        "version": "3.0",
        "created": dt.datetime.now().strftime("%Y-%m-%dT%H:%M:%S"),
    })
    ctor = ET.SubElement(root, "CONSTRUCTOR", {
        "name": "ebimport_splash", "registration": "", "version": "2.0"})
    ET.SubElement(ctor, "CONTACT", {"name": "ebimport_splash", "email": ""})

    meets = ET.SubElement(root, "MEETS")
    meet = ET.SubElement(meets, "MEET", {
        "name": MEET_NAME, "city": MEET_CITY, "nation": MEET_NATION,
        "course": MEET_COURSE, "timing": "AUTOMATIC",
    })
    ET.SubElement(meet, "AGEDATE", {"value": AGE_DATE.isoformat(), "type": "CAN.FNQ"})
    ET.SubElement(meet, "POOL", {"lanemin": "1", "lanemax": "8"})

    # Sessions + Events from template
    sessions_xml = ET.SubElement(meet, "SESSIONS")
    # Query sessions from template
    session_events: dict[int, list] = defaultdict(list)
    for ev_list in template.events_by_uid_gender.values():
        for tev in ev_list:
            session_events[tev.session_id or 0].append(tev)

    event_id_map = {}  # template swim_event_id -> lenex eventid string
    for ses_id, tevents in sorted(session_events.items()):
        ses_xml = ET.SubElement(sessions_xml, "SESSION", {
            "number": str(ses_id), "date": AGE_DATE.isoformat(),
            "course": MEET_COURSE,
        })
        evts_xml = ET.SubElement(ses_xml, "EVENTS")
        for tev in sorted(tevents, key=lambda e: e.event_number or 0):
            eid_str = str(tev.swim_event_id)
            event_id_map[tev.swim_event_id] = eid_str
            style = template.styles_by_uid.get(tev.uniqueid)
            ev_xml = ET.SubElement(evts_xml, "EVENT", {
                "eventid": eid_str,
                "number": str(tev.event_number or 0),
                "gender": lenex_gender(tev.gender),
                "round": "TIM" if tev.round == 1 else "PRE",
            })
            ss_attrs = {
                "stroke": "UNKNOWN",
                "distance": str(style.distance or 0),
                "relaycount": str(style.relay_count or 1),
                "name": style.name or "",
            }
            ET.SubElement(ev_xml, "SWIMSTYLE", ss_attrs)
            # Age groups
            if tev.agegroups:
                ags_xml = ET.SubElement(ev_xml, "AGEGROUPS")
                for ag in tev.agegroups:
                    ET.SubElement(ags_xml, "AGEGROUP", {
                        "agegroupid": str(ag.agegroup_id),
                        "agemin": str(ag.amin if ag.amin is not None else -1),
                        "agemax": str(ag.amax if ag.amax is not None else -1),
                        "gender": lenex_gender(ag.gender) if ag.gender else "A",
                    })

    # Clubs + Athletes + Entries + Relays
    clubs_xml = ET.SubElement(meet, "CLUBS")
    athlete_id_map: dict[tuple, int] = {}
    uid_counter = 10000

    # Derive athlete gender from individual entries
    athlete_gender: dict[tuple, int] = {}
    for (akey, ekey), _ in best_by.items():
        ev = events_in_xlsx[ekey]
        if not ev.is_relay and akey not in athlete_gender:
            athlete_gender[akey] = ev.gender

    for cnorm, cname in sorted(clubs.items(), key=lambda kv: kv[1].lower()):
        club_xml = ET.SubElement(clubs_xml, "CLUB", {
            "name": cname, "code": cname[:10], "nation": MEET_NATION,
        })

        # Athletes in this club (only canonical keys — no duplicates)
        club_aths = [(ak, a) for ak, a in athletes.items()
                     if norm_key(a.club) == cnorm
                     and name_to_key.get(norm_key(a.first, a.last)) == ak]
        if not club_aths:
            continue

        aths_xml = ET.SubElement(club_xml, "ATHLETES")
        for akey, ins in sorted(club_aths, key=lambda x: (x[1].last, x[1].first)):
            uid_counter += 1
            athlete_id_map[akey] = uid_counter
            attrs = {
                "athleteid": str(uid_counter),
                "firstname": ins.first,
                "lastname": ins.last,
                "gender": lenex_gender(athlete_gender.get(akey, 0)),
            }
            if ins.birthdate:
                attrs["birthdate"] = ins.birthdate.strftime("%Y-%m-%d")
            if ins.license:
                # Check if athlete has any Masters entries
                is_masters = any(events_in_xlsx[ek].age_code == "MASTERS"
                                 for (ak, ek), _ in best_by.items() if ak == akey)
                attrs["license"] = ins.license or ""
            ath_xml = ET.SubElement(aths_xml, "ATHLETE", attrs)
            if is_masters:
                ET.SubElement(ath_xml, "HANDICAP", {"exception": "X"})

            # Individual entries for this athlete
            my_entries = [(ekey, ms) for (ak, ekey), ms in best_by.items()
                          if ak == akey]
            if my_entries:
                entries_xml = ET.SubElement(ath_xml, "ENTRIES")
                for ekey, ms in my_entries:
                    ev = events_in_xlsx[ekey]
                    # All athletes go to prelim — Masters are marked via
                    # HANDICAP exception='X' and transferred after prelims by VBS
                    tevent = template.find_event(
                        ev.uniqueid, ev.gender, masters=False)
                    if tevent is None:
                        continue
                    athlete_age = age_at(ins.birthdate)
                    # If the event is a prelim, Masters go to [19-99]
                    # If it's a timed final (Masters-only event like UID 541),
                    # use the actual Masters bracket
                    if tevent.round == 2:
                        ag_code = "OPEN" if ev.age_code == "MASTERS" else ev.age_code
                    else:
                        ag_code = ev.age_code
                    ag = pick_agegroup_for_individual(
                        tevent, ag_code, athlete_age)
                    if ag is None:
                        continue
                    entry_attrs = {
                        "eventid": str(tevent.swim_event_id),
                        "agegroupid": str(ag.agegroup_id),
                    }
                    et = ms_to_lenex(ms)
                    if et:
                        entry_attrs["entrytime"] = et
                        entry_attrs["entrycourse"] = MEET_COURSE
                    ET.SubElement(entries_xml, "ENTRY", entry_attrs)

        # Relays for this club
        club_relays = [(ekey, squads) for (cn, ekey), squads
                       in relay_squads.items() if cn == cnorm]
        if club_relays:
            relays_xml = ET.SubElement(club_xml, "RELAYS")
            for ekey, squads in club_relays:
                ev = events_in_xlsx[ekey]
                # Route relays same as individuals: Masters go to prelim
                # (Masters bracket), non-Masters go to prelim too.
                if ev.age_code == "MASTERS":
                    tevent = template.find_prelim_for_dual_entry(
                        ev.uniqueid, ev.gender)
                    if tevent is None:
                        tevent = template.find_event(
                            ev.uniqueid, ev.gender, masters=True)
                else:
                    tevent = template.find_event(
                        ev.uniqueid, ev.gender, masters=False)
                if tevent is None:
                    continue
                style = template.styles_by_uid[ev.uniqueid]
                relay_size = style.relay_count or 4

                for team_no, squad in enumerate(squads, start=1):
                    if len(squad) < relay_size:
                        continue
                    # Route age group
                    ages = [age_at(athletes[ak].birthdate)
                            for ak, _ in squad[:relay_size]]
                    age_sum = sum(a for a in ages if a is not None) if all(a is not None for a in ages) else None
                    youngest = min((a for a in ages if a is not None), default=None)
                    ag = pick_agegroup_for_relay(
                        tevent, ev.age_code, age_sum, oldest_age=youngest)
                    if ag is None:
                        continue

                    # For Lenex relay import, SPLASH matches the relay's
                    # agemin/agemax to an AGEGROUP on the event.  For
                    # Masters relays on a prelim event, use the Open
                    # bracket [19,99] so SPLASH accepts them.
                    lenex_ag = ag
                    if ev.age_code == "MASTERS" and not tevent.masters:
                        # Find the [19,99] bracket on this prelim
                        for a in tevent.agegroups:
                            if a.amin == 19 and a.amax in (99, -1, None):
                                lenex_ag = a; break

                    relay_name = "/".join(
                        athletes[ak].last for ak, _ in squad[:relay_size])
                    if ev.age_code == "1518":
                        rel_amin, rel_amax = 15, 18
                        rel_totalmin, rel_totalmax = -1, -1
                    elif ev.age_code == "OPEN" or (ev.age_code == "MASTERS" and not tevent.masters):
                        rel_amin, rel_amax = 19, 99
                        rel_totalmin, rel_totalmax = -1, -1
                    else:
                        # Masters relay on Masters final
                        if ag.amin is not None and ag.amin < 100:
                            # Individual-style brackets (Corde duo)
                            rel_amin = ag.amin if ag.amin is not None else -1
                            rel_amax = ag.amax if ag.amax is not None else -1
                            rel_totalmin, rel_totalmax = -1, -1
                        else:
                            # Age-sum brackets
                            rel_amin, rel_amax = -1, -1
                            rel_totalmin = ag.amin if ag.amin is not None else -1
                            rel_totalmax = ag.amax if ag.amax is not None else -1
                    rel_attrs = {
                        "number": str(team_no),
                        "name": relay_name[:50],
                        "gender": lenex_gender(ev.gender),
                        "agemin": str(rel_amin),
                        "agemax": str(rel_amax),
                        "agetotalmin": str(rel_totalmin),
                        "agetotalmax": str(rel_totalmax),
                    }
                    rel_xml = ET.SubElement(relays_xml, "RELAY", rel_attrs)
                    # Entry — use the relay-row's team time (squad[0][1]),
                    # not a sum of teammates' individual best times.
                    entry_time = squad[0][1] if squad else None
                    ents_xml = ET.SubElement(rel_xml, "ENTRIES")
                    entry_attrs = {
                        "eventid": str(tevent.swim_event_id),
                        "agegroupid": str(lenex_ag.agegroup_id),
                    }
                    et = ms_to_lenex(entry_time)
                    if et:
                        entry_attrs["entrytime"] = et
                        entry_attrs["entrycourse"] = MEET_COURSE
                    entry_xml = ET.SubElement(ents_xml, "ENTRY", entry_attrs)
                    # Positions
                    pos_xml = ET.SubElement(entry_xml, "RELAYPOSITIONS")
                    for leg, (ak, _) in enumerate(squad[:relay_size], start=1):
                        aid = athlete_id_map.get(ak)
                        if aid is None:
                            continue
                        ET.SubElement(pos_xml, "RELAYPOSITION", {
                            "number": str(leg),
                            "athleteid": str(aid),
                        })

    if db: db.close()

    # --- Write output ---
    xml_str = minidom.parseString(
        ET.tostring(root, encoding="unicode")
    ).toprettyxml(indent="  ", encoding=None)
    # Remove extra blank lines from minidom
    xml_str = "\n".join(l for l in xml_str.splitlines() if l.strip())

    out_path = args.out
    if args.zip or out_path.suffix.lower() == ".lxf":
        with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("meet.lef", xml_str)
        print(f"  Written: {out_path} (zipped)")
    else:
        out_path.write_text(xml_str, encoding="utf-8")
        print(f"  Written: {out_path}")

    # Masters identification: HANDICAP exception='X' on athlete element
    pass

    issues.report("Issues found while generating Lenex", full=True)


if __name__ == "__main__":
    main()
