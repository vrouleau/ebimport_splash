#!/usr/bin/env python3
"""Post-prelim script: transfer Masters results from prelim to final.

For each Masters athlete on a non-Masters prelim event (in a Masters
age bracket) that has a SWIMTIME recorded:
  1. Read the SWIMTIME (and REACTIONTIME, RESULTSTATUS)
  2. Delete the prelim SWIMRESULT row
  3. Create a SWIMRESULT on the corresponding Masters timed-final event
     with that SWIMTIME, assigned to a HEAT and LANE (max 8 per heat)

Usage:
    python copy_prelim_to_masters_final.py --mdb meet.mdb [--dry-run]

Idempotent: rows without SWIMTIME are skipped.
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import math
import os
import sys

import jaydebeapi

AGE_DATE = dt.date(2026, 5, 31)
INT_MAX = 2147483647


def connect(mdb_path: str):
    ucanaccess_dir = os.environ.get(
        "UCANACCESS_DIR", "/opt/ucanaccess/UCanAccess-5.0.1.bin")
    jars = (glob.glob(f"{ucanaccess_dir}/ucanaccess-*.jar") +
            glob.glob(f"{ucanaccess_dir}/lib/*.jar") +
            glob.glob(f"{ucanaccess_dir}/*.jar"))
    jars = list(dict.fromkeys(jars))
    return jaydebeapi.connect(
        "net.ucanaccess.jdbc.UcanaccessDriver",
        f"jdbc:ucanaccess://{mdb_path};openExclusive=false", [], jars)


def age_at(birthdate, ref=AGE_DATE) -> int | None:
    if birthdate is None:
        return None
    if isinstance(birthdate, str):
        birthdate = dt.date.fromisoformat(birthdate[:10])
    elif hasattr(birthdate, "date"):
        birthdate = birthdate.date()
    y = ref.year - birthdate.year
    if (ref.month, ref.day) < (birthdate.month, birthdate.day):
        y -= 1
    return y


def main():
    ap = argparse.ArgumentParser(description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mdb", required=True, help="Path to the meet .mdb")
    ap.add_argument("--dry-run", action="store_true",
                    help="Show what would be done without writing")
    args = ap.parse_args()

    conn = connect(args.mdb)
    if args.dry_run:
        conn.jconn.setAutoCommit(False)
    c = conn.cursor()

    # Get BS_GLOBAL_UID
    c.execute("SELECT LASTUID FROM BSUIDTABLE WHERE NAME='BS_GLOBAL_UID'")
    next_uid = int(c.fetchone()[0]) + 1

    # Find prelim events with Masters brackets
    c.execute("""
        SELECT e.SWIMEVENTID, e.SWIMSTYLEID, e.GENDER,
               a.AGEGROUPID, a.AGEMIN, a.AGEMAX
        FROM SWIMEVENT e
        INNER JOIN AGEGROUP a ON a.SWIMEVENTID = e.SWIMEVENTID
        WHERE e.ROUND = 2 AND e.MASTERS = 'F'
          AND a.AGEMIN >= 25 AND a.AGEMIN < 100
    """)
    prelim_masters_brackets = {}
    prelim_meta = {}
    for eid, styid, gen, agid, amin, amax in c.fetchall():
        eid, styid, gen = int(eid), int(styid), int(gen)
        prelim_masters_brackets.setdefault(eid, []).append(
            (int(agid), int(amin), int(amax)))
        prelim_meta[eid] = (styid, gen)

    if not prelim_masters_brackets:
        print("No non-Masters prelim events with Masters brackets found.")
        conn.close()
        return

    # Find Masters timed-final events
    c.execute("""
        SELECT e.SWIMEVENTID, e.SWIMSTYLEID, e.GENDER, e.EVENTNUMBER,
               a.AGEGROUPID, a.AGEMIN, a.AGEMAX
        FROM SWIMEVENT e
        INNER JOIN AGEGROUP a ON a.SWIMEVENTID = e.SWIMEVENTID
        WHERE e.ROUND = 1 AND e.MASTERS = 'T'
          AND a.AGEMIN >= 25 AND a.AGEMIN < 100
    """)
    masters_finals = {}
    for eid, styid, gen, enum, agid, amin, amax in c.fetchall():
        key = (int(styid), int(gen))
        if key not in masters_finals:
            masters_finals[key] = {"eid": int(eid), "enum": enum, "brackets": []}
        masters_finals[key]["brackets"].append((int(agid), int(amin), int(amax)))

    # Collect all transfers grouped by final event
    # transfers_by_final[final_eid] = [(ath_id, target_agid, swimtime, reaction, status, entrytime, prelim_srid)]
    transfers_by_final: dict[int, list] = {}
    prelim_rows_to_delete = []

    for prelim_eid, brackets in prelim_masters_brackets.items():
        styid, gender = prelim_meta[prelim_eid]
        final_info = masters_finals.get((styid, gender))
        if final_info is None:
            continue

        final_eid = final_info["eid"]
        final_brackets = final_info["brackets"]

        ag_ids = [b[0] for b in brackets]
        placeholders = ",".join("?" * len(ag_ids))
        c.execute(f"""
            SELECT sr.SWIMRESULTID, sr.ATHLETEID, sr.AGEGROUPID,
                   sr.SWIMTIME, sr.REACTIONTIME, sr.RESULTSTATUS,
                   sr.ENTRYTIME, ath.BIRTHDATE
            FROM SWIMRESULT sr
            INNER JOIN ATHLETE ath ON ath.ATHLETEID = sr.ATHLETEID
            WHERE sr.SWIMEVENTID = ? AND sr.AGEGROUPID IN ({placeholders})
              AND sr.SWIMTIME IS NOT NULL AND sr.SWIMTIME > 0
        """, [prelim_eid] + ag_ids)
        rows = c.fetchall()
        if not rows:
            continue

        c.execute("SELECT EVENTNUMBER FROM SWIMEVENT WHERE SWIMEVENTID=?",
                  [prelim_eid])
        prelim_enum = c.fetchone()[0]

        event_count = 0
        for sr_id, ath_id, ag_id, swimtime, reaction, status, entrytime, birthdate in rows:
            sr_id, ath_id = int(sr_id), int(ath_id)
            swimtime = int(swimtime)

            athlete_age = age_at(birthdate)
            if athlete_age is None:
                continue
            target_agid = None
            for fagid, famin, famax in final_brackets:
                hi = 10**9 if (famax < 0 or famax >= 99) else famax
                if famin <= athlete_age <= hi:
                    target_agid = fagid
                    break
            if target_agid is None:
                continue

            transfers_by_final.setdefault(final_eid, []).append(
                (ath_id, target_agid, swimtime, reaction, status, entrytime, sr_id))
            prelim_rows_to_delete.append(sr_id)
            event_count += 1

        if event_count:
            print(f"  prelim #{prelim_enum} → Masters final #{final_info['enum']}: "
                  f"{event_count} athlete(s)")

    if not transfers_by_final:
        print("No Masters prelim results with SWIMTIME found.")
        conn.close()
        return

    # Delete prelim rows FIRST
    for sr_id in prelim_rows_to_delete:
        c.execute("DELETE FROM SWIMRESULT WHERE SWIMRESULTID=?", [sr_id])

    # Now process each final event: create heats and assign lanes
    # Get pool lane configuration from the session
    c.execute("SELECT LANEMIN, LANEMAX FROM SWIMSESSION FETCH FIRST 1 ROWS ONLY")
    row = c.fetchone()
    lane_min = int(row[0]) if row and row[0] else 1
    lane_max = int(row[1]) if row and row[1] else 8
    lanes_per_heat = lane_max - lane_min + 1

    total_transferred = 0
    total_heats = 0

    for final_eid, athletes_data in transfers_by_final.items():
        # Get existing max heat number for this event
        c.execute("SELECT MAX(HEATNUMBER) FROM HEAT WHERE SWIMEVENTID=?",
                  [final_eid])
        row = c.fetchone()
        max_heat = int(row[0]) if row[0] else 0

        # Create heats (max lanes_per_heat athletes per heat)
        n_heats_needed = math.ceil(len(athletes_data) / lanes_per_heat)

        # Sort athletes by swimtime (fastest first) for seeding
        athletes_data.sort(key=lambda x: x[2])

        lane_idx = 0
        current_heat_id = None
        current_heat_num = max_heat

        for ath_id, target_agid, swimtime, reaction, status, entrytime, prelim_srid in athletes_data:
            # Need a new heat?
            if lane_idx % lanes_per_heat == 0:
                current_heat_num += 1
                current_heat_id = next_uid; next_uid += 1
                c.execute("""INSERT INTO HEAT
                    (HEATID, SWIMEVENTID, HEATNUMBER, SORTCODE,
                     AGEGROUPID, AGEGROUPORDER, RACESTATUS)
                    VALUES (?, ?, ?, ?, 0, 0, 2)""",
                    [current_heat_id, final_eid, current_heat_num, current_heat_num])
                total_heats += 1
                lane_idx = 0

            lane = lane_min + lane_idx
            lane_idx += 1

            # Check if athlete already has a result on this final
            c.execute("""SELECT SWIMRESULTID FROM SWIMRESULT
                WHERE ATHLETEID=? AND SWIMEVENTID=?""", [ath_id, final_eid])
            existing = c.fetchone()

            if existing:
                c.execute("""UPDATE SWIMRESULT SET SWIMTIME=?, REACTIONTIME=?,
                    RESULTSTATUS=?, HEATID=?, LANE=?
                    WHERE SWIMRESULTID=?""",
                    [swimtime,
                     reaction if reaction is not None else -32768,
                     status if status is not None else 0,
                     current_heat_id, lane, int(existing[0])])
            else:
                new_id = next_uid; next_uid += 1
                c.execute("""INSERT INTO SWIMRESULT
                    (SWIMRESULTID, ATHLETEID, SWIMEVENTID, AGEGROUPID,
                     ENTRYTIME, SWIMTIME, REACTIONTIME, RESULTSTATUS,
                     HEATID, LANE,
                     ENTRYCOURSE, BONUSENTRY, DSQNOTIFIED, FINALFIX,
                     LATEENTRY, NOADVANCE, BACKUPTIME1, BACKUPTIME2,
                     BACKUPTIME3, FINISHJUDGE, PADTIME, QTCOURSE,
                     QTTIME, QTTIMING)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                            0, 'F', 'F', 'F', 'F', 'F', 0, 0, 0, 0, ?, 0, ?, 0)""",
                    [new_id, ath_id, final_eid, target_agid,
                     int(entrytime) if entrytime else swimtime,
                     swimtime,
                     reaction if reaction is not None else -32768,
                     status if status is not None else 0,
                     current_heat_id, lane,
                     INT_MAX, INT_MAX])

            total_transferred += 1

    # Update UID counter
    c.execute("UPDATE BSUIDTABLE SET LASTUID=? WHERE NAME='BS_GLOBAL_UID'",
              [next_uid - 1])

    print(f"\n{'='*60}")
    print(f"  Summary")
    print(f"{'='*60}")
    print(f"  {total_transferred} athlete(s) moved to Masters finals")
    print(f"  {total_heats} heat(s) created")
    print(f"  {len(prelim_rows_to_delete)} prelim row(s) deleted")

    if args.dry_run:
        conn.rollback()
        print(f"\n  [DRY-RUN] — rolled back, no changes written.")
    else:
        print(f"\n  Changes committed.")

    conn.close()


if __name__ == "__main__":
    main()
