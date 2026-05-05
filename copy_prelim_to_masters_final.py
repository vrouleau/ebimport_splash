#!/usr/bin/env python3
"""Post-meet script: copy SWIMTIME from Masters athletes on non-Masters
prelim events to their corresponding Masters timed-final event, then
delete the prelim SWIMRESULT row so they don't appear as DQ/NoShow in
prelim results.

Usage:
    python copy_prelim_to_masters_final.py --mdb meet.mdb [--dry-run]

Behaviour:
  1. Scan all non-Masters prelim events (ROUND=2, MASTERS='F') for
     age-group brackets with AGEMIN in [25..99] (Masters-style).
  2. For each such event, find the matching Masters timed-final event
     (same SWIMSTYLEID, same GENDER, MASTERS='T', ROUND=1).
  3. For each SWIMRESULT on the prelim in a Masters bracket that has a
     non-NULL SWIMTIME:
       a. Find (or create) the SWIMRESULT on the Masters final for the
          same athlete, in the matching 5-year bracket.
       b. Copy SWIMTIME (and REACTIONTIME, STATUS if present).
       c. Delete the prelim SWIMRESULT row.
  4. Commit (or rollback if --dry-run).

Idempotent: rows with NULL SWIMTIME on the prelim are skipped.
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
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
    # Deduplicate (flat layout vs nested layout)
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
    conn.jconn.setAutoCommit(False)
    c = conn.cursor()

    # --- Build mapping: prelim events with Masters brackets ---
    # (prelim_event_id, SWIMSTYLEID, GENDER) -> list of Masters age brackets
    c.execute("""
        SELECT e.SWIMEVENTID, e.SWIMSTYLEID, e.GENDER,
               a.AGEGROUPID, a.AGEMIN, a.AGEMAX
        FROM SWIMEVENT e
        INNER JOIN AGEGROUP a ON a.SWIMEVENTID = e.SWIMEVENTID
        WHERE e.ROUND = 2 AND e.MASTERS = 'F'
          AND a.AGEMIN >= 25 AND a.AGEMIN < 100
    """)
    prelim_masters_brackets = {}  # (prelim_eid) -> [(agid, amin, amax)]
    prelim_meta = {}  # prelim_eid -> (styid, gender)
    for eid, styid, gen, agid, amin, amax in c.fetchall():
        eid, styid, gen = int(eid), int(styid), int(gen)
        agid, amin, amax = int(agid), int(amin), int(amax)
        prelim_masters_brackets.setdefault(eid, []).append((agid, amin, amax))
        prelim_meta[eid] = (styid, gen)

    if not prelim_masters_brackets:
        print("No non-Masters prelim events with Masters brackets found.")
        conn.close()
        return

    # --- Build mapping: Masters timed-final events ---
    # (SWIMSTYLEID, GENDER) -> {event_id, agegroups: [(agid, amin, amax)]}
    c.execute("""
        SELECT e.SWIMEVENTID, e.SWIMSTYLEID, e.GENDER, e.EVENTNUMBER,
               a.AGEGROUPID, a.AGEMIN, a.AGEMAX
        FROM SWIMEVENT e
        INNER JOIN AGEGROUP a ON a.SWIMEVENTID = e.SWIMEVENTID
        WHERE e.ROUND = 1 AND e.MASTERS = 'T'
          AND a.AGEMIN >= 25 AND a.AGEMIN < 100
    """)
    masters_finals = {}  # (styid, gender) -> {eid, brackets: [(agid, amin, amax)]}
    for eid, styid, gen, enum, agid, amin, amax in c.fetchall():
        eid, styid, gen = int(eid), int(styid), int(gen)
        agid, amin, amax = int(agid), int(amin), int(amax)
        key = (styid, gen)
        if key not in masters_finals:
            masters_finals[key] = {"eid": eid, "enum": enum, "brackets": []}
        masters_finals[key]["brackets"].append((agid, amin, amax))

    # --- Get BS_GLOBAL_UID for new row allocation ---
    c.execute("SELECT LASTUID FROM BSUIDTABLE WHERE NAME='BS_GLOBAL_UID'")
    next_uid = int(c.fetchone()[0]) + 1

    # --- Process each prelim event ---
    total_copied = 0
    total_created = 0

    for prelim_eid, brackets in prelim_masters_brackets.items():
        styid, gender = prelim_meta[prelim_eid]
        final_info = masters_finals.get((styid, gender))
        if final_info is None:
            print(f"  [SKIP] prelim eid={prelim_eid}: no Masters final "
                  f"for (styid={styid}, gender={gender})")
            continue

        final_eid = final_info["eid"]
        final_brackets = final_info["brackets"]

        # Get all SWIMRESULT rows on this prelim in Masters brackets
        ag_ids = [b[0] for b in brackets]
        placeholders = ",".join("?" * len(ag_ids))
        c.execute(f"""
            SELECT sr.SWIMRESULTID, sr.ATHLETEID, sr.AGEGROUPID,
                   sr.SWIMTIME, sr.REACTIONTIME, sr.RESULTSTATUS,
                   ath.BIRTHDATE
            FROM SWIMRESULT sr
            INNER JOIN ATHLETE ath ON ath.ATHLETEID = sr.ATHLETEID
            WHERE sr.SWIMEVENTID = ? AND sr.AGEGROUPID IN ({placeholders})
              AND sr.SWIMTIME IS NOT NULL AND sr.SWIMTIME > 0
        """, [prelim_eid] + ag_ids)
        rows = c.fetchall()

        if not rows:
            continue

        # Get event number for reporting
        c.execute("SELECT EVENTNUMBER FROM SWIMEVENT WHERE SWIMEVENTID=?",
                  [prelim_eid])
        prelim_enum = c.fetchone()[0]

        for sr_id, ath_id, ag_id, swimtime, reaction, status, birthdate in rows:
            sr_id, ath_id = int(sr_id), int(ath_id)
            swimtime = int(swimtime) if swimtime is not None else None

            # Determine athlete age → find matching final bracket
            athlete_age = age_at(birthdate)
            if athlete_age is None:
                continue

            target_agid = None
            for fagid, famin, famax in final_brackets:
                hi = 10**9 if famax < 0 else famax
                if famin <= athlete_age <= hi:
                    target_agid = fagid
                    break
            if target_agid is None:
                continue

            # Check if athlete already has a result on the Masters final
            c.execute("""
                SELECT SWIMRESULTID, SWIMTIME FROM SWIMRESULT
                WHERE ATHLETEID=? AND SWIMEVENTID=?
            """, [ath_id, final_eid])
            existing = c.fetchone()

            if existing:
                existing_srid = int(existing[0])
                # Update with the prelim time
                c.execute("""
                    UPDATE SWIMRESULT SET SWIMTIME=?, REACTIONTIME=?,
                           RESULTSTATUS=?
                    WHERE SWIMRESULTID=?
                """, [swimtime,
                      reaction if reaction is not None else -32768,
                      status if status is not None else 0,
                      existing_srid])
            else:
                # Create a new SWIMRESULT on the Masters final
                new_id = next_uid
                next_uid += 1
                c.execute("""
                    INSERT INTO SWIMRESULT
                    (SWIMRESULTID, ATHLETEID, SWIMEVENTID, AGEGROUPID,
                     ENTRYTIME, SWIMTIME, REACTIONTIME, RESULTSTATUS,
                     ENTRYCOURSE, BONUSENTRY, DSQNOTIFIED, FINALFIX,
                     LATEENTRY, NOADVANCE, BACKUPTIME1, BACKUPTIME2,
                     BACKUPTIME3, FINISHJUDGE, PADTIME, QTCOURSE,
                     QTTIME, QTTIMING)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, 0, 'F', 'F', 'F',
                            'F', 'F', 0, 0, 0, 0, ?, 0, ?, 0)
                """, [new_id, ath_id, final_eid, target_agid,
                      swimtime, swimtime,
                      reaction if reaction is not None else -32768,
                      status if status is not None else 0,
                      INT_MAX, INT_MAX])
                total_created += 1

            # Delete the prelim SWIMRESULT row entirely so the athlete
            # doesn't appear as DQ/NoShow in prelim results.
            c.execute("DELETE FROM SWIMRESULT WHERE SWIMRESULTID=?", [sr_id])
            total_copied += 1

        print(f"  event #{prelim_enum} (prelim eid={prelim_eid}) → "
              f"Masters final eid={final_eid}: "
              f"{len(rows)} time(s) transferred")

    # Update BS_GLOBAL_UID if we allocated new IDs
    if total_created > 0:
        c.execute("UPDATE BSUIDTABLE SET LASTUID=? WHERE NAME='BS_GLOBAL_UID'",
                  [next_uid - 1])

    # Summary
    print(f"\n{'=' * 60}")
    print(f"  Summary")
    print(f"{'=' * 60}")
    print(f"  {total_copied} prelim time(s) copied to Masters final")
    print(f"  {total_created} new SWIMRESULT row(s) created on Masters final")
    print(f"  {total_copied} prelim row(s) deleted")

    if args.dry_run:
        conn.rollback()
        print("\n  [DRY-RUN] — rolled back, no changes written.")
    else:
        conn.commit()
        print("\n  Changes committed.")

    conn.close()


if __name__ == "__main__":
    main()
