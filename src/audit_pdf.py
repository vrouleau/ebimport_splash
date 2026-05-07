"""
Audit a SPLASH heat-sheet PDF against the source xlsx.

Checks:
  - NT athletes are in the slowest heats (heat 1-2)
  - No duplicate entries (same athlete, same event)
  - All xlsx individual athletes appear in the PDF
  - Entry times in PDF match xlsx
  - Seeding order within age brackets is correct

Usage:
    python audit_pdf.py --pdf HEATS.pdf --xlsx ATTENDEES.xlsx
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from collections import defaultdict
from pathlib import Path

import fitz  # PyMuPDF

# Allow importing from same directory
sys.path.insert(0, str(Path(__file__).parent))
from load_to_mdb import read_attendees, IssueCollector


# --------------------------------------------------------------------------- #
# PDF parsing
# --------------------------------------------------------------------------- #
_RE_EVENT = re.compile(r"Epreuve\s+(\d+)")
_RE_HEAT = re.compile(r"Série\s+(\d+)\s+de\s+(\d+)")
_RE_ATHLETE_HEAT = re.compile(
    r"^(\d+)\s+([A-ZÀÂÄÉÈÊËÏÎÔÙÛÜŸÇÆŒ'' -]+),\s+(.+)$"
)
_RE_RANK = re.compile(r"^(\d+)\.$|^disq\.$|^dns$|^dnf$")
_RE_NAME = re.compile(
    r"^([A-ZÀÂÄÉÈÊËÏÎÔÙÛÜŸÇÆŒ'' -]+),\s+(.+)$"
)
_RE_TIME = re.compile(r"^(\d+:)*\d+[.:]\d+$")


def parse_pdf(pdf_path: Path) -> list[dict]:
    """Extract athlete entries from a SPLASH heat-sheet or results PDF."""
    doc = fitz.open(str(pdf_path))
    entries = []
    event_num = None
    heat_num = None
    total_heats = None
    is_results = False

    for page in doc:
        lines = page.get_text().splitlines()
        i = 0

        # Detect results format
        if any("Liste résultats" in l for l in lines):
            is_results = True

        while i < len(lines):
            line = lines[i].strip()

            m = _RE_EVENT.match(line)
            if m:
                event_num = int(m.group(1))
                i += 1
                continue

            m = _RE_HEAT.match(line)
            if m:
                heat_num = int(m.group(1))
                total_heats = int(m.group(2))
                i += 1
                continue

            # Heat-sheet format: "lane  LASTNAME, First"
            m = _RE_ATHLETE_HEAT.match(line)
            if m and event_num and not is_results:
                lane = int(m.group(1))
                last = m.group(2).strip()
                first = m.group(3).strip()
                birth_year = lines[i + 1].strip() if i + 1 < len(lines) else ""
                club = lines[i + 2].strip() if i + 2 < len(lines) else ""
                time_str = lines[i + 3].strip() if i + 3 < len(lines) else ""
                if not (_RE_TIME.match(time_str) or time_str == "NT"):
                    time_str = None
                entries.append({
                    "last": last, "first": first, "lane": lane,
                    "event": event_num, "time": time_str,
                    "club": club, "birth_year": birth_year,
                    "heat": heat_num, "total_heats": total_heats,
                })
                i += 4
                continue

            # Results format: "rank." then "LASTNAME, First" on next line
            m = _RE_RANK.match(line)
            if m and event_num:
                # Peek at next line for name
                if i + 1 < len(lines):
                    nm = _RE_NAME.match(lines[i + 1].strip())
                    if nm:
                        last = nm.group(1).strip()
                        first = nm.group(2).strip()
                        # birth_year (may be missing, e.g. "NODOB, Nora" has no year)
                        j = i + 2
                        birth_year = ""
                        club = ""
                        time_str = None
                        if j < len(lines):
                            val = lines[j].strip()
                            if re.match(r"^\d{2,4}$", val):
                                birth_year = val
                                j += 1
                            # club
                            if j < len(lines):
                                val = lines[j].strip()
                                if val and not _RE_TIME.match(val) and val not in ("A", "B", "R", "disq."):
                                    club = val
                                    j += 1
                            # time
                            if j < len(lines):
                                val = lines[j].strip()
                                if _RE_TIME.match(val):
                                    time_str = val
                                    j += 1
                        entries.append({
                            "last": last, "first": first, "lane": None,
                            "event": event_num, "time": time_str,
                            "club": club, "birth_year": birth_year,
                            "heat": None, "total_heats": None,
                        })
                        i = j
                        continue
            i += 1
    doc.close()
    return entries


def parse_time_ms(t: str | None) -> int | None:
    """Parse a PDF time string to milliseconds."""
    if not t or t == "NT":
        return None
    # H:MM:SS.hh
    m = re.match(r"(\d+):(\d+):(\d+)\.(\d+)", t)
    if m:
        return (int(m.group(1)) * 3600000 + int(m.group(2)) * 60000
                + int(m.group(3)) * 1000 + int(m.group(4)) * 10)
    # M:SS.hh
    m = re.match(r"(\d+):(\d+)\.(\d+)", t)
    if m:
        return (int(m.group(1)) * 60000 + int(m.group(2)) * 1000
                + int(m.group(3)) * 10)
    # SS.hh
    m = re.match(r"(\d+)\.(\d+)", t)
    if m:
        return int(m.group(1)) * 1000 + int(m.group(2)) * 10
    return None


# --------------------------------------------------------------------------- #
# Audit logic
# --------------------------------------------------------------------------- #
def audit(pdf_path: Path, xlsx_path: Path) -> dict:
    """Run all checks. Returns a dict with results."""
    entries = parse_pdf(pdf_path)

    issues_coll = IssueCollector()
    inscriptions = read_attendees(xlsx_path, issues_coll)

    # Build xlsx lookup
    xlsx_by_name: dict[tuple, list] = defaultdict(list)
    for ins in inscriptions:
        if not ins.event.is_relay:
            key = (ins.last.upper().strip(), ins.first.strip())
            xlsx_by_name[key].append(ins)

    results: dict = {
        "pdf_entries": len(entries),
        "pdf_athletes": len(set((e["last"], e["first"]) for e in entries)),
        "pdf_events": sorted(set(e["event"] for e in entries)),
        "checks": {},
    }

    # --- Check 1: NT placement ---
    nt_entries = [e for e in entries if e["time"] == "NT"]
    nt_in_high = [e for e in nt_entries if e["heat"] and e["heat"] > 2]
    results["checks"]["nt_placement"] = {
        "total_nt": len(nt_entries),
        "in_heat_1_2": len(nt_entries) - len(nt_in_high),
        "in_heat_gt2": len(nt_in_high),
        "details": [
            f"{e['first']} {e['last']} ev#{e['event']} heat {e['heat']}/{e['total_heats']}"
            for e in nt_in_high
        ],
        "ok": len(nt_in_high) == 0,
    }

    # --- Check 2: Duplicates ---
    counts = defaultdict(int)
    for e in entries:
        counts[(e["last"], e["first"], e["event"])] += 1
    dupes = {k: v for k, v in counts.items() if v > 1}
    results["checks"]["duplicates"] = {
        "count": len(dupes),
        "details": [
            f"{k[1]} {k[0]} ev#{k[2]}: {v}x" for k, v in dupes.items()
        ],
        "ok": len(dupes) == 0,
    }

    # --- Check 3: Missing from PDF ---
    pdf_names = set((e["last"], e["first"]) for e in entries)
    xlsx_names = set(xlsx_by_name.keys())
    missing = xlsx_names - pdf_names
    results["checks"]["missing_from_pdf"] = {
        "count": len(missing),
        "details": [f"{first} {last}" for last, first in sorted(missing)],
        "ok": len(missing) == 0,
    }

    # --- Check 4: Time accuracy ---
    xlsx_times_by_name: dict[tuple, set] = {}
    for key, ins_list in xlsx_by_name.items():
        xlsx_times_by_name[key] = {
            ins.best_time_ms for ins in ins_list if ins.best_time_ms is not None
        }

    time_mismatches = []
    for e in entries:
        pdf_ms = parse_time_ms(e["time"])
        if pdf_ms is None:
            continue
        key = (e["last"], e["first"])
        xlsx_times = xlsx_times_by_name.get(key, set())
        if xlsx_times and pdf_ms not in xlsx_times:
            # Allow ±5% tolerance (simulated results vary from entry time)
            if not any(abs(pdf_ms - t) <= t * 0.05 for t in xlsx_times):
                time_mismatches.append(
                    f"{e['first']} {e['last']} ev#{e['event']}: "
                    f"PDF={e['time']} xlsx={sorted(xlsx_times)}"
                )
    results["checks"]["time_accuracy"] = {
        "mismatches": len(time_mismatches),
        "details": time_mismatches[:20],
        "ok": len(time_mismatches) == 0,
    }

    # --- Check 5: Seeding order ---
    by_event_heat: dict[int, dict[int, list]] = defaultdict(lambda: defaultdict(list))
    for e in entries:
        by_event_heat[e["event"]][e["heat"]].append(e)

    seeding_violations = []
    for ev_num in sorted(by_event_heat):
        heats = by_event_heat[ev_num]
        for h in sorted(h for h in heats if h is not None):
            if h + 1 not in heats:
                continue
            t_h = [parse_time_ms(e["time"]) for e in heats[h]]
            t_h = [t for t in t_h if t is not None]
            t_next = [parse_time_ms(e["time"]) for e in heats[h + 1]]
            t_next = [t for t in t_next if t is not None]
            if t_h and t_next and min(t_h) < max(t_next):
                seeding_violations.append(
                    f"ev#{ev_num} heat {h} fastest={min(t_h)}ms "
                    f"< heat {h+1} slowest={max(t_next)}ms"
                )
    results["checks"]["seeding_order"] = {
        "violations": len(seeding_violations),
        "note": "Violations at age-bracket boundaries are normal",
        "details": seeding_violations,
        "ok": len(seeding_violations) == 0,
    }

    # --- Overall ---
    critical = (
        results["checks"]["duplicates"]["ok"]
        and results["checks"]["missing_from_pdf"]["ok"]
        and results["checks"]["time_accuracy"]["ok"]
    )
    results["all_critical_ok"] = critical

    return results


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--pdf", required=True, type=Path)
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--json", action="store_true", help="Output as JSON")
    args = ap.parse_args()

    results = audit(args.pdf, args.xlsx)

    if args.json:
        print(json.dumps(results, indent=2, ensure_ascii=False))
    else:
        print(f"PDF: {results['pdf_entries']} entries, "
              f"{results['pdf_athletes']} athletes, "
              f"events {results['pdf_events']}")
        print()
        for name, check in results["checks"].items():
            status = "✓" if check["ok"] else "✗"
            print(f"  {status} {name}")
            if not check["ok"] and check.get("details"):
                for d in check["details"][:10]:
                    print(f"      {d}")
                if len(check.get("details", [])) > 10:
                    print(f"      … and {len(check['details']) - 10} more")
        print()
        if results["all_critical_ok"]:
            print("All critical checks passed.")
        else:
            print("CRITICAL ISSUES FOUND — review above.")


if __name__ == "__main__":
    main()
