#!/usr/bin/env python3
"""Format audit JSON output into a readable summary."""
import json
import sys

def main():
    path = sys.argv[1] if len(sys.argv) > 1 else None
    if path:
        with open(path) as f:
            data = json.load(f)
    else:
        data = json.load(sys.stdin)

    if "error" in data:
        print(f"ERROR: {data['error']}")
        sys.exit(1)

    print("=" * 60)
    print("  AUDIT RESULTS")
    print("=" * 60)
    print(f"  Athletes in PDF:  {data['pdf_athletes']}")
    print(f"  Entries in PDF:   {data['pdf_entries']}")
    print(f"  Events in PDF:    {len(data['pdf_events'])}")
    print("-" * 60)

    checks = data["checks"]
    all_ok = True

    # Duplicates
    c = checks["duplicates"]
    status = "PASS" if c["ok"] else "FAIL"
    if not c["ok"]: all_ok = False
    print(f"\n  [{status}] Duplicates: {c['count']}")
    for d in c.get("details", [])[:10]:
        print(f"         {d}")

    # Missing from PDF
    c = checks["missing_from_pdf"]
    status = "PASS" if c["ok"] else "WARN"
    if not c["ok"] and c["count"] > 20: all_ok = False
    print(f"\n  [{status}] Missing from PDF: {c['count']}")
    for d in c.get("details", [])[:15]:
        print(f"         {d}")
    if c["count"] > 15:
        print(f"         ... and {c['count'] - 15} more")

    # Time accuracy
    c = checks["time_accuracy"]
    status = "PASS" if c["ok"] else "FAIL"
    if not c["ok"]: all_ok = False
    print(f"\n  [{status}] Time accuracy (±5%): {c['mismatches']} mismatches")
    for d in c.get("details", [])[:10]:
        print(f"         {d}")
    if c["mismatches"] > 10:
        print(f"         ... and {c['mismatches'] - 10} more")

    # NT placement
    c = checks["nt_placement"]
    status = "PASS" if c["ok"] else "WARN"
    print(f"\n  [{status}] NT placement: {c['total_nt']} NT entries")
    if not c["ok"]:
        print(f"         {c['in_heat_gt2']} in heat > 2 (should be in slowest heats)")

    # Seeding order
    c = checks["seeding_order"]
    status = "PASS" if c["ok"] else "WARN"
    print(f"\n  [{status}] Seeding order: {c['violations']} violations")
    for d in c.get("details", [])[:5]:
        print(f"         {d}")

    print("\n" + "=" * 60)
    if data["all_critical_ok"]:
        print("  ALL CRITICAL CHECKS PASSED")
    else:
        print("  ISSUES FOUND - review above")
    print("=" * 60)


if __name__ == "__main__":
    main()
