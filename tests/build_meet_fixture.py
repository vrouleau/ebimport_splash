"""Augment the committed Gatineau meet .lxf with synthetic Masters events
and the missing mixed-relay UIDs (541-544) so the integration tests cover
the full validator surface.

Idempotent: re-running rebuilds the fixture from scratch using the original
Gatineau .lef as input. Run from the repo root:

    python tests/build_meet_fixture.py

Writes:
    tests/fixtures/meet_template.lxf    (overwritten)
"""
from __future__ import annotations

import re
import shutil
import zipfile
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET

REPO = Path(__file__).resolve().parent.parent
SRC_LXF = REPO / "tests" / "fixtures" / "meet_template_base.lxf"
DST_LXF = REPO / "tests" / "fixtures" / "meet_template.lxf"

# 5-year Masters brackets used everywhere a "Masters AGEGROUPs" requirement applies.
MASTERS_BRACKETS = [(a, a + 4) for a in range(25, 100, 5)]

# Synthetic SWIMSTYLE descriptors for UIDs that may already (UID 502 etc) or may
# not (UID 541-544) live in the original template. Distance/relaycount values
# come from src/core.py docstring.
STYLE_INFO = {
    502: ("100 m Portage Mannequin palmes",       100, 1),
    504: ("12 m Lancer de la corde",               12, 2),
    506: ("100 m Remorquage mannequin palmes",    100, 1),
    507: ("50 m Portage du mannequin plein",       50, 1),
    508: ("200 m Sauveteur d'acier",              200, 1),
    531: ("100 m Sauvetage combine",              100, 1),
    541: ("100 m Nage avec obstacles (Masters)",  100, 1),
    542: ("4 x 50 m Relais obstacle mixte",        50, 4),
    543: ("2 x 50 m Relais mixte portage",         50, 2),
    544: ("4 x 50 m Relais mixte sauve combine",   50, 4),
}

# What we need to add to the template, derived from the FATAL list run_validation
# emits when the unmodified Gatineau template is fed test_attendees.xlsx:
#   - Masters individuals: UIDs 502, 506, 507, 508, 531, 541 for M and F
#   - Masters relay:       UID 504 for M and F
#   - Mixed relays (gender=X), UIDs 542, 543, 544 — both 15-18/Open and Masters
NEW_EVENTS: list[tuple[int, str, str, list[tuple[int, int]]]] = []
for uid in (502, 506, 507, 508, 531, 541):
    for g in ("M", "F"):
        NEW_EVENTS.append((uid, g, "MASTERS", MASTERS_BRACKETS))
for g in ("M", "F"):
    NEW_EVENTS.append((504, g, "MASTERS", MASTERS_BRACKETS))
for uid in (542, 543, 544):
    NEW_EVENTS.append((uid, "X", "",        [(15, 18), (19, -1)]))
    NEW_EVENTS.append((uid, "X", "MASTERS", MASTERS_BRACKETS))


def build() -> bytes:
    with zipfile.ZipFile(SRC_LXF) as z:
        lef_name = next(n for n in z.namelist() if n.endswith(".lef"))
        xml = z.read(lef_name).decode("utf-8")

    # ElementTree drops the XML decl and any DOCTYPE; preserve the original
    # by patching the bytes after parsing to keep the .lef byte-clean.
    root = ET.fromstring(xml)

    # Find current max ids so we don't collide.
    next_event_id = max(
        (int(e.get("eventid", 0)) for e in root.iter("EVENT")),
        default=10_000,
    ) + 1
    next_agegroup_id = max(
        (int(a.get("agegroupid", 0)) for a in root.iter("AGEGROUP")),
        default=10_000,
    ) + 1
    next_event_number = max(
        (int(e.get("number", 0)) for e in root.iter("EVENT")),
        default=100,
    ) + 1

    # Pick the last SESSION as the host for the synthetic events.
    sessions = list(root.iter("SESSION"))
    if not sessions:
        raise RuntimeError("template has no SESSION element")
    target_session = sessions[-1]
    events_el = target_session.find("EVENTS")
    if events_el is None:
        events_el = ET.SubElement(target_session, "EVENTS")

    for uid, gender, ev_type, brackets in NEW_EVENTS:
        name, distance, relaycount = STYLE_INFO[uid]
        ev_attrs = {
            "eventid": str(next_event_id),
            "number": str(next_event_number),
            "order": str(next_event_number),
            "round": "TIM",
            "status": "OFFICIAL",
            "preveventid": "-1",
            "gender": gender,
        }
        if ev_type:
            ev_attrs["type"] = ev_type
        ev_el = ET.SubElement(events_el, "EVENT", ev_attrs)
        ET.SubElement(ev_el, "SWIMSTYLE", {
            "swimstyleid": str(uid),
            "name": name,
            "distance": str(distance),
            "relaycount": str(relaycount),
            "stroke": "UNKNOWN",
        })
        ags_el = ET.SubElement(ev_el, "AGEGROUPS")
        for amin, amax in brackets:
            ET.SubElement(ags_el, "AGEGROUP", {
                "agegroupid": str(next_agegroup_id),
                "agemin": str(amin),
                "agemax": str(amax),
            })
            next_agegroup_id += 1
        next_event_id += 1
        next_event_number += 1

    # Re-serialize. Preserve XML decl from the original.
    decl_match = re.match(rb"^\s*<\?xml[^>]*\?>\s*", xml.encode("utf-8"))
    decl = decl_match.group(0) if decl_match else b'<?xml version="1.0" encoding="UTF-8"?>\n'
    body = ET.tostring(root, encoding="utf-8")
    out_xml = decl + body

    # Repackage as .lxf (zip with the original .lef name preserved).
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        zout.writestr(lef_name, out_xml)
    return buf.getvalue()


def main() -> None:
    new_bytes = build()
    DST_LXF.write_bytes(new_bytes)
    print(f"Wrote {DST_LXF} ({len(new_bytes)} bytes, {len(NEW_EVENTS)} events added)")


if __name__ == "__main__":
    main()
