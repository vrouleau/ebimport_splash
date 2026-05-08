"""Parse a SPLASH meet export .lxf into event structure.

Used by both ebimport_splash and meetmanager-app to get event IDs,
agegroups, and swimstyles from the authoritative SPLASH export.
"""
from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from xml.etree import ElementTree as ET


@dataclass
class MeetAgeGroup:
    agegroupid: int
    agemin: int
    agemax: int


@dataclass
class MeetEvent:
    eventid: int
    number: int
    gender: str  # "F", "M", "X"
    round: str  # "TIM", "PRE", "FIN"
    event_type: str  # "MASTERS" or ""
    swimstyleid: int
    distance: int
    relaycount: int
    style_name: str
    agegroups: list[MeetAgeGroup] = field(default_factory=list)

    @property
    def is_masters(self) -> bool:
        return self.event_type == "MASTERS"

    @property
    def is_prelim(self) -> bool:
        return self.round == "PRE"

    @property
    def is_final(self) -> bool:
        return self.round in ("TIM", "FIN")

    @property
    def gender_int(self) -> int:
        return {"M": 1, "F": 2, "X": 3}.get(self.gender, 0)


@dataclass
class MeetSession:
    number: int
    name: str
    events: list[MeetEvent] = field(default_factory=list)


@dataclass
class ParsedMeet:
    sessions: list[MeetSession] = field(default_factory=list)

    @property
    def all_events(self) -> list[MeetEvent]:
        return [e for s in self.sessions for e in s.events]

    def find_event(self, swimstyleid: int, gender_int: int, masters: bool = False) -> MeetEvent | None:
        """Find event by style UID + gender + masters flag. Prefer prelim for non-masters."""
        gender_str = {1: "M", 2: "F", 3: "X"}.get(gender_int, "")
        candidates = [e for e in self.all_events
                      if e.swimstyleid == swimstyleid and e.gender == gender_str
                      and e.is_masters == masters]
        # Prefer prelim
        for e in candidates:
            if e.is_prelim:
                return e
        return candidates[0] if candidates else None

    def find_event_any(self, swimstyleid: int, gender_int: int) -> MeetEvent | None:
        """Find any event for this style+gender (fallback)."""
        gender_str = {1: "M", 2: "F", 3: "X"}.get(gender_int, "")
        candidates = [e for e in self.all_events
                      if e.swimstyleid == swimstyleid and e.gender == gender_str]
        return candidates[0] if candidates else None


def parse_meet_lxf(source) -> ParsedMeet:
    """Parse a meet .lxf (path, bytes, or file-like) into ParsedMeet.

    Accepts: Path, str (file path), bytes, or BytesIO.
    """
    if isinstance(source, (str, Path)):
        with open(source, "rb") as f:
            raw = f.read()
    elif isinstance(source, bytes):
        raw = source
    else:
        raw = source.read()

    # Unzip
    with zipfile.ZipFile(BytesIO(raw)) as z:
        lef_name = [n for n in z.namelist() if n.endswith(".lef")][0]
        xml_bytes = z.read(lef_name)

    root = ET.fromstring(xml_bytes)
    meet = ParsedMeet()

    for session_el in root.iter("SESSION"):
        ses = MeetSession(
            number=int(session_el.get("number", 0)),
            name=session_el.get("name", ""),
        )
        for event_el in session_el.iter("EVENT"):
            style_el = event_el.find("SWIMSTYLE")
            ev = MeetEvent(
                eventid=int(event_el.get("eventid", 0)),
                number=int(event_el.get("number", 0)),
                gender=event_el.get("gender", ""),
                round=event_el.get("round", "TIM"),
                event_type=event_el.get("type", ""),
                swimstyleid=int(style_el.get("swimstyleid", 0)) if style_el is not None else 0,
                distance=int(style_el.get("distance", 0)) if style_el is not None else 0,
                relaycount=int(style_el.get("relaycount", 1)) if style_el is not None else 1,
                style_name=(style_el.get("name", "") if style_el is not None else ""),
            )
            for ag_el in event_el.iter("AGEGROUP"):
                ev.agegroups.append(MeetAgeGroup(
                    agegroupid=int(ag_el.get("agegroupid", 0)),
                    agemin=int(ag_el.get("agemin", -1)),
                    agemax=int(ag_el.get("agemax", -1)),
                ))
            ses.events.append(ev)
        meet.sessions.append(ses)

    return meet
