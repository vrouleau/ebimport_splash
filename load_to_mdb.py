#!/usr/bin/env python3
"""
load_to_mdb.py
Populate a SPLASH Meet Manager 11 .mdb file with clubs, athletes and
individual / relay session inscriptions read from the "Attendees" sheet
of a registration workbook (e.g. CPLC2026FINAL.xlsx).

The target .mdb is an empty meet shell created by SPLASH Meet Manager 11.
All primary keys come from BSUIDTABLE.BS_GLOBAL_UID (monotonic counter
shared by every table). We insert:

    - custom SWIMSTYLE rows for the lifesaving strokes
      (Obstacle, Corde, Portage, Remorquage, Sauveteur d'acier,
       Medley + the three mixed relays)
    - one placeholder SWIMSESSION (user re-organises afterwards in MM)
    - one SWIMEVENT per distinct (age-group, gender, stroke, distance)
    - one AGEGROUP row per SWIMEVENT
    - one CLUB row per distinct club
    - one ATHLETE row per distinct (first name, last name, NRAN)
    - one SWIMRESULT row per individual inscription (RESULTSTATUS=0)
    - one RELAY row per (club, age group, relay style) with the
      athletes of that club as RELAYPOSITION rows

Run from Linux/WSL with Java 8+ available and UCanAccess unpacked in
/tmp/ucanaccess (or change UCANACCESS_DIR below).  On Windows you can
point pyodbc at the MS Access driver instead – see README.

Usage:
    python3 load_to_mdb.py --xlsx CPLC2026FINAL.xlsx --mdb Canadien.mdb
                           [--dry-run] [--wipe]
"""
from __future__ import annotations

import argparse
import datetime as dt
import glob
import os
import re
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import openpyxl
import jaydebeapi  # JDBC bridge over UCanAccess

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #
UCANACCESS_DIR = os.environ.get(
    "UCANACCESS_DIR", "/tmp/ucanaccess/UCanAccess-5.0.1.bin"
)

MEET_NATION = "CAN"
PLACEHOLDER_SESSION_NAME = "Placeholder – reorganise in Meet Manager"
# Session date must exist; user will edit it inside MM.
PLACEHOLDER_SESSION_DATE = dt.datetime(2026, 6, 20, 9, 0, 0)
# Age reference date (FINA "age at end of meet year" — matches sample Lenex).
# Used when choosing an AGEGROUP for an entry if MM-side age-splits exist.
AGE_DATE = dt.date(2026, 6, 20)
POOLTYPE_LCM = 1  # 1=LCM, 0=SCM, 2=SCY — matches the sample .lef "LCM"
COURSE_LCM = 1

# SPLASH/Lenex gender encoding in SMALLINT columns.
# 0 = All / Mixed, 1 = Male, 2 = Female.
# Note: the federation reference DB uses 0 (All) even for mixed-gender relays,
# never 3 (which some Lenex tools emit for "Mixed").  Sticking to 0/1/2
# avoids any chance of hitting a Delphi case-else that isn't handled.
GENDER_MALE   = 1
GENDER_FEMALE = 2
GENDER_ALL    = 0
GENDER_MIXED  = 0        # alias — treat mixed relays as "All" per SS convention

# ----------------------------------------------------------------------------
# Société de Sauvetage lifesaving catalog (extracted from a real SPLASH DB
# produced by the federation — "30-Deux 25 octobre 2025.mdb").  SPLASH stores
# lifesaving events with STROKE=0 / TECHNIQUE=0 and keeps the identity in the
# CODE, NAME and UNIQUEID columns.  STROKE=0 tells TBSwLanguage.StrokeName to
# render the CODE/NAME instead of looking up one of the swim strokes (1..5).
#
# UNIQUEIDs 501-540 come straight from the federation catalog.  550-552 are
# reserved here for events in the Canadien meet that don't exist in the
# original catalog (Medley 200 m, Portage relay 4x50 m, Obstacle 100 m for
# Masters).  Bilingual French / English names are emitted in both
# SWIMSTYLE.NAME (primary, French) and the Lenex NAME attribute.
# ----------------------------------------------------------------------------
LIFESAVING_STROKE    = 0        # STROKE value used by SPLASH for catalog items
LIFESAVING_TECHNIQUE = 0

# Ticket-type parsing helpers
NON_RACE_PREFIXES = (
    "Banquet", "Coach", "Cosmod", "Couloir", "Officiel", "Priorit",
    "Sheraton",
)

# label -> catalog entry.  All distances agreed with the meet organiser.
#   key    : (ticket_label, is_relay)
#   value  : (UNIQUEID, CODE, distance_m, relay_count, name_fr, name_en)
LIFESAVING_CATALOG: dict[tuple, tuple] = {
    # individuals (Corde stays individual for now; pairing will be handled in MM)
    ("Corde",             False): (504, "ID504",  12, 1,
                                   "12 m Lancer de la corde",
                                   "12 m Line Throw"),
    ("Obstacle",          False): (501, "ID501", 200, 1,
                                   "200 m Nage avec obstacles",
                                   "200 m Obstacle Swim"),
    ("Obstacle100",       False): (552, "ID552", 100, 1,
                                   "100 m Nage avec obstacles",
                                   "100 m Obstacle Swim"),
    ("Portage",           False): (502, "ID502", 100, 1,
                                   "100 m Portage du mannequin plein avec palmes",
                                   "100 m Manikin Carry w/ Fins"),
    ("Portage50",         False): (507, "ID507",  50, 1,
                                   "50 m Portage du mannequin plein",
                                   "50 m Manikin Carry"),
    ("Remorquage",        False): (506, "ID506", 100, 1,
                                   "100 m Remorquage du mannequin ½ plein + palmes",
                                   "100 m Manikin Tow with Fins"),
    ("Sauveteur d'acier", False): (508, "ID508", 200, 1,
                                   "200 m Sauveteur d'acier",
                                   "200 m Super Lifesaver"),
    ("Medley",            False): (550, "ID550", 200, 1,
                                   "200 m Medley de sauvetage",
                                   "200 m Rescue Medley"),

    # relays — all 4x50 m per the organiser's spec
    ("Medley",            True):  (538, "ID538",  50, 4,
                                   "4 x 50 m Relais Medley",
                                   "4 x 50 m Rescue Medley Relay"),
    ("Obstacle",          True):  (540, "ID540",  50, 4,
                                   "4 x 50 m Relais obstacles",
                                   "4 x 50 m Obstacle Relay"),
    ("Portage",           True):  (551, "ID551",  50, 4,
                                   "4 x 50 m Relais portage du mannequin",
                                   "4 x 50 m Manikin Carry Relay"),
}

AGE_GROUPS = {  # code -> (AGEMIN, AGEMAX, display name FR/EN)
    "1518":    (15, 18, "15-18 ans"),
    "MASTERS": (30, 99, "Maîtres"),
    "OPEN":    ( 0, 99, "Toutes catégories"),
}

# Combined events ("Cumulatifs") — one per (age_code × gender).
# Each cumulatif sums points from a list of catalog UIDs.  Point schedule
# comes from the 30-Deux federation sample: 1st=20, 2nd=18, ..., 16th=1,
# 17th+=0.  Corde (UID 504) and Medley (UID 550) are intentionally excluded
# per federation convention; relays aren't counted either.
#
#   (age_code, gender) -> (display_name, [UIDs])
CUMULATIF_POINTS = "20,18,16,14,13,12,11,10,8,7,6,5,4,3,2,1"
CUMULATIFS: dict[tuple, tuple[str, list[int]]] = {
    ("1518",    GENDER_FEMALE): ("Cumulatif 15-18 ans - dames",
                                 [501, 502, 506, 507, 508]),
    ("1518",    GENDER_MALE):   ("Cumulatif 15-18 ans - hommes",
                                 [501, 502, 506, 507, 508]),
    ("OPEN",    GENDER_FEMALE): ("Cumulatif Open - dames",
                                 [501, 502, 506, 507, 508]),
    ("OPEN",    GENDER_MALE):   ("Cumulatif Open - hommes",
                                 [501, 502, 506, 507, 508]),
    # Masters uses the 100 m Obstacle variant (UID 552) instead of 200 m (501).
    ("MASTERS", GENDER_FEMALE): ("Cumulatif Maîtres - dames",
                                 [552, 502, 506, 507, 508]),
    ("MASTERS", GENDER_MALE):   ("Cumulatif Maîtres - hommes",
                                 [552, 502, 506, 507, 508]),
}


@dataclass
class EventKey:
    age_code: str          # '1518' | 'MASTERS' | 'OPEN'
    gender: int            # GENDER_MALE | GENDER_FEMALE | GENDER_MIXED
    uniqueid: int          # Société de Sauvetage catalog UID
    code: str              # 'ID501' etc.
    distance: int
    relay_count: int       # 1 for individual, 4 for relay
    name_fr: str
    name_en: str

    def key(self) -> tuple:
        # Identity of an event = (age bracket, gender, catalog UID).
        # The catalog UID already encodes distance + relay count.
        return (self.age_code, self.gender, self.uniqueid)


def parse_ticket(ticket: str) -> EventKey | None:
    """Return an EventKey if the ticket is a race, else None."""
    t = ticket.strip()
    for p in NON_RACE_PREFIXES:
        if t.startswith(p):
            return None
    m = re.match(r"^(15-18|MA|Open)\s+(.*)$", t)
    if not m:
        return None
    age_code = {"15-18": "1518", "MA": "MASTERS", "Open": "OPEN"}[m.group(1)]
    rest = m.group(2).strip()

    # Relay?
    mr = re.match(r"^Relais Mixte\s+(\S+)", rest)
    if mr:
        style = mr.group(1).strip()   # 'Medley' | 'Obstacle' | 'Portage'
        cat = LIFESAVING_CATALOG.get((style, True))
        if cat is None:
            return None
        uid, code, dist, rc, fr, en = cat
        return EventKey(age_code, GENDER_MIXED, uid, code, dist, rc, fr, en)

    # Individual: "<F|M> <label> [<n> m]"
    mi = re.match(r"^([FM])\s+(.*)$", rest)
    if not mi:
        return None
    gender = GENDER_MALE if mi.group(1) == "M" else GENDER_FEMALE
    body = mi.group(2).strip()
    mb = re.match(r"^(.*?)(?:\s+(\d+)\s*m)?$", body)
    label = mb.group(1).strip()
    dist_txt = mb.group(2)

    # Disambiguate Portage 50 / 100 and Obstacle 100 (Masters) / 200
    lookup_label = label
    if label == "Portage" and dist_txt == "50":
        lookup_label = "Portage50"
    elif label == "Obstacle" and dist_txt == "100":
        lookup_label = "Obstacle100"

    cat = LIFESAVING_CATALOG.get((lookup_label, False))
    if cat is None:
        return None
    uid, code, dist, rc, fr, en = cat
    return EventKey(age_code, gender, uid, code, dist, rc, fr, en)


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def norm_key(*parts: Any) -> str:
    """Normalise a string to act as a dedup key (lower, strip accents)."""
    s = " ".join((str(p) if p is not None else "") for p in parts).strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s)


def truncate(s: str | None, n: int) -> str | None:
    if s is None:
        return None
    s = str(s)
    return s[:n]


def upper_key(s: str | None, n: int) -> str | None:
    if s is None:
        return None
    return unicodedata.normalize("NFKD", s).upper().encode("ascii", "ignore") \
        .decode("ascii")[:n]


def age_at(birthdate: dt.datetime | dt.date | None,
           ref: dt.date = AGE_DATE) -> int | None:
    """Age in whole years at `ref`, or None if no birthdate."""
    if birthdate is None:
        return None
    bd = birthdate.date() if isinstance(birthdate, dt.datetime) else birthdate
    years = ref.year - bd.year
    if (ref.month, ref.day) < (bd.month, bd.day):
        years -= 1
    return years


def pick_agegroup(agegroups: list[tuple],
                  athlete_age: int | None,
                  fallback_min: int, fallback_max: int) -> int | None:
    """From a list of (AGEGROUPID, AGEMIN, AGEMAX), pick the one that best
    contains the athlete's age.  Retained for reference / potential reuse
    (not called directly in the current per-age-bracket model, which uses
    event_targets routing instead).
    """
    if not agegroups:
        return None

    def in_bounds(amin, amax, x):
        lo = -10**9 if amin is None or amin < 0 else amin
        hi = 10**9  if amax is None or amax < 0 else amax
        return lo <= x <= hi

    if athlete_age is not None:
        matches = [(agid, amin, amax) for agid, amin, amax in agegroups
                   if in_bounds(amin, amax, athlete_age)]
        if matches:
            def span(t):
                _, amin, amax = t
                lo = 0   if amin is None or amin < 0 else amin
                hi = 999 if amax is None or amax < 0 else amax
                return hi - lo
            matches.sort(key=span)
            return matches[0][0]
    for agid, amin, amax in agegroups:
        if amin == fallback_min and amax == fallback_max:
            return agid
    for agid, amin, amax in agegroups:
        lo = -10**9 if amin is None or amin < 0 else amin
        hi = 10**9  if amax is None or amax < 0 else amax
        if lo <= fallback_min and fallback_max <= hi:
            return agid
    return agegroups[0][0]


def short_code_from_name(name: str, length: int = 10) -> str:
    """Build a club short code: keep capitals+digits, fall back to prefix."""
    caps = "".join(c for c in name if c.isupper() or c.isdigit())
    if 2 <= len(caps) <= length:
        return caps
    cleaned = re.sub(r"[^A-Za-z0-9]", "", unicodedata.normalize("NFKD", name)
                     .encode("ascii", "ignore").decode("ascii"))
    return cleaned[:length].upper() or "CLUB"


# --------------------------------------------------------------------------- #
# Fuzzy duplicate detection helpers
# --------------------------------------------------------------------------- #
import difflib as _difflib

FUZZY_CLUB_THRESHOLD    = 0.90   # club name similarity
FUZZY_ATHLETE_THRESHOLD = 0.90   # same-club athlete full-name similarity


def fuzzy_key(s: str) -> str:
    """Strong normalisation for dedup: lowercase, NFKD, strip accents and
    punctuation, collapse whitespace.  Two strings with the same
    fuzzy_key are almost certainly the same entity."""
    if s is None:
        return ""
    s = unicodedata.normalize("NFKD", str(s))
    s = "".join(c for c in s if not unicodedata.combining(c))
    s = s.lower()
    s = re.sub(r"[^\w\s]", " ", s)     # drop punctuation
    s = re.sub(r"\s+", " ", s).strip()
    return s


def similarity(a: str, b: str) -> float:
    """Ratio in [0, 1].  1.0 = identical, 0.0 = unrelated."""
    return _difflib.SequenceMatcher(None, a, b).ratio()


def find_fuzzy_club_duplicates(
        club_counts: dict[str, int],
        threshold: float = FUZZY_CLUB_THRESHOLD) -> list[tuple[str, str, float, int, int]]:
    """Return list of (name_a, name_b, similarity, count_a, count_b) for
    club-name pairs that look like typos/variants of each other.

    Pairs are reported when either:
      - their fuzzy_key is identical (e.g. 'Rouville SurfClub' / 'Rouville Surfclub'), or
      - their similarity on normalised text is >= threshold.
    """
    names = sorted(club_counts.keys(), key=str.lower)
    out: list[tuple[str, str, float, int, int]] = []
    for i in range(len(names)):
        a = names[i]
        ka = fuzzy_key(a)
        for j in range(i + 1, len(names)):
            b = names[j]
            kb = fuzzy_key(b)
            if not ka or not kb:
                continue
            if ka == kb:
                out.append((a, b, 1.0, club_counts[a], club_counts[b]))
                continue
            # Skip comparisons with wildly different lengths — saves time and
            # avoids false positives on very short names.
            if abs(len(ka) - len(kb)) > max(4, min(len(ka), len(kb)) // 2):
                continue
            s = similarity(ka, kb)
            if s >= threshold:
                out.append((a, b, s, club_counts[a], club_counts[b]))
    return out


def find_fuzzy_athlete_duplicates(
        athletes: dict[tuple, "Inscription"],
        threshold: float = FUZZY_ATHLETE_THRESHOLD
) -> dict[str, list[tuple]]:
    """Scan the athletes dict for suspect duplicates.

    Returns a dict with three keys:
        'same_license':          pairs sharing a LICENSE but with different names
        'same_club_fuzzy':       pairs in the same club whose full name is similar
        'cross_club_same_person':pairs in different clubs with same normalised
                                 first+last name AND same birthdate
    """
    results = {
        "same_license":           [],
        "same_club_fuzzy":        [],
        "cross_club_same_person": [],
    }

    # Index athletes by license
    by_license: dict[str, list[tuple[tuple, "Inscription"]]] = {}
    for akey, ins in athletes.items():
        lic = (ins.license or "").strip()
        if lic:
            by_license.setdefault(lic, []).append((akey, ins))

    # 1) same license, different name
    for lic, group in by_license.items():
        if len(group) < 2:
            continue
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                a_key, a = group[i]
                b_key, b = group[j]
                name_a = f"{a.first} {a.last}"
                name_b = f"{b.first} {b.last}"
                if fuzzy_key(name_a) != fuzzy_key(name_b):
                    results["same_license"].append(
                        (name_a, a.club, name_b, b.club, lic))

    # Group athletes by club for same-club fuzzy check
    by_club: dict[str, list[tuple[tuple, "Inscription"]]] = {}
    for akey, ins in athletes.items():
        by_club.setdefault(norm_key(ins.club), []).append((akey, ins))

    # 2) same club, similar full names
    for cnorm, group in by_club.items():
        n = len(group)
        if n < 2:
            continue
        keys = [fuzzy_key(f"{ins.first} {ins.last}") for _, ins in group]
        for i in range(n):
            ka = keys[i]
            if not ka:
                continue
            for j in range(i + 1, n):
                kb = keys[j]
                if not kb or ka == kb:
                    # Exact match after normalisation — same person twice.
                    # (Usually caught upstream by the dedup key, but guard
                    # for cases where license differs so the dedup key
                    # doesn't coalesce them.)
                    if ka == kb:
                        a_key, a = group[i]; b_key, b = group[j]
                        if a.license != b.license:
                            results["same_club_fuzzy"].append(
                                (f"{a.first} {a.last}", a.license or "-",
                                 f"{b.first} {b.last}", b.license or "-",
                                 a.club, 1.0))
                    continue
                if abs(len(ka) - len(kb)) > max(4, min(len(ka), len(kb)) // 2):
                    continue
                s = similarity(ka, kb)
                if s >= threshold:
                    # Extra guard: require BOTH first AND last to be reasonably
                    # similar (so "Alice Tremblay" vs "Alice Gauthier" doesn't
                    # trip the trigger just on the shared first name).
                    a_key, a = group[i]; b_key, b = group[j]
                    s_first = similarity(fuzzy_key(a.first), fuzzy_key(b.first))
                    s_last  = similarity(fuzzy_key(a.last),  fuzzy_key(b.last))
                    if s_first >= 0.70 and s_last >= 0.70:
                        results["same_club_fuzzy"].append(
                            (f"{a.first} {a.last}", a.license or "-",
                             f"{b.first} {b.last}", b.license or "-",
                             a.club, s))

    # 3) cross-club: same normalised name + same birthdate
    by_name_dob: dict[tuple, list[tuple[tuple, "Inscription"]]] = {}
    for akey, ins in athletes.items():
        bd = ins.birthdate
        bd_key = bd.date() if isinstance(bd, dt.datetime) else bd
        k = (fuzzy_key(f"{ins.first} {ins.last}"), bd_key)
        if k[0] and k[1] is not None:
            by_name_dob.setdefault(k, []).append((akey, ins))
    for (name_k, bd_k), group in by_name_dob.items():
        # Skip when everyone is in the same club (already handled above)
        clubs_ = {norm_key(ins.club) for _, ins in group}
        if len(clubs_) < 2:
            continue
        # Report each cross-club pair once
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                _, a = group[i]; _, b = group[j]
                if norm_key(a.club) == norm_key(b.club):
                    continue
                results["cross_club_same_person"].append(
                    (f"{a.first} {a.last}", a.club, b.club,
                     bd_k.isoformat() if bd_k else ""))

    return results


_TIME_RE_H = re.compile(r"^\s*(\d+):(\d{1,2}):(\d{1,2})(?:[.,](\d{1,3}))?\s*$")
_TIME_RE_M = re.compile(r"^\s*(\d+):(\d{1,2})(?:[.,](\d{1,3}))?\s*$")
_TIME_RE_S = re.compile(r"^\s*(\d+)(?:[.,](\d{1,3}))?\s*$")


def parse_best_time(val: Any) -> int | None:
    """Parse a swim best time into MILLISECONDS, or None.

    SPLASH's SWIMRESULT.ENTRYTIME / SWIMTIME column stores milliseconds,
    not hundredths of a second — confirmed against the 30-Deux federation
    sample where a 50 m obstacle result of SWIMTIME=45520 corresponds
    to 45.520 s (as milliseconds), not 7:35.20 (as centiseconds).

    Recognised formats:
        hh:mm:ss[.fff]
        mm:ss[.fff]          <-- most common in lifesaving meets
        ss[.fff]
        An Excel time object (datetime.time / timedelta-like)
        A float/int in seconds
    Fractional part is padded to 3 digits ('1:05' -> 65000 ms,
    '1:05.3' -> 65300 ms, '1:05.37' -> 65370 ms, '1:05.371' -> 65371 ms).
    """
    if val is None:
        return None
    if isinstance(val, dt.time):
        return ((val.hour * 3600 + val.minute * 60 + val.second) * 1000
                + val.microsecond // 1000)
    if isinstance(val, dt.timedelta):
        return int(round(val.total_seconds() * 1000))
    if isinstance(val, (int, float)):
        x = float(val)
        if 0 < x < 1:                    # Excel fraction-of-day
            return int(round(x * 24 * 3600 * 1000))
        if x > 0:                        # seconds
            return int(round(x * 1000))
        return None
    s = str(val).strip()
    if not s or s.lower() in ("nt", "n/a", "na", "-"):
        return None

    def _ms(fs: str | None) -> int:
        """Normalise a fractional-seconds string to milliseconds (0-999)."""
        fs = (fs or "0")
        fs = (fs + "000")[:3]
        return int(fs)

    m = _TIME_RE_H.match(s)
    if m:
        h, mm, ss, fs = m.group(1), m.group(2), m.group(3), m.group(4)
        total = (int(h) * 3600 + int(mm) * 60 + int(ss)) * 1000 + _ms(fs)
        return total if total > 0 else None
    m = _TIME_RE_M.match(s)
    if m:
        mm, ss, fs = m.group(1), m.group(2), m.group(3)
        total = (int(mm) * 60 + int(ss)) * 1000 + _ms(fs)
        return total if total > 0 else None
    m = _TIME_RE_S.match(s)
    if m:
        ss, fs = m.group(1), m.group(2)
        total = int(ss) * 1000 + _ms(fs)
        return total if total > 0 else None
    return None


def parse_birthdate(val: Any) -> dt.datetime | None:
    if val is None or (isinstance(val, float) and val != val):
        return None
    if isinstance(val, dt.datetime):
        return val
    if isinstance(val, dt.date):
        return dt.datetime(val.year, val.month, val.day)
    s = str(val).strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt)
        except ValueError:
            continue
    return None


# --------------------------------------------------------------------------- #
# Excel reader
# --------------------------------------------------------------------------- #
@dataclass
class Inscription:
    first: str
    last: str
    email: str | None
    club: str
    birthdate: dt.datetime | None
    license: str | None
    best_time_ms: int | None
    event: EventKey
    teammates: str | None = None  # raw string, used only for relay debug


# --------------------------------------------------------------------------- #
# Issue collector — data-quality warnings surfaced at end of run
# --------------------------------------------------------------------------- #
@dataclass
class Issue:
    severity: str     # 'WARNING' | 'NOTE'
    category: str     # short tag (e.g. 'bad_time', 'no_dob', 'age_mismatch')
    message: str      # human-readable description
    row: int | None = None      # xlsx row number (1-based) if applicable


class IssueCollector:
    """Bucket for data-quality findings.  At end of run we print a summary
    grouped by category, capped at N items per category."""

    def __init__(self, max_per_category: int = 10):
        self.issues: list[Issue] = []
        self.max_per_category = max_per_category

    def add(self, severity: str, category: str, message: str,
            row: int | None = None) -> None:
        self.issues.append(Issue(severity, category, message, row))

    def warn(self, category: str, message: str, row: int | None = None):
        self.add("WARNING", category, message, row)

    def note(self, category: str, message: str, row: int | None = None):
        self.add("NOTE", category, message, row)

    def by_category(self):
        from collections import defaultdict
        out: dict[tuple, list[Issue]] = defaultdict(list)
        for i in self.issues:
            out[(i.severity, i.category)].append(i)
        return out

    def report(self, title: str = "Issues",
               out_file=None,
               full: bool = False) -> None:
        """Print (and optionally write to `out_file`) the issues section.

        full=True removes the per-category cap (every issue listed)."""
        if not self.issues:
            return
        buckets = self.by_category()
        ordered = sorted(buckets.items(),
                         key=lambda kv: (kv[0][0] != "WARNING", -len(kv[1])))
        cap = 10**9 if full else self.max_per_category

        lines: list[str] = []
        lines.append("")
        lines.append("=" * 60)
        lines.append(f"  {title}")
        lines.append("=" * 60)
        for (sev, cat), items in ordered:
            lines.append(f"  [{sev}] {cat}: {len(items)}")
            for it in items[:cap]:
                suffix = f" (row {it.row})" if it.row else ""
                lines.append(f"       - {it.message}{suffix}")
            if len(items) > cap:
                lines.append(f"       … and {len(items) - cap} more")
        lines.append("=" * 60)

        block = "\n".join(lines)
        print(block)
        if out_file is not None:
            out_file.write(block + "\n")


def read_attendees(xlsx_path: Path,
                   issues: IssueCollector | None = None) -> list[Inscription]:
    """Parse the Attendees sheet into a list of Inscription records.

    If an IssueCollector is provided, data-quality problems encountered
    while parsing (missing names, unparseable ticket types, bad times,
    bad birthdates, truncated names, duplicate athlete/event pairs) are
    reported into it.
    """
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    if "Attendees" not in wb.sheetnames:
        raise SystemExit(f"Sheet 'Attendees' not found in {xlsx_path}")
    ws = wb["Attendees"]
    rows = list(ws.iter_rows(values_only=True))
    header = [str(c or "").strip() for c in rows[0]]
    def col(name): return header.index(name)
    i_first = col("First Name")
    i_last = col("Last Name")
    i_email = col("Email")
    i_ticket = col("Ticket Type")
    i_best = col("Best time")
    i_club = col("Club")
    i_dob = col("DD/MM/YYYY")
    i_lic = col("NRAN")
    try:
        i_team = header.index("Teammate(s) + NRAN")
    except ValueError:
        i_team = None

    out: list[Inscription] = []
    # Track duplicate (athlete, event) pairs from the xlsx itself
    seen_pairs: dict[tuple, tuple[int, int | None]] = {}

    for row_idx, r in enumerate(rows[1:], start=2):
        if not r:
            continue
        first = r[i_first]; last = r[i_last]
        if not first or not last:
            # Skip fully empty rows silently; warn on partial ones.
            if any(cell not in (None, "") for cell in r):
                if issues:
                    issues.warn("missing_name",
                                f"row missing first or last name",
                                row=row_idx)
            continue
        ticket = (str(r[i_ticket] or "").strip()) if r[i_ticket] else ""
        if not ticket:
            if issues:
                issues.warn("missing_ticket",
                            f"{first} {last}: empty Ticket Type",
                            row=row_idx)
            continue
        ev = parse_ticket(ticket)
        if ev is None:
            # Distinguish intentional non-race tickets from unknown ones
            if any(ticket.startswith(p) for p in NON_RACE_PREFIXES):
                # legitimate non-race (Banquet, Coach, etc.)
                pass
            else:
                if issues:
                    issues.warn("unknown_ticket",
                                f"{first} {last}: unrecognised ticket "
                                f"{ticket!r}",
                                row=row_idx)
            continue

        # Parse best time + birthdate; track failures
        raw_time = r[i_best]
        best_ms = parse_best_time(raw_time)
        if raw_time not in (None, "") and best_ms is None \
                and str(raw_time).strip().lower() not in ("nt", "n/a", "na", "-"):
            if issues:
                issues.warn("bad_time",
                            f"{first} {last} {ticket!r}: "
                            f"can't parse time {raw_time!r}",
                            row=row_idx)

        raw_dob = r[i_dob]
        bd = parse_birthdate(raw_dob)
        if raw_dob not in (None, "") and bd is None:
            if issues:
                issues.warn("bad_birthdate",
                            f"{first} {last}: can't parse DOB {raw_dob!r}",
                            row=row_idx)

        # Name-length warnings (will be truncated for SPLASH columns)
        if len(str(first)) > 30 and issues:
            issues.note("truncated_name",
                        f"first name truncated (>30 chars): {first!r}",
                        row=row_idx)
        if len(str(last)) > 50 and issues:
            issues.note("truncated_name",
                        f"last name truncated (>50 chars): {last!r}",
                        row=row_idx)
        club_raw = (r[i_club] or "Unattached")
        if len(str(club_raw)) > 80 and issues:
            issues.note("truncated_name",
                        f"club name truncated (>80 chars): {club_raw!r}",
                        row=row_idx)

        # Duplicate (athlete, event) in xlsx — we'll keep the best time, but
        # let the user know so they can clean up the source sheet.
        pair_key = (norm_key(first, last), r[i_lic] or "", ev.key())
        if pair_key in seen_pairs and issues:
            prev_row, prev_cs = seen_pairs[pair_key]
            issues.note("duplicate_entry",
                        f"{first} {last} entered in {ticket!r} again "
                        f"(first seen row {prev_row}); keeping best time",
                        row=row_idx)
        seen_pairs[pair_key] = (row_idx, best_ms)

        out.append(Inscription(
            first=str(first).strip(),
            last=str(last).strip(),
            email=(r[i_email] or None),
            club=str(club_raw).strip(),
            birthdate=bd,
            license=(str(r[i_lic]).strip() if r[i_lic] else None),
            best_time_ms=best_ms,
            event=ev,
            teammates=(str(r[i_team]).strip()
                       if i_team is not None and r[i_team] else None),
        ))
    return out


# --------------------------------------------------------------------------- #
# MDB writer
# --------------------------------------------------------------------------- #
class MDB:
    """Thin wrapper around a UCanAccess JDBC connection."""

    def __init__(self, path: Path, dry_run: bool = False):
        self.path = path
        self.dry_run = dry_run
        jars = (
            glob.glob(os.path.join(UCANACCESS_DIR, "ucanaccess-*.jar"))
            + glob.glob(os.path.join(UCANACCESS_DIR, "lib", "*.jar"))
            # Support the flattened layout used by the Docker image,
            # where all five jars live directly under /opt/ucanaccess/
            # with no lib/ subdir.
            + glob.glob(os.path.join(UCANACCESS_DIR, "*.jar"))
        )
        # Dedup while preserving order
        seen = set()
        jars = [j for j in jars if not (j in seen or seen.add(j))]
        if not jars:
            raise SystemExit(
                f"UCanAccess jars not found under {UCANACCESS_DIR}. "
                "Set UCANACCESS_DIR env var."
            )
        url = (f"jdbc:ucanaccess://{path};openExclusive=false;"
               "memory=true;ignoreCase=true")
        self.conn = jaydebeapi.connect(
            "net.ucanaccess.jdbc.UcanaccessDriver", url, [], jars)
        self.cur = self.conn.cursor()
        self._uid = self._read_uid()
        self._start_uid = self._uid
        # JPype is started by jaydebeapi.connect(); grab Timestamp class
        import jpype  # noqa
        self._Timestamp = jpype.JClass("java.sql.Timestamp")

    # ----- UID allocator -----
    def _read_uid(self) -> int:
        self.cur.execute("SELECT LASTUID FROM BSUIDTABLE WHERE NAME='BS_GLOBAL_UID'")
        return int(self.cur.fetchone()[0])

    def next_id(self) -> int:
        self._uid += 1
        return self._uid

    def flush_uid(self):
        self.cur.execute(
            "UPDATE BSUIDTABLE SET LASTUID=? WHERE NAME='BS_GLOBAL_UID'",
            [self._uid])

    # ----- DML -----
    def exec(self, sql: str, params: list | None = None):
        if self.dry_run:
            return
        self.cur.execute(sql, params or [])

    def exec_many(self, sql: str, batch: list[list]):
        if self.dry_run or not batch:
            return
        self.cur.executemany(sql, batch)

    def insert(self, table: str, row: dict):
        """INSERT with only non-None columns – avoids UCanAccess NPE on nulls."""
        if self.dry_run:
            return
        cols = [c for c, v in row.items() if v is not None]
        vals = [self._to_jdbc(row[c]) for c in cols]
        if not cols:
            return
        placeholders = ",".join("?" * len(cols))
        sql = (f"INSERT INTO {table} ({','.join(cols)}) "
               f"VALUES ({placeholders})")
        self.cur.execute(sql, vals)

    def insert_many(self, table: str, rows: list[dict]):
        if self.dry_run or not rows:
            return
        # group by the set of non-null columns so each batch has homogeneous SQL
        from itertools import groupby
        def key(r): return tuple(c for c, v in r.items() if v is not None)
        rows_sorted = sorted(rows, key=key)
        for cols, grp in groupby(rows_sorted, key=key):
            cols = list(cols)
            if not cols:
                continue
            placeholders = ",".join("?" * len(cols))
            sql = (f"INSERT INTO {table} ({','.join(cols)}) "
                   f"VALUES ({placeholders})")
            batch = [[self._to_jdbc(r[c]) for c in cols] for r in grp]
            self.cur.executemany(sql, batch)

    def update(self, table: str, where: dict, updates: dict):
        """UPDATE with non-None updates only, filtered by `where` dict."""
        if self.dry_run or not updates:
            return
        set_cols = list(updates.keys())
        where_cols = list(where.keys())
        sql = (f"UPDATE {table} SET "
               + ", ".join(f"{c}=?" for c in set_cols)
               + " WHERE "
               + " AND ".join(f"{c}=?" for c in where_cols))
        params = ([self._to_jdbc(updates[c]) for c in set_cols]
                  + [self._to_jdbc(where[c])  for c in where_cols])
        self.cur.execute(sql, params)

    def query(self, sql: str, params: list | None = None):
        """SELECT and return all rows as list[tuple]."""
        self.cur.execute(sql, params or [])
        return self.cur.fetchall()

    def _to_jdbc(self, v):
        """Convert Python value into something JDBC's setObject accepts."""
        if isinstance(v, dt.datetime):
            # java.sql.Timestamp(long ms since epoch)
            epoch_ms = int(v.timestamp() * 1000)
            return self._Timestamp(epoch_ms)
        if isinstance(v, dt.date):
            epoch_ms = int(dt.datetime(v.year, v.month, v.day).timestamp() * 1000)
            return self._Timestamp(epoch_ms)
        return v

    def commit(self):
        if self.dry_run:
            print("[dry-run] skipping commit; rolling back")
            self.conn.rollback()
        else:
            self.flush_uid()
            self.conn.commit()

    def close(self):
        self.conn.close()


# --------------------------------------------------------------------------- #
# Main loader
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--xlsx", required=True, type=Path)
    ap.add_argument("--mdb", required=True, type=Path,
                    help="Target .mdb file (will be modified).")
    ap.add_argument("--dry-run", action="store_true",
                    help="Parse and plan everything but rollback at the end.")
    ap.add_argument("--wipe", action="store_true",
                    help="Delete any existing clubs/athletes/events/entries "
                         "before loading.")
    ap.add_argument("--issues-full", action="store_true",
                    help="List every issue, no per-category cap. "
                         "Useful when handing the report to someone who "
                         "has to fix the source xlsx row by row.")
    ap.add_argument("--issues-out", type=Path,
                    help="Also write the issues section to the given "
                         "file (plain text). Implies --issues-full.")
    args = ap.parse_args()

    if args.issues_out:
        args.issues_full = True

    if not args.xlsx.exists():
        sys.exit(f"xlsx not found: {args.xlsx}")
    if not args.mdb.exists():
        sys.exit(f"mdb not found: {args.mdb}")

    print(f"Reading {args.xlsx}...")
    issues = IssueCollector()
    inscriptions = read_attendees(args.xlsx, issues)
    print(f"  {len(inscriptions)} race inscriptions extracted")

    # Aggregate
    clubs: dict[str, str] = {}          # norm_name -> display name
    athletes: dict[tuple, Inscription] = {}  # (norm first,last,license) -> record
    events: dict[tuple, EventKey] = {}  # EventKey.key() -> EventKey
    ind_entries: list[tuple] = []       # (athlete_key, event_key, best_ms)
    relay_groups: dict[tuple, list[tuple]] = defaultdict(list)
    # relay_groups key: (club_norm, event_key) -> list[(athlete_key, best_ms)]

    for ins in inscriptions:
        club_norm = norm_key(ins.club)
        clubs.setdefault(club_norm, ins.club)

        ath_key = (norm_key(ins.first, ins.last), ins.license or "")
        if ath_key not in athletes:
            athletes[ath_key] = ins
        events.setdefault(ins.event.key(), ins.event)

        if ins.event.relay_count == 1:
            ind_entries.append((ath_key, ins.event.key(), ins.best_time_ms))
        else:
            relay_groups[(club_norm, ins.event.key())].append(
                (ath_key, ins.best_time_ms))

    print(f"  distinct clubs:    {len(clubs)}")
    print(f"  distinct athletes: {len(athletes)}")
    print(f"  distinct events:   {len(events)}  "
          f"(individual: {sum(1 for k,e in events.items() if e.relay_count==1)}, "
          f"relay: {sum(1 for k,e in events.items() if e.relay_count>1)})")
    print(f"  individual entries: {len(ind_entries)}")
    print(f"  relay squads:      {len(relay_groups)}")

    # ----- Cross-row data-quality checks -----
    # athletes missing a birthdate (will fall back to xlsx bracket on any
    # future MM age-split — fine but worth flagging).
    for akey, ins in athletes.items():
        if ins.birthdate is None:
            issues.warn(
                "no_dob",
                f"{ins.first} {ins.last} ({ins.club}) has no birthdate — "
                f"age-based routing will fall back to ticket bracket")

    # Individual inscription age sanity-check: compare the computed age
    # at AGE_DATE to the age bracket the ticket is for.  Flag mismatches.
    bracket_by_code = {k: (v[0], v[1]) for k, v in AGE_GROUPS.items()}
    for ins in inscriptions:
        if ins.event.relay_count != 1:
            continue       # relays aren't individually age-bracketed here
        age = age_at(ins.birthdate)
        if age is None:
            continue
        amin, amax = bracket_by_code[ins.event.age_code]
        if age < amin or age > amax:
            issues.warn(
                "age_bracket_mismatch",
                f"{ins.first} {ins.last} age {age} outside ticket bracket "
                f"{ins.event.age_code} ({amin}-{amax}) "
                f"for {ins.event.name_fr}")

    # Relay squads that are incomplete (fewer athletes than required
    # by the event's relay_count) — MM user will need to fill them in.
    for (cnorm, ekey), members in relay_groups.items():
        ev = events[ekey]
        leftovers = len(members) % ev.relay_count
        if len(members) < ev.relay_count:
            issues.warn(
                "incomplete_relay",
                f"{clubs[cnorm]}: {len(members)}/{ev.relay_count} athletes "
                f"for {ev.name_fr} ({ev.age_code})")
        elif leftovers:
            n_squads = (len(members) + ev.relay_count - 1) // ev.relay_count
            issues.note(
                "extra_relay_members",
                f"{clubs[cnorm]}: {len(members)} athletes for "
                f"{ev.name_fr} ({ev.age_code}) — split into {n_squads} "
                f"squads, the last one has only {leftovers}/{ev.relay_count}")

    # Non-race-only clubs / athletes (informational).  Re-scan the xlsx
    # so we can compare what's in the sheet vs what we imported.
    wb = openpyxl.load_workbook(args.xlsx, data_only=True)
    ws_all = wb["Attendees"]
    rows_all = list(ws_all.iter_rows(values_only=True))
    hdr = [str(c or "").strip() for c in rows_all[0]]
    i_f  = hdr.index("First Name")
    i_l  = hdr.index("Last Name")
    i_cl = hdr.index("Club")
    i_lc = hdr.index("NRAN")
    all_clubs: set[str] = set()
    all_names: set[str] = set()   # athlete name only (not keyed on license)
    for r in rows_all[1:]:
        if not r or not r[i_f] or not r[i_l]:
            continue
        all_clubs.add(norm_key(r[i_cl] or "Unattached"))
        all_names.add(norm_key(r[i_f], r[i_l]))
    # Athletes that DID get imported — collapse to name-only identifiers
    # so that someone who has both a race row (with license) and a Banquet
    # row (no license) is counted as one person, not two.
    race_names = {akey[0] for akey in athletes.keys()}
    race_clubs = set(clubs.keys())
    n_club_skipped = len(all_clubs - race_clubs)
    n_ath_skipped = len(all_names - race_names)
    if n_club_skipped:
        issues.note("non_race_only_club",
                    f"{n_club_skipped} club(s) appear only on non-race "
                    f"tickets (Banquet/Coach/Officiel/…) — not imported")
    if n_ath_skipped:
        issues.note("non_race_only_athlete",
                    f"{n_ath_skipped} attendee(s) only bought non-race "
                    f"tickets (supporters, coaches, officials, hotel) "
                    f"— not imported as athletes")

    # ----- Fuzzy duplicate detection (clubs + athletes) -----
    # Count rows per club display-name so the warning can show scale.
    club_row_counts: dict[str, int] = defaultdict(int)
    for ins in inscriptions:
        club_row_counts[ins.club] += 1

    for a, b, sim, ca, cb in find_fuzzy_club_duplicates(dict(club_row_counts)):
        issues.warn(
            "possible_duplicate_club",
            f"{a!r} ({ca} rows) vs {b!r} ({cb} rows) "
            f"— similarity {sim:.2f}")

    fuzzy = find_fuzzy_athlete_duplicates(athletes)
    for (name_a, club_a, name_b, club_b, lic) in fuzzy["same_license"]:
        issues.warn(
            "license_name_mismatch",
            f"license {lic!r}: {name_a!r} ({club_a}) vs {name_b!r} ({club_b})"
            f" — same license, different name spelling")
    for (name_a, lic_a, name_b, lic_b, club, sim) in fuzzy["same_club_fuzzy"]:
        issues.warn(
            "possible_duplicate_athlete",
            f"{club}: {name_a!r} (NRAN {lic_a}) vs {name_b!r} "
            f"(NRAN {lic_b}) — similarity {sim:.2f}")
    for (name, club_a, club_b, dob) in fuzzy["cross_club_same_person"]:
        issues.warn(
            "same_person_different_club",
            f"{name!r} born {dob} appears in both {club_a!r} and "
            f"{club_b!r} — probably the same person")

    # ----- open MDB -----
    print(f"\nOpening {args.mdb}...")
    db = MDB(args.mdb, dry_run=args.dry_run)
    print(f"  starting BS_GLOBAL_UID = {db._uid}")

    if args.wipe:
        for sql in [
            "DELETE FROM RELAYPOSITION",
            "DELETE FROM RELAY",
            "DELETE FROM SWIMRESULT",
            "DELETE FROM AGEGROUP",
            "DELETE FROM SWIMEVENT",
            "DELETE FROM SWIMSESSION",
            "DELETE FROM ATHLETE",
            "DELETE FROM CLUB",
            "DELETE FROM SWIMSTYLE WHERE SWIMSTYLEID > 1058",
        ]:
            if not args.dry_run:
                db.cur.execute(sql)
        print("  wiped existing clubs/athletes/events/entries")

    # ----- Preload existing rows (for additive/idempotent re-runs) -----
    stats = {
        "club_new": 0, "club_updated": 0,
        "athlete_new": 0, "athlete_gender_fix": 0,
        "athlete_license_fix": 0, "athlete_birthdate_fix": 0,
        "athlete_club_fix": 0,
        "style_new": 0,
        "session_new": 0,
        "event_new": 0, "agegroup_new": 0,
        "entry_new": 0, "entry_time_faster": 0,
        "relay_new": 0, "relayposition_new": 0,
        "combined_new": 0,
    }

    # CLUB lookup: norm_name -> (CLUBID, NAME)
    existing_clubs: dict[str, tuple[int, str]] = {}
    for cid, name in db.query("SELECT CLUBID, NAME FROM CLUB"):
        if name:
            existing_clubs[norm_key(name)] = (int(cid), name)

    # ATHLETE lookup: (norm first+last, license_norm) -> full row dict
    existing_athletes: dict[tuple, dict] = {}
    for aid, clubid, first, last, gender, bdate, lic in db.query(
        "SELECT ATHLETEID, CLUBID, FIRSTNAME, LASTNAME, GENDER, BIRTHDATE, "
        "LICENSE FROM ATHLETE"):
        k = (norm_key(first or "", last or ""), (lic or "").strip())
        existing_athletes[k] = {
            "ATHLETEID": int(aid),
            "CLUBID":    int(clubid) if clubid is not None else None,
            "GENDER":    int(gender) if gender is not None else None,
            "BIRTHDATE": bdate,
            "LICENSE":   lic,
        }

    # SWIMSTYLE lookup: UNIQUEID -> SWIMSTYLEID   (only catalog rows, STROKE=0)
    existing_styles: dict[int, int] = {}
    for sid, uid in db.query(
        "SELECT SWIMSTYLEID, UNIQUEID FROM SWIMSTYLE "
        "WHERE STROKE=0 AND UNIQUEID IS NOT NULL"):
        if uid is not None:
            existing_styles[int(uid)] = int(sid)

    # SWIMEVENT lookup: events grouped by (SWIMSTYLEID, GENDER).
    # Each entry is a list of sub-events (SWIMEVENTID, AGEGROUPID, AMIN, AMAX).
    # Our model keeps one AGEGROUP per SWIMEVENT (= one age bracket per
    # event), but MM users may have split an event into several sub-events
    # after we created them (e.g. 15-18 -> 15-16 + 17-18).  The re-run logic
    # recognises that case via "subset partitioning".
    events_by_sg: dict[tuple, list[tuple]] = defaultdict(list)
    for eid, styid, gen, agid, amin, amax in db.query(
        "SELECT e.SWIMEVENTID, e.SWIMSTYLEID, e.GENDER, "
        "       a.AGEGROUPID, a.AGEMIN, a.AGEMAX "
        "FROM SWIMEVENT e LEFT JOIN AGEGROUP a "
        "       ON a.SWIMEVENTID=e.SWIMEVENTID"):
        if styid is None or gen is None:
            continue
        events_by_sg[(int(styid), int(gen))].append((
            int(eid),
            int(agid) if agid is not None else None,
            int(amin) if amin is not None else None,
            int(amax) if amax is not None else None,
        ))

    # SWIMRESULT lookup: (ATHLETEID, SWIMEVENTID) -> (SWIMRESULTID, ENTRYTIME)
    existing_results: dict[tuple, tuple[int, int | None]] = {}
    for srid, aid, seid, etime in db.query(
        "SELECT SWIMRESULTID, ATHLETEID, SWIMEVENTID, ENTRYTIME FROM SWIMRESULT"):
        if aid is None or seid is None:
            continue
        existing_results[(int(aid), int(seid))] = (
            int(srid),
            int(etime) if etime is not None else None,
        )

    # RELAY lookup: (CLUBID, SWIMEVENTID, TEAMNUMBER) -> RELAYID
    existing_relays: dict[tuple, int] = {}
    for rid, cid, seid, tnum in db.query(
        "SELECT RELAYID, CLUBID, SWIMEVENTID, TEAMNUMBER FROM RELAY"):
        if cid is None or seid is None:
            continue
        existing_relays[(int(cid), int(seid), int(tnum or 0))] = int(rid)

    # RELAYPOSITION lookup: set of (RELAYID, RELAYNUMBER)
    existing_relay_pos: set[tuple] = set()
    for rid, rnum in db.query(
        "SELECT RELAYID, RELAYNUMBER FROM RELAYPOSITION"):
        if rid is None or rnum is None:
            continue
        existing_relay_pos.add((int(rid), int(rnum)))

    # Is this a fresh meet or a re-run?
    total_events_present = sum(len(v) for v in events_by_sg.values())
    is_update = bool(existing_clubs or existing_athletes or total_events_present)
    if is_update:
        print(f"\n  Existing data detected — running in ADDITIVE mode.")
        print(f"    clubs already present:     {len(existing_clubs)}")
        print(f"    athletes already present:  {len(existing_athletes)}")
        print(f"    events already present:    {total_events_present}")
        print(f"    entries already present:   {len(existing_results)}")
        print(f"    relays already present:    {len(existing_relays)}")
    else:
        print(f"\n  Empty meet — doing a fresh load.")

    if args.wipe:
        if is_update:
            print("  WARNING: --wipe will delete existing data before loading.")
        for sql in [
            "DELETE FROM RELAYPOSITION",
            "DELETE FROM RELAY",
            "DELETE FROM SWIMRESULT",
            "DELETE FROM AGEGROUP",
            "DELETE FROM SWIMEVENT",
            "DELETE FROM SWIMSESSION",
            "DELETE FROM ATHLETE",
            "DELETE FROM CLUB",
            "DELETE FROM SWIMSTYLE WHERE SWIMSTYLEID > 1058",
        ]:
            if not args.dry_run:
                db.cur.execute(sql)
        print("  wiped existing clubs/athletes/events/entries")
        # Reset lookups to empty — we just deleted everything
        existing_clubs.clear()
        existing_athletes.clear()
        existing_styles.clear()
        events_by_sg.clear()
        existing_results.clear()
        existing_relays.clear()
        existing_relay_pos.clear()
        is_update = False

    INT_MAX = 2147483647

    # ----- SWIMSTYLE (Société de Sauvetage catalog) -----
    # STROKE=0 TECHNIQUE=0 tells SPLASH this is a catalog item.  Keyed by
    # UNIQUEID; already-present catalog entries reuse the same SWIMSTYLEID.
    style_ids: dict[int, int] = dict(existing_styles)
    sort_seed = 1000 + len(existing_styles)
    distinct_styles: dict[int, "EventKey"] = {}
    for ev in events.values():
        distinct_styles.setdefault(ev.uniqueid, ev)
    for uid, ev in sorted(distinct_styles.items()):
        if uid in style_ids:
            continue
        sid = db.next_id()
        style_ids[uid] = sid
        sort_seed += 1
        db.insert("SWIMSTYLE", {
            "SWIMSTYLEID": sid,
            "CODE":        truncate(ev.code, 10) or None,
            "NAME":        truncate(ev.name_fr, 50),
            "DISTANCE":    ev.distance,
            "RELAYCOUNT":  ev.relay_count,
            "STROKE":      LIFESAVING_STROKE,
            "TECHNIQUE":   LIFESAVING_TECHNIQUE,
            "UNIQUEID":    uid,
            "SORTCODE":    sort_seed,
        })
        stats["style_new"] += 1

    # ----- SWIMSESSION (placeholder) -----
    # Reuse the first existing session if one is present; otherwise create
    # a new placeholder.  We never modify an existing session (user may
    # have reorganised it in MM).
    rows = db.query("SELECT SWIMSESSIONID FROM SWIMSESSION "
                    "ORDER BY SWIMSESSIONID")
    if rows:
        session_id = int(rows[0][0])
    else:
        session_id = db.next_id()
        db.insert("SWIMSESSION", {
            "SWIMSESSIONID":     session_id, "SESSIONNUMBER": 1,
            "NAME":              truncate(PLACEHOLDER_SESSION_NAME, 100),
            "STARTDATE":         PLACEHOLDER_SESSION_DATE,
            "DAYTIME":           PLACEHOLDER_SESSION_DATE,
            "COURSE":            COURSE_LCM, "POOLTYPE": POOLTYPE_LCM,
            "LANEMIN": 1, "LANEMAX": 8, "TIMING": 1, "TOUCHPADMODE": 1,
            "MAXENTRIESATHLETE": 10, "MAXENTRIESRELAY": 5,
            "ROUNDTOTENTHS":     "F",
            "FOLLOWING":         "F",
            "POOLGLOBAL":        "F",
        })
        stats["session_new"] = 1

    # ----- SWIMEVENT + AGEGROUP -----
    # Each xlsx event key resolves to a list of target (SWIMEVENTID, AGEGROUPID,
    # AGEMIN, AGEMAX). For the common case there is exactly one target; for
    # MM-split cases (the meet director replaced one event with several
    # narrower age-bracket events) there are several — entries are then
    # routed to the one covering the athlete's real age.
    event_targets: dict[tuple, list[tuple]] = {}

    def _span(amin, amax):
        lo = 0   if amin is None or amin < 0 else amin
        hi = 999 if amax is None or amax < 0 else amax
        return lo, hi

    # Figure out the next EVENTNUMBER starting point
    rows = db.query("SELECT COALESCE(MAX(EVENTNUMBER),0) FROM SWIMEVENT")
    next_event_no = int(rows[0][0]) if rows and rows[0][0] is not None else 0

    # Sort key for the event list:
    #   1) by catalog UID           — all "Nage avec obstacles" together, etc.
    #   2) by age bracket           — 15-18, then MASTERS, then OPEN
    #   3) by gender, F before M    — (gender 2 before gender 1; 0=Mixed last)
    # The age bracket order happens to fall out of alphabetical string order
    # on the age_code values "1518" < "MASTERS" < "OPEN".
    _AGE_ORDER = {"1518": 0, "MASTERS": 1, "OPEN": 2}
    _GENDER_ORDER = {GENDER_FEMALE: 0, GENDER_MALE: 1, GENDER_ALL: 2}

    for ek, ev in sorted(
            events.items(),
            key=lambda kv: (kv[1].uniqueid,
                            _AGE_ORDER.get(kv[1].age_code, 99),
                            _GENDER_ORDER.get(kv[1].gender, 99))):
        style_id = style_ids[ev.uniqueid]
        age_min, age_max, _age_name = AGE_GROUPS[ev.age_code]
        xmin, xmax = age_min, age_max
        sg_key = (style_id, ev.gender)
        candidates = events_by_sg.get(sg_key, [])

        #
        # 1) Exact match (same [amin, amax]).
        exact = [c for c in candidates
                 if c[2] == xmin and c[3] == xmax]
        if exact:
            event_targets[ek] = [exact[0]]
            continue

        #
        # 2) MM-split: multiple existing events, each ⊆ [xmin, xmax], whose
        #    union tiles [xmin, xmax] contiguously.  This takes precedence
        #    over a single wider container (e.g. an Open [0,99] event) so
        #    that narrower event brackets are preferred when they partition
        #    our xlsx bracket exactly.
        subsets = []
        for c in candidates:
            lo, hi = _span(c[2], c[3])
            if xmin <= lo and hi <= xmax and (lo, hi) != (xmin, xmax):
                subsets.append(c)
        if subsets:
            bounds = sorted([_span(c[2], c[3]) for c in subsets])
            covered = xmin
            gap = False
            for lo, hi in bounds:
                if lo > covered:
                    gap = True; break
                covered = max(covered, hi + 1)
            if not gap and covered > xmax:
                # Split detected — route entries to these sub-events
                event_targets[ek] = subsets
                continue

        #
        # 3) Single existing event whose bracket fully contains ours.
        #    Prefer the tightest one.
        containers = []
        for c in candidates:
            lo, hi = _span(c[2], c[3])
            if lo <= xmin and xmax <= hi:
                containers.append(c)
        if containers:
            containers.sort(key=lambda c: _span(c[2], c[3])[1] - _span(c[2], c[3])[0])
            event_targets[ek] = [containers[0]]
            continue

        #
        # 4) No match — create a brand-new SWIMEVENT + AGEGROUP.
        next_event_no += 1
        eid = db.next_id()
        db.insert("SWIMEVENT", {
            "SWIMEVENTID":         eid,
            "EVENTNUMBER":         next_event_no,
            "SWIMSESSIONID":       session_id,
            "SWIMSTYLEID":         style_id,
            "SORTCODE":            next_event_no,
            "GENDER":              ev.gender,
            "ROUND":               1,
            "FINALORDER":          1,
            "PREVEVENTID":         -1,
            "ENTRYTIMECONVERSION": 1,
            "ENTRYTIMEPERCENT":    2,
            "SEEDINGGLOBAL":       "T",
            "SEEDBONUSLAST":       "F",
            "SEEDEXHLAST":         "F",
            "SEEDLATEENTRYLAST":   "F",
            "SPLASHMECANEDIT":     "F",
            "TWOPERLANE":          "F",
            "PFINEIGNORE":         "F",
            "COMBINEAGEGROUPS":    "F",
            "MASTERS": "T" if ev.age_code == "MASTERS" else "F",
            "LANEMAX":             0,
            "MAXENTRIES":          32767,
            "FEE":                 0.0,
        })
        stats["event_new"] += 1

        ag_id = db.next_id()
        db.insert("AGEGROUP", {
            "AGEGROUPID":    ag_id,
            "SWIMEVENTID":   eid,
            "AGEMIN":        age_min,
            "AGEMAX":        age_max,
            "GENDER":        ev.gender,
            "SORTCODE":      1,
            "AGEBYTOTAL":    "F", "ALLOFFICIAL": "T",
            "FORCEPRELIM":   "T", "USEFORMEDALS": "T",
            "USEFORSCORING": "T", "SEEDWITHTSONLY": "F",
            "SCORETYPE":     1,   "RESULTCOUNT": 0,
        })
        stats["agegroup_new"] += 1
        events_by_sg[sg_key].append((eid, ag_id, age_min, age_max))
        event_targets[ek] = [(eid, ag_id, age_min, age_max)]

    # ----- CLUB -----
    club_ids: dict[str, int] = {}
    for cnorm, cname in sorted(clubs.items(), key=lambda kv: kv[1].lower()):
        if cnorm in existing_clubs:
            club_ids[cnorm] = existing_clubs[cnorm][0]
            continue
        cid = db.next_id()
        club_ids[cnorm] = cid
        short = short_code_from_name(cname, 10)
        db.insert("CLUB", {
            "CLUBID":    cid,
            "NAME":      truncate(cname, 80),
            "SHORTNAME": truncate(short, 30),
            "NATION":    MEET_NATION,
            "CLUBTYPE":  1,
            "CODE":      truncate(short, 10),
        })
        stats["club_new"] += 1
        existing_clubs[cnorm] = (cid, cname)

    # ----- ATHLETE -----
    athlete_ids: dict[tuple, int] = {}
    # Pre-compute per-athlete inferred gender from individual tickets
    inferred_gender: dict[tuple, int] = {}
    for e in inscriptions:
        if e.event.relay_count != 1:
            continue
        k = (norm_key(e.first, e.last), (e.license or "").strip())
        # first individual-ticket gender wins (consistent for a given athlete)
        inferred_gender.setdefault(k, e.event.gender)

    for akey, ins in athletes.items():
        club_id = club_ids[norm_key(ins.club)]
        new_gender = inferred_gender.get(akey, GENDER_ALL)
        if akey in existing_athletes:
            existing = existing_athletes[akey]
            aid = existing["ATHLETEID"]
            athlete_ids[akey] = aid
            # Apply safe, additive updates
            updates = {}
            # Gender: fix from 0/NULL -> 1/2 if we now know it
            if (existing["GENDER"] in (None, GENDER_ALL)
                    and new_gender in (GENDER_MALE, GENDER_FEMALE)):
                updates["GENDER"] = new_gender
                stats["athlete_gender_fix"] += 1
            # License: fill in if missing
            if not existing["LICENSE"] and ins.license:
                updates["LICENSE"] = truncate(ins.license, 20)
                stats["athlete_license_fix"] += 1
            # Birthdate: fill in if missing
            if existing["BIRTHDATE"] is None and ins.birthdate is not None:
                updates["BIRTHDATE"] = ins.birthdate
                stats["athlete_birthdate_fix"] += 1
            # Club: update if the athlete moved between meets (rare)
            if existing["CLUBID"] != club_id:
                updates["CLUBID"] = club_id
                stats["athlete_club_fix"] += 1
            if updates:
                db.update("ATHLETE", {"ATHLETEID": aid}, updates)
            continue

        # New athlete
        aid = db.next_id()
        athlete_ids[akey] = aid
        db.insert("ATHLETE", {
            "ATHLETEID":       aid,
            "CLUBID":          club_id,
            "FIRSTNAME":       truncate(ins.first, 30),
            "LASTNAME":        truncate(ins.last, 50),
            "FIRSTNAME_UPPER": upper_key(ins.first, 5),
            "LASTNAME_UPPER":  upper_key(ins.last, 10),
            "GENDER":          new_gender,
            "BIRTHDATE":       ins.birthdate,
            "LICENSE":         truncate(ins.license, 20),
            "NATION":          MEET_NATION,
            "HANDICAPS":       0,
            "HANDICAPSB":      0,
            "HANDICAPSM":      0,
            "SDMSID":          0,
            "SWRID":           0,
        })
        stats["athlete_new"] += 1
        existing_athletes[akey] = {
            "ATHLETEID": aid, "CLUBID": club_id,
            "GENDER": new_gender, "BIRTHDATE": ins.birthdate,
            "LICENSE": ins.license,
        }

    # ----- SWIMRESULT (individual entries) -----
    # Deduplicate (athlete, event-key) pairs from the xlsx; keep the
    # fastest best time when the same athlete appears multiple times.
    best_by: dict[tuple, tuple[int | None, dt.date | None]] = {}
    for akey, ekey, cs in ind_entries:
        ath = athletes[akey]
        bd = ath.birthdate
        cur = best_by.get((akey, ekey))
        if cur is None or (cs is not None and (cur[0] is None or cs < cur[0])):
            best_by[(akey, ekey)] = (cs, bd)

    def _route(targets: list[tuple], athlete_age: int | None,
               fallback_min: int, fallback_max: int) -> tuple:
        """Choose one (eid, agid, amin, amax) from the candidate targets
        for an athlete.  Single target -> trivial.  Multiple targets
        (MM-split) -> pick by athlete age; fall back to first target
        if age is unknown."""
        if len(targets) == 1:
            return targets[0]
        if athlete_age is not None:
            for t in targets:
                lo, hi = _span(t[2], t[3])
                if lo <= athlete_age <= hi:
                    return t
        # No birthdate known — default to the first (e.g. lowest bracket)
        return sorted(targets, key=lambda t: _span(t[2], t[3])[0])[0]

    sr_batch: list[dict] = []
    for (akey, ekey), (cs, bd) in best_by.items():
        aid = athlete_ids[akey]
        ev = events[ekey]
        age_min, age_max, _ = AGE_GROUPS[ev.age_code]
        target = _route(event_targets[ekey], age_at(bd), age_min, age_max)
        eid, ag_id, _, _ = target

        if (aid, eid) in existing_results:
            _sr_id, cur_cs = existing_results[(aid, eid)]
            if cs is not None and (cur_cs is None or cs < cur_cs):
                db.update("SWIMRESULT", {"SWIMRESULTID": _sr_id},
                          {"ENTRYTIME": cs})
                stats["entry_time_faster"] += 1
            continue
        sr_id = db.next_id()
        sr_batch.append({
            "SWIMRESULTID":  sr_id,
            "ATHLETEID":     aid,
            "SWIMEVENTID":   eid,
            "AGEGROUPID":    ag_id,
            "ENTRYTIME":     cs,
            "ENTRYCOURSE":   0,
            "RESULTSTATUS":  0,
            "BONUSENTRY":    "F",
            "DSQNOTIFIED":   "F",
            "FINALFIX":      "F",
            "LATEENTRY":     "F",
            "NOADVANCE":     "F",
            "BACKUPTIME1":   0,
            "BACKUPTIME2":   0,
            "BACKUPTIME3":   0,
            "FINISHJUDGE":   0,
            "PADTIME":       INT_MAX,
            "QTCOURSE":      0,
            "QTTIME":        INT_MAX,
            "QTTIMING":      0,
            "REACTIONTIME":  -32768,
        })
        stats["entry_new"] += 1
        existing_results[(aid, eid)] = (sr_id, cs)
    db.insert_many("SWIMRESULT", sr_batch)

    # ----- RELAY + RELAYPOSITION -----
    # Global TEAMNUMBER / RELAYCODE counter — both must be unique across
    # ALL relays in the meet (not just per club).  SPLASH's seeding module
    # crashes (TBSItem.DoLoadColumns) when two relays in the same event
    # share a TEAMNUMBER/RELAYCODE.  We keep a per-club squad index only
    # for the display name.
    rows = db.query("SELECT COALESCE(MAX(TEAMNUMBER), 0) FROM RELAY")
    next_team_no = int(rows[0][0]) if rows and rows[0][0] is not None else 0

    # Also remember which (club, event, squad-index-within-club) we've seen
    # — that's the *stable* identity across re-runs (TEAMNUMBER isn't,
    # because the global counter shifts if rows were inserted/deleted).
    existing_relays_stable: dict[tuple, int] = {}
    # Build it from the current DB state
    _club_squad_count: dict[tuple, int] = defaultdict(int)
    for rid_row, club_row, event_row, _tn in db.query(
        "SELECT RELAYID, CLUBID, SWIMEVENTID, TEAMNUMBER "
        "FROM RELAY ORDER BY RELAYID"):
        if club_row is None or event_row is None:
            continue
        key_ce = (int(club_row), int(event_row))
        _club_squad_count[key_ce] += 1
        sub_idx = _club_squad_count[key_ce]
        existing_relays_stable[(int(club_row), int(event_row), sub_idx)] = int(rid_row)

    for (cnorm, ekey), members in relay_groups.items():
        ev = events[ekey]
        age_min, age_max, _age_name = AGE_GROUPS[ev.age_code]
        club_id = club_ids[cnorm]
        # Relays don't split by age; pick the first target (relays live on
        # the bracket as a whole, not per-athlete)
        rel_target = event_targets[ekey][0]
        event_id, relay_ag, _, _ = rel_target
        # Chunk members into squads of ev.relay_count.  Any leftover
        # athletes become an additional (incomplete) squad so MM can
        # show them — the coach then picks the final line-up inside MM.
        chunks: list[list[tuple]] = []
        buf: list[tuple] = []
        for m in members:
            buf.append(m)
            if len(buf) == ev.relay_count:
                chunks.append(buf); buf = []
        if buf:
            chunks.append(buf)    # leftover squad (may be <relay_count)

        for club_squad_idx, squad in enumerate(chunks, start=1):
            stable_key = (club_id, event_id, club_squad_idx)
            if stable_key in existing_relays_stable:
                rid = existing_relays_stable[stable_key]
                # Only add any relay positions that don't yet exist.
                for leg_no, (akey, _bt) in enumerate(squad[:ev.relay_count],
                                                     start=1):
                    if (rid, leg_no) in existing_relay_pos:
                        continue
                    db.insert("RELAYPOSITION", {
                        "RELAYID":       rid,
                        "ATHLETEID":     athlete_ids[akey],
                        "RELAYNUMBER":   leg_no,
                        "RESULTSTATUS":  0,
                        "QTCOURSE":      0,
                        "QTISLAP":       "F",
                        "QTTIME":        INT_MAX,
                        "QTTIMING":      0,
                        "REACTIONTIME":  -32768,
                    })
                    existing_relay_pos.add((rid, leg_no))
                    stats["relayposition_new"] += 1
                continue

            # New relay squad — allocate a fresh globally-unique TEAMNUMBER
            next_team_no += 1
            rid = db.next_id()
            entry_time = None
            if (all(bt is not None for _, bt in squad)
                    and len(squad) >= ev.relay_count):
                entry_time = sum(bt for _, bt in squad[:ev.relay_count])
            db.insert("RELAY", {
                "RELAYID":      rid,
                "CLUBID":       club_id,
                "SWIMEVENTID":  event_id,
                "AGEGROUPID":   relay_ag,
                "GENDER":       ev.gender,
                "TEAMNUMBER":   next_team_no,       # globally unique
                "RELAYCODE":    next_team_no,       # globally unique
                "AGEMIN":       age_min,
                "AGEMAX":       age_max,
                "AGETOTAL":     0,
                "ATHLETES":     ev.relay_count,
                "ENTRYTIME":    entry_time,
                "ENTRYCOURSE":  0,
                "RESULTSTATUS": 0,
                "NAME":         truncate(f"{clubs[cnorm]} {club_squad_idx}", 100),
                "BONUSENTRY":   "F",
                "DSQNOTIFIED":  "F",
                "FINALFIX":     "F",
                "LATEENTRY":    "F",
                "NOADVANCE":    "F",
                "BACKUPTIME1":  0,
                "BACKUPTIME2":  0,
                "BACKUPTIME3":  0,
                "FINISHJUDGE":  0,
                "PADTIME":      INT_MAX,
                "QTCOURSE":     0,
                "QTTIME":       INT_MAX,
                "QTTIMING":     0,
                "REACTIONTIME": -32768,
                "USETIMETYPE":  0,
            })
            stats["relay_new"] += 1
            existing_relays_stable[stable_key] = rid
            for leg_no, (akey, _bt) in enumerate(squad[:ev.relay_count],
                                                 start=1):
                db.insert("RELAYPOSITION", {
                    "RELAYID":      rid,
                    "ATHLETEID":    athlete_ids[akey],
                    "RELAYNUMBER":  leg_no,
                    "RESULTSTATUS": 0,
                    "QTCOURSE":     0,
                    "QTISLAP":      "F",
                    "QTTIME":       INT_MAX,
                    "QTTIMING":     0,
                    "REACTIONTIME": -32768,
                })
                existing_relay_pos.add((rid, leg_no))
                stats["relayposition_new"] += 1

    # ----- COMBINEDEVENTS (Cumulatifs) -----
    # On fresh load only — if the BSGLOBAL row already exists we leave it
    # untouched, since the meet director may have tweaked the combined
    # events manually in MM.
    rows = db.query(
        "SELECT DATA FROM BSGLOBAL WHERE NAME='COMBINEDEVENTS'")
    if not rows:
        # Build: age_code -> gender -> uniqueid -> SWIMEVENTID
        # Lookup existing events via (SWIMSTYLEID, GENDER, AGEMIN, AGEMAX)
        # matching the age bracket.
        existing_events_by_sg_age: dict[tuple, int] = {}
        for eid, styid, gen, amin, amax in db.query(
            "SELECT e.SWIMEVENTID, e.SWIMSTYLEID, e.GENDER, "
            "       a.AGEMIN, a.AGEMAX "
            "FROM SWIMEVENT e LEFT JOIN AGEGROUP a "
            "       ON a.SWIMEVENTID=e.SWIMEVENTID"):
            if styid is None or gen is None:
                continue
            existing_events_by_sg_age[(int(styid), int(gen),
                                        int(amin) if amin is not None else None,
                                        int(amax) if amax is not None else None
                                        )] = int(eid)
        # Reverse-map SWIMSTYLEID -> UNIQUEID so we can resolve by UID
        uid_to_sid: dict[int, int] = dict(style_ids)
        # Build cumulatifs
        ce_blocks: list[str] = []
        ce_counter = 0
        for (age_code, gender), (ce_name, uid_list) in CUMULATIFS.items():
            amin, amax, _ = AGE_GROUPS[age_code]
            event_ids_for_ce: list[int] = []
            for uid in uid_list:
                sid = uid_to_sid.get(uid)
                if sid is None:
                    continue
                eid = existing_events_by_sg_age.get((sid, gender, amin, amax))
                if eid is None:
                    # Fall back: any event for this (style, gender), pick first
                    for (s, g, _mn, _mx), e in existing_events_by_sg_age.items():
                        if s == sid and g == gender:
                            eid = e
                            break
                if eid is not None:
                    event_ids_for_ce.append(eid)
            if len(event_ids_for_ce) < 2:
                # Not enough matching events to form a cumulatif; skip
                continue
            ce_counter += 1
            ce_id = db.next_id()
            anchor_eid = event_ids_for_ce[0]
            events_xml = "\n        ".join(
                f'<EVENT eventid="{e}" mandatory="F" />'
                for e in event_ids_for_ce)
            ce_blocks.append(
                f'    <COMBINEDEVENT combinedeventid="{ce_id}" '
                f'name="{ce_name}" titleforprints="{ce_name}" '
                f'sumtype="2" '
                f'pointsforplaces="{CUMULATIF_POINTS}" '
                f'maxresults="100" sortbyresfirst="F" '
                f'penalty="10" inpercent="T" completedsq="F" '
                f'finalusetype="2" agegroupeventid="{anchor_eid}">\n'
                f'      <EVENTS>\n'
                f'        {events_xml}\n'
                f'      </EVENTS>\n'
                f'    </COMBINEDEVENT>')
        if ce_blocks:
            xml = (
                '<?xml version="1.0" encoding="UTF-16"?>\r\n'
                '<COMBINEDEVENTDEFINITION>\r\n'
                '  <COMBINEDEVENTS>\r\n'
                + "\r\n".join(ce_blocks) + "\r\n"
                '  </COMBINEDEVENTS>\r\n'
                '</COMBINEDEVENTDEFINITION>\r\n'
            )
            db.insert("BSGLOBAL", {
                "NAME": "COMBINEDEVENTS",
                "DATA": xml,
            })
            stats["combined_new"] = len(ce_blocks)

    # ----- Summary of changes -----
    print("\n" + "=" * 60)
    print("  Summary of changes")
    print("=" * 60)
    def line(label, n, unit=""):
        if n:
            print(f"  +{n:<5d} {label}{unit}")
    if stats["style_new"]:      line("new SWIMSTYLE (catalog)", stats["style_new"])
    if stats["session_new"]:    line("new SWIMSESSION",         stats["session_new"])
    line("new clubs",                          stats["club_new"])
    line("new athletes",                       stats["athlete_new"])
    line("athlete gender corrections",         stats["athlete_gender_fix"])
    line("athlete license fills",              stats["athlete_license_fix"])
    line("athlete birthdate fills",            stats["athlete_birthdate_fix"])
    line("athlete club changes",               stats["athlete_club_fix"])
    line("new events",                         stats["event_new"])
    line("new age-group rows",                 stats["agegroup_new"])
    line("new individual entries",             stats["entry_new"])
    line("entries updated (faster time)",      stats["entry_time_faster"])
    line("new relay squads",                   stats["relay_new"])
    line("new relay positions",                stats["relayposition_new"])
    line("new combined events (cumulatifs)",   stats["combined_new"])
    total_changes = sum(stats.values())
    if total_changes == 0:
        print("  (no changes — database already in sync with xlsx)")
    print("=" * 60)

    print(f"\nAllocated UIDs {db._start_uid+1}..{db._uid}  "
          f"({db._uid - db._start_uid} new rows)")

    # Emit the issues section (data-quality findings) if any were collected.
    if args.issues_out:
        with open(args.issues_out, "w", encoding="utf-8") as fh:
            # Write the xlsx filename at the top of the output file so the
            # recipient knows which workbook it's from.
            fh.write(f"Data-quality report for: {args.xlsx}\n")
            fh.write(f"Generated: {dt.datetime.now():%Y-%m-%d %H:%M:%S}\n")
            issues.report("Issues found while parsing",
                          out_file=fh, full=True)
        print(f"\n(issues written to {args.issues_out})")
    else:
        issues.report("Issues found while parsing / loading",
                      full=args.issues_full)

    db.commit()
    db.close()
    if args.dry_run:
        print("\nDRY RUN complete — no changes written to disk.")
    else:
        print("\nDone. Open the .mdb in SPLASH Meet Manager to verify.")


if __name__ == "__main__":
    main()
