"""Shared aggregation and validation logic for ebimport_splash."""
from __future__ import annotations

import re
from collections import Counter, defaultdict
from dataclasses import dataclass
from typing import Any

from core import (
    Inscription, EventKey, IssueCollector,
    norm_key, age_at, TICKET_UID,
    GENDER_MALE, GENDER_FEMALE,
)


@dataclass
class AggregatedData:
    clubs: dict[str, str]                          # norm -> display
    athletes: dict[tuple, Inscription]             # (norm_name, license) -> ins
    name_to_key: dict[str, tuple]                  # norm_name -> canonical key
    events_in_xlsx: dict[tuple, EventKey]           # ek.key() -> EventKey
    ind_entries: list[tuple]                        # [(ath_key, event_key, best_ms)]
    relay_squads: dict[tuple, list[list[tuple]]]   # (club_norm, ekey) -> squads


def aggregate(inscriptions: list[Inscription],
              issues: IssueCollector,
              name_to_dob: dict | None = None) -> AggregatedData:
    """First + second pass: build clubs, athletes, entries, relay squads.

    ``name_to_dob`` (norm_key(first, last) -> datetime) lets teammate-only
    swimmers — those listed in column I but with no entry row — still get a
    birthdate from a related non-race ticket the same person bought.
    """
    if name_to_dob is None:
        name_to_dob = {}
    clubs: dict[str, str] = {}
    athletes: dict[tuple, Inscription] = {}
    events_in_xlsx: dict[tuple, EventKey] = {}
    ind_entries: list[tuple] = []
    relay_squads: dict[tuple, list[list[tuple]]] = defaultdict(list)

    # First pass
    for ins in inscriptions:
        club_norm = norm_key(ins.club)
        clubs.setdefault(club_norm, ins.club)
        ath_key = (norm_key(ins.first, ins.last), ins.license or "")
        if ath_key not in athletes:
            athletes[ath_key] = ins
        else:
            if ins.birthdate and not athletes[ath_key].birthdate:
                athletes[ath_key] = ins
        events_in_xlsx.setdefault(ins.event.key(), ins.event)

    # Name lookup — prefer key with license
    name_to_key: dict[str, tuple] = {}
    for akey, ins in athletes.items():
        nk = norm_key(ins.first, ins.last)
        if nk in name_to_key:
            if akey[1] and not name_to_key[nk][1]:
                name_to_key[nk] = akey
        else:
            name_to_key[nk] = akey

    # Warn about duplicate keys — only when two distinct non-empty licenses clash.
    # Empty license on relay rows (NRAN absent) merging with a real license is normal.
    _name_licenses: dict[str, list[tuple]] = defaultdict(list)
    for akey in athletes:
        _name_licenses[akey[0]].append(akey)
    for nk, akeys in _name_licenses.items():
        if len(akeys) <= 1:
            continue
        real_licenses = {k[1] for k in akeys if k[1]}
        if len(real_licenses) > 1:
            ins = athletes[akeys[0]]
            issues.warn("duplicate_athlete_key",
                f"{ins.first} {ins.last}: conflicting license values "
                f"{sorted(real_licenses)} — merged to one")

    # Second pass
    for ins in inscriptions:
        club_norm = norm_key(ins.club)
        nk = norm_key(ins.first, ins.last)
        ath_key = name_to_key.get(nk, (nk, ins.license or ""))

        if ins.event.is_relay:
            squad: list[tuple] = [(ath_key, ins.best_time_ms)]
            for tname in _parse_teammates(ins.teammates):
                tkey = _resolve_teammate(tname, name_to_key, issues)
                if tkey is not None:
                    if tkey == ath_key:
                        continue
                    squad.append((tkey, athletes[tkey].best_time_ms))
                else:
                    parts = tname.split()
                    tfirst = " ".join(parts[:-1]) if len(parts) >= 2 else tname
                    tlast = parts[-1] if len(parts) >= 2 else ""
                    # Try to recover a DOB from non-race rows the teammate
                    # may have on their own (banquet/coach tickets).
                    supp_dob = (name_to_dob.get(norm_key(tfirst, tlast))
                                or name_to_dob.get(tname))
                    pkey = (tname, "")
                    if pkey not in athletes:
                        athletes[pkey] = Inscription(
                            first=tfirst.title(), last=tlast.title(),
                            email=None, club=ins.club, birthdate=supp_dob,
                            license=None, best_time_ms=None, event=ins.event)
                        name_to_key[tname] = pkey
                    elif supp_dob and not athletes[pkey].birthdate:
                        athletes[pkey].birthdate = supp_dob
                    squad.append((pkey, None))

            squad_sig = frozenset(k for k, _ in squad)
            ekey = ins.event.key()
            existing = relay_squads[(club_norm, ekey)]
            if not any(frozenset(k for k, _ in s) == squad_sig
                       for s in existing):
                existing.append(squad)
        else:
            ind_entries.append((ath_key, ins.event.key(), ins.best_time_ms))

    return AggregatedData(
        clubs=clubs, athletes=athletes, name_to_key=name_to_key,
        events_in_xlsx=events_in_xlsx, ind_entries=ind_entries,
        relay_squads=relay_squads,
    )


def run_sanity_checks(template: Any) -> list[str]:
    """Return list of fatal errors if template is incompatible. Empty = OK."""
    fatals = []
    missing_uids = [uid for uid in set(TICKET_UID.values())
                    if uid not in template.styles_by_uid]
    if missing_uids:
        fatals.append(
            f"TICKET_UID references UIDs not in template SWIMSTYLE: "
            f"{sorted(missing_uids)}")
    return fatals


def run_validation(events_in_xlsx: dict[tuple, EventKey],
                   template: Any) -> list[str]:
    """Validate xlsx events against template structure. Returns fatal errors."""
    fatal: list[str] = []
    for ek, ev in events_in_xlsx.items():
        style = template.styles_by_uid.get(ev.uniqueid)
        if style is None:
            fatal.append(
                f"Ticket {ev.label}: template has no SWIMSTYLE "
                f"with UNIQUEID={ev.uniqueid}")
            continue
        tevent = template.find_event(ev.uniqueid, ev.gender,
                                      masters=(ev.age_code == "MASTERS"))
        if tevent is None:
            fatal.append(
                f"Ticket {ev.label}: no SWIMEVENT with gender={ev.gender} "
                f"for age code {ev.age_code!r}")
            continue
        if ev.is_relay:
            if ev.age_code in ("1518", "OPEN"):
                need_min = 15 if ev.age_code == "1518" else 19
                if not any(a.amin == need_min for a in tevent.agegroups):
                    fatal.append(
                        f"Ticket {ev.label}: SWIMEVENT #{tevent.event_number} "
                        f"has no AGEGROUP for bracket {ev.age_code}")
            elif ev.age_code == "MASTERS":
                has_any = any(a.amin is not None and a.amin >= 25
                              for a in tevent.agegroups)
                if not has_any:
                    fatal.append(
                        f"Ticket {ev.label}: Masters relay but no Masters AGEGROUPs")
        else:
            if ev.age_code == "1518":
                if not any(a.amin == 15 and a.amax == 18
                            for a in tevent.agegroups):
                    fatal.append(
                        f"Ticket {ev.label}: no 15-18 AGEGROUP")
            elif ev.age_code == "OPEN":
                if not any(a.amin == 19 for a in tevent.agegroups):
                    fatal.append(
                        f"Ticket {ev.label}: no 19+ AGEGROUP")
            elif ev.age_code == "MASTERS":
                if not any(a.amin is not None and 25 <= a.amin < 100
                            for a in tevent.agegroups):
                    fatal.append(
                        f"Ticket {ev.label}: no 5-year Masters AGEGROUPs")
    return fatal


def run_cross_row_checks(data: AggregatedData, template: Any,
                         issues: IssueCollector) -> None:
    """Emit warnings for data-quality issues across rows."""
    athletes = data.athletes
    events_in_xlsx = data.events_in_xlsx
    relay_squads = data.relay_squads
    clubs = data.clubs

    # No DOB
    for akey, ins in athletes.items():
        if ins.birthdate is None:
            issues.warn("no_dob",
                f"{ins.first} {ins.last} ({ins.club}) has no birthdate")

    # Individual age bracket mismatch
    for akey, ekey, _ in data.ind_entries:
        ev = events_in_xlsx[ekey]
        ins = athletes[akey]
        age = age_at(ins.birthdate)
        if age is None:
            continue
        ac = ev.age_code
        if ac == "1518" and not (15 <= age <= 18):
            issues.warn("age_bracket_mismatch",
                f"{ins.first} {ins.last} age {age} outside 15-18 bracket")
        elif ac == "OPEN" and age < 19:
            issues.warn("age_bracket_mismatch",
                f"{ins.first} {ins.last} age {age} too young for Open (19+)")
        elif ac == "MASTERS" and age < 25:
            issues.warn("age_bracket_mismatch",
                f"{ins.first} {ins.last} age {age} too young for Masters (25+)")

    # Relay member checks
    for (cnorm, ekey), squads in relay_squads.items():
        ev = events_in_xlsx[ekey]
        style = template.styles_by_uid.get(ev.uniqueid)
        if style is None:
            continue
        relay_size = style.relay_count or 4
        for squad in squads:
            if len(squad) < relay_size:
                first_ath = athletes[squad[0][0]]
                issues.warn("incomplete_relay",
                    f"{clubs[cnorm]}: {len(squad)}/{relay_size} athletes "
                    f"for UID {ev.uniqueid} ({ev.age_code}) "
                    f"— registrant: {first_ath.first} {first_ath.last}")
            for akey, _ in squad[:relay_size]:
                member = athletes[akey]
                m_age = age_at(member.birthdate)
                if m_age is not None:
                    if ev.age_code == "1518" and not (15 <= m_age <= 18):
                        issues.warn("relay_member_age",
                            f"{member.first} {member.last} age {m_age} "
                            f"in 15-18 relay ({clubs[cnorm]})")
                    elif ev.age_code == "OPEN" and m_age < 15:
                        issues.warn("relay_member_age",
                            f"{member.first} {member.last} age {m_age} "
                            f"too young for Open relay ({clubs[cnorm]})")

            # --- Relay rule warnings (non-blocking) ---
            members = squad[:relay_size]
            relay_label = "/".join(athletes[ak].last for ak, _ in members)

            # 1) Lower-age limit (meet setting: max 2 below agegroup min)
            _RELAY_LOWER_AGE_ALLOWED = 2
            if ev.age_code == "1518":
                _age_floor = 15
            elif ev.age_code in ("OPEN", "MASTERS"):
                _age_floor = 19
            else:
                _age_floor = None
            if _age_floor is not None:
                below = [athletes[ak] for ak, _ in members
                         if age_at(athletes[ak].birthdate) is not None
                         and age_at(athletes[ak].birthdate) < _age_floor]
                if len(below) > _RELAY_LOWER_AGE_ALLOWED:
                    names = ", ".join(f"{a.first} {a.last}" for a in below)
                    issues.warn("relay_lower_age",
                        f"{relay_label} ({clubs[cnorm]}): "
                        f"{len(below)} below age {_age_floor} "
                        f"(max {_RELAY_LOWER_AGE_ALLOWED}): {names}")

            # 2) Quad mixed relay gender balance: must be 2M + 2F
            if relay_size == 4 and style.name and "mixte" in style.name.lower():
                genders = []
                for ak, _ in members:
                    # Derive gender from athlete's gendered individual entry
                    for a2, ek2, _ in data.ind_entries:
                        if a2 == ak:
                            g = events_in_xlsx[ek2].gender
                            if g in (GENDER_MALE, GENDER_FEMALE):
                                genders.append(g)
                                break
                if len(genders) == 4:
                    m_cnt = genders.count(GENDER_MALE)
                    f_cnt = genders.count(GENDER_FEMALE)
                    if m_cnt != 2 or f_cnt != 2:
                        issues.warn("relay_gender_balance",
                            f"{relay_label} ({clubs[cnorm]}): "
                            f"{m_cnt}M+{f_cnt}F (must be 2M+2F)")

            # 3) Masters/Open mixing
            if ev.age_code == "MASTERS":
                non_ma = [athletes[ak] for ak, _ in members
                          if not any(events_in_xlsx[ek2].age_code == "MASTERS"
                                     for a2, ek2, _ in data.ind_entries
                                     if a2 == ak and not events_in_xlsx[ek2].is_relay)]
                if non_ma:
                    names = ", ".join(f"{a.first} {a.last}" for a in non_ma)
                    issues.warn("relay_masters_mixing",
                        f"{relay_label} ({clubs[cnorm]}): "
                        f"non-Masters in Masters relay: {names}")
            elif ev.age_code == "OPEN":
                ma_in_open = [athletes[ak] for ak, _ in members
                              if any(events_in_xlsx[ek2].age_code == "MASTERS"
                                     for a2, ek2, _ in data.ind_entries
                                     if a2 == ak)]
                if ma_in_open:
                    names = ", ".join(f"{a.first} {a.last}" for a in ma_in_open)
                    issues.warn("relay_masters_mixing",
                        f"{relay_label} ({clubs[cnorm]}): "
                        f"Masters athlete(s) in Open relay: {names}")

            # 4) Duo relay: cannot mix age groups
            if relay_size == 2:
                codes = []
                for ak, _ in members:
                    # Use the age_code from the athlete's individual entries
                    ind_codes = {events_in_xlsx[ek2].age_code
                                 for a2, ek2, _ in data.ind_entries
                                 if a2 == ak and not events_in_xlsx[ek2].is_relay}
                    if "MASTERS" in ind_codes:
                        codes.append("MASTERS")
                    elif "OPEN" in ind_codes:
                        codes.append("OPEN")
                    elif "1518" in ind_codes:
                        codes.append("1518")
                    else:
                        codes.append(None)
                if all(c is not None for c in codes) and codes[0] != codes[1]:
                    m0, m1 = athletes[members[0][0]], athletes[members[1][0]]
                    issues.warn("relay_duo_mixing",
                        f"{relay_label} ({clubs[cnorm]}): "
                        f"mixed age groups — "
                        f"{m0.first} {m0.last} ({codes[0]}) + "
                        f"{m1.first} {m1.last} ({codes[1]})")


# --------------------------------------------------------------------------- #
# Internal helpers
# --------------------------------------------------------------------------- #
def _parse_teammates(raw: str | None) -> list[str]:
    if not raw:
        return []
    names = []
    for line in raw.split("\n"):
        line = line.strip()
        if not line or re.match(r"^\(.*\)$", line):
            continue
        tokens = [t.strip(",") for t in line.split()]
        while len(tokens) > 2:
            last = tokens[-1]
            if re.match(r"^[A-Z0-9]{3,8}$", last):
                tokens.pop()
            elif re.match(r"^\d+$", last):
                tokens.pop()
            elif last.lower() in ("years", "old", "ans"):
                tokens.pop()
            else:
                break
        names.append(norm_key(" ".join(tokens)))
    return names


def _resolve_teammate(name_norm: str, name_to_key: dict, issues=None) -> tuple | None:
    if name_norm in name_to_key:
        return name_to_key[name_norm]
    tokens = name_norm.split()
    orig_tokens = list(tokens)
    while len(tokens) > 2:
        tokens.pop()
        attempt = " ".join(tokens)
        if attempt in name_to_key:
            if issues:
                issues.note("teammate_auto_fix",
                    f"'{name_norm}' -> '{attempt}' (trimmed trailing tokens)")
            return name_to_key[attempt]
    # Prefix match: "phil skinder" -> "philip skinder"
    if len(orig_tokens) == 2:
        first, last = orig_tokens
        for key in name_to_key:
            parts = key.split()
            if len(parts) >= 2 and parts[-1] == last and parts[0].startswith(first):
                if issues:
                    issues.note("teammate_auto_fix",
                        f"'{name_norm}' -> '{key}' (prefix match)")
                return name_to_key[key]
    # First+last fallback: "luis ismail gana-akkor" -> "luis gana-akkor"
    if len(orig_tokens) >= 3:
        first_last = f"{orig_tokens[0]} {orig_tokens[-1]}"
        if first_last in name_to_key:
            if issues:
                issues.note("teammate_auto_fix",
                    f"'{name_norm}' -> '{first_last}' (dropped middle name)")
            return name_to_key[first_last]
    # Reversed name: "barter ying" -> "ying barter"
    if len(orig_tokens) == 2:
        reversed_name = f"{orig_tokens[1]} {orig_tokens[0]}"
        if reversed_name in name_to_key:
            if issues:
                issues.note("teammate_auto_fix",
                    f"'{name_norm}' -> '{reversed_name}' (reversed name)")
            return name_to_key[reversed_name]
    return None
