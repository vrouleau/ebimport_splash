# Expected output of manual regression against `test_attendees.xlsx`

`test_attendees.xlsx` is a **deterministic** synthetic workbook produced
by `generate_test_xlsx.py` (seeded RNG `20260504`). Each run should
reproduce the same values. Use this file as the reference for manual
regression of `load_to_mdb.py`.

Regenerate the xlsx any time with:

```bash
python tests/generate_test_xlsx.py --out tests/test_attendees.xlsx
```

---

## Composition

- **100 base athletes** across 5 clubs of sizes **30 / 25 / 20 / 15 / 10**
  (Aurora Test Club, Béluga Sauvetage, Cedar Creek LSC, Dauphins de
  l'Est, Elite Rescue).
- Age-bracket mix roughly **40 % 15-18 / 40 % Open / 20 % Masters**,
  with seeded guarantees so the relay tests have the right sizes.
- Individual tickets for every one of the **7 styles × 3 age brackets
  × 2 genders** (42 combinations) — at least one athlete per combo.
- Mixed relays in each of the **3 relay styles × 3 age brackets** (9
  combos), with three squad sizes exercised:
  - full squad of 4 (most clubs)
  - **under-subscribed** squad of 2 (Cedar Creek LSC 15-18)
  - **over-subscribed** squad of 5 (Béluga Sauvetage Open — triggers
    `extra_relay_members`)
- **8 non-race tickets** (Banquet, Banquet Officiel, Coach, Cosmodôme,
  Couloir de nage, Officiel 3 jours, Priorité - SERC, Sheraton).
- **6 injected defect rows** (one per defect type) plus the existing
  `under-age` / `over-age` / `duplicate` / `no-DOB` athletes.

---

## Expected counts (MDB fresh load)

```
===== Summary of changes =====
  +11    new SWIMSTYLE (catalog)
  +1     new SWIMSESSION
  +5     new clubs
  +114   new athletes               # 100 base + 14 injected defect/fuzzy athletes
  +51    new events                  # 42 ind + 9 relay
  +51    new age-group rows
  +368   new individual entries
  +18    new relay squads
  +66    new relay positions
  +6     new combined events (cumulatifs)
```

---

## Expected Issues section

The loader should emit these exact categories on a fresh run (order
may vary, counts are stable):

| Severity | Category | Count | Trigger |
|---|---|---|---|
| WARNING | `no_dob` | 3 | Beth BadDOB, Nora NoDOB, Noel NoDOB |
| WARNING | `incomplete_relay` | 3 | Cedar Creek LSC 15-18 × 3 relay styles (2/4 athletes) |
| WARNING | `age_bracket_mismatch` | 2 | Under AgeTooYoung (13) + Over AgeTooOld (20), both in 15-18 ticket |
| WARNING | `possible_duplicate_club` | 2 | `"Aurora Test  Club"` (double space) vs `"Aurora Test Club"`; `"Beluga Sauvetage"` (no accent) vs `"Béluga Sauvetage"` |
| WARNING | `unknown_ticket` | 1 | Zach Unknown `"Not A Real Ticket"` |
| WARNING | `missing_name` | 1 | row with no first name |
| WARNING | `bad_time` | 1 | Bob BadTime `"not-a-time"` |
| WARNING | `bad_birthdate` | 1 | Beth BadDOB `"maybe 2001?"` |
| WARNING | `license_name_mismatch` | 1 | `CHIU_SAME` license on "Henri Chiu" and "Henri Tsz Hin Chiu" |
| WARNING | `possible_duplicate_athlete` | 1 | "Stephen Kennedy" vs "Stphen Kennedy" in Elite Rescue (similarity 0.97) |
| WARNING | `same_person_different_club` | 1 | "Gabrielle Fortin" born 2000-03-03 in both Dauphins de l'Est and Elite Rescue |
| NOTE    | `extra_relay_members` | 3 | Béluga Sauvetage Open × 3 relay styles (5 athletes; 1 extra tucked onto last squad) |
| NOTE    | `duplicate_entry` | 1 | Héloïse Lavoie entered in `Open F Obstacle` twice |
| NOTE    | `non_race_only_athlete` | 1 | the Sheraton-only attendee |

Notes:

- Beth BadDOB is reported **twice** — once as `bad_birthdate` (the
  `"maybe 2001?"` string) and once as `no_dob` (because the parse
  failed so she effectively has no usable birthdate).
- The over-subscribed Béluga squad contributes to `relay positions`
  with 5 × 3 = 15 positions across the 3 relay styles
  (vs the 4 × 3 = 12 a fully balanced club would contribute).

---

## Expected re-run (additive) behaviour

Running the MDB loader **a second time** against the same produced
`.mdb` should report:

```
===== Summary of changes =====
  (no changes — database already in sync with xlsx)
```

and **`0 new rows`** allocated.

Running it against a version of the xlsx with the Héloïse Lavoie
duplicate row's second time **faster than the first** (not the
default test data — that would be a separate scenario) should
report `+1 entries updated (faster time)` on re-run.

---

## Verified on

- Python 3.12
- UCanAccess 5.0.1
- commit/tag: *initial release* (see git log)
