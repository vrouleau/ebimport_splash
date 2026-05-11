# ebimport_splash

Python tool + Flask web app that converts a JotForm registration spreadsheet (xlsx) into a Lenex 3.0 .lxf file, ready to load into SPLASH for a lifesaving meet.

## What it does

1. Reads a JotForm "Attendees" xlsx export with registration rows (athlete, club, events, age category, best times).
2. Validates entries against the SPLASH meet template (event structure, age groups, relay constraints).
3. Outputs a Lenex `.lxf` file for import through Meet Manager's standard Lenex import (`load_to_lenex.py`).

The web UI (`webapp/app.py`) wraps the Lenex path as a stateless Flask app.

## Repo layout

```
ebimport_splash/
├── src/
│   ├── core.py             # Main logic: read xlsx, validate, route age groups, build inscriptions
│   ├── load_to_lenex.py    # CLI: xlsx + meet .lxf → output .lxf
│   ├── meet_parser.py      # Parse SPLASH meet .lxf → ParsedMeet (shared with meetmanager-app)
│   ├── common.py           # Shared validation, sanity checks
│   └── audit_pdf.py        # Generate audit PDF from issues
├── webapp/
│   └── app.py              # Flask web app wrapping load_to_lenex
├── scripts/
│   ├── simulate_results.bat / .vbs   # Windows scripts to seed fake results in SPLASH
│   ├── masters_transfer.bat / .vbs   # Transfer Masters athletes between databases
│   └── audit.bat / format_audit.py  # Audit report generation
├── forms/
│   └── jotform_inscription.json     # JotForm form definition (for reference)
├── tests/
│   ├── test_integration.py
│   ├── test_attendees.xlsx          # Integration test fixture
│   └── build_meet_fixture.py        # Generate test meet .lxf
├── docs/
│   └── MASTERS_TRANSFER.md
├── Dockerfile
└── docker-compose.yml
```

## Key source files

### `src/core.py`
Central module. Contains:
- `read_attendees(xlsx)` — parse the JotForm xlsx into `Inscription` dataclasses
- `IssueCollector` — accumulate WARNING / NOTE issues; surfaced in the output report
- `TemplateStyle`, `TemplateAgeGroup`, `TemplateEvent` — dataclasses used by `MeetLxfTemplate`
- `MeetLxfTemplate` (in `load_to_lenex.py`) — wraps a `ParsedMeet` for event/age-group lookup
- `pick_agegroup_for_individual(age, event_codes, template)` — route individual entry to correct age group
- `pick_agegroup_for_relay(ages, template)` — route relay by sum-of-ages for Masters relays
- Age/gender constants: `GENDER_MALE=1, GENDER_FEMALE=2, GENDER_ALL=0, GENDER_MIXED=3`
- Round constants: `ROUND_TIMED_FINAL=1, ROUND_PRELIM=2, ROUND_FINAL=9`
- `AGE_DATE` — global, set by `load_to_lenex.py` (defaults to 2026-12-31 if not overridden)

### `src/meet_parser.py`
Parses a SPLASH-exported meet `.lxf` (zip containing `.lef` XML) into:
- `ParsedMeet` — meet name, course, masters flag, `meet_fees: dict[str,int]`, currency, sessions
- `MeetSession` — session number, name, list of `MeetEvent`
- `MeetEvent` — eventid, number, gender, round, swimstyleid, distance, relaycount, style_name, fee_cents, agegroups

**This file is shared with `meetmanager-app/backend/app/meet_parser.py` — keep them in sync.**

### `src/load_to_lenex.py`
CLI entry point for the Lenex output path:
```bash
python load_to_lenex.py --xlsx CPLC2026FINAL.xlsx --meet splash_results_meet.lxf --out meet.lxf
```
Uses `MeetLxfTemplate` (adapter wrapping `ParsedMeet`) to provide the same interface as `TemplateIndex`.
Produces two output zips: `splash-inscription.zip` (entries) and `splash-dryrun.zip` (dry-run preview).

### `webapp/app.py`
Stateless Flask app:
- Accepts xlsx + meet .lxf upload
- Runs `load_to_lenex.py` in a subprocess (dry-run or write mode)
- Parses stdout for Summary + Issues sections
- Serves resulting zip + issues report as download
- Temp dirs auto-cleaned after 10 min (`STAGING_TTL_SECS`)
- Port 5000, deployed via Docker

## Age-bracket routing

| xlsx ticket | Routes to |
|---|---|
| `15-18` | AGEGROUP [15, 18] |
| `Open` | AGEGROUP [19, 99] |
| `Masters` individual | 5-year bracket containing athlete age at `AGE_DATE` |
| `Masters` relay | Sum-of-ages bracket containing squad's total age |

## Masters athletes

- Individual Masters: routed to 5-year age brackets in SPLASH
- Masters relay: routed by sum of team members' ages
- Identified in Lenex output with `HANDICAP exception=X` attribute
- `AGE_DATE` defaults to 2026-12-31; override by setting `core.AGE_DATE` before calling

## Exported zip naming

- `splash-inscription.zip` — real entry output
- `splash-dryrun.zip` — dry-run preview (no writes)

## Running locally

```bash
# Web app
docker compose up --build
# Available at http://localhost:5000

# CLI (Lenex path)
python src/load_to_lenex.py --xlsx registrations.xlsx --meet meet.lxf --out output.lxf
```

## Testing

```bash
cd tests && pip install -r requirements-test.txt && pytest test_integration.py
```

## Environment

- Nation: hardcoded `CAN` in `core.py`
- Build timestamp injected via Docker ARG `BUILD_TIMESTAMP`
