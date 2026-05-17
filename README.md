# ebimport_splash

[![CI](https://github.com/vrouleau/ebimport_splash/actions/workflows/ci.yml/badge.svg)](https://github.com/vrouleau/ebimport_splash/actions/workflows/ci.yml)

Convert an Eventbrite/registration Excel workbook into a Lenex `.lxf` file for import into **SPLASH Meet Manager 11**.

## How It Works

Upload an xlsx (with an `Attendees` sheet) and a meet `.lxf` (exported from SPLASH). The tool:

1. Parses the xlsx: athletes, clubs, events, entry times, relay squads
2. Validates against the meet structure (event UIDs, age brackets, birthdates)
3. Generates a Lenex `.lxf` ready for SPLASH import via **Transfers → Import entries**

No Java, no MDB, no UCanAccess — pure Python + Lenex XML.

## Features

- **Lenex-only output** — generates `.lxf` compatible with SPLASH import
- **Meet .lxf as template** — uploaded per-request, defines event structure
- **Validation mode** — dry-run without meet .lxf (parse-only, reports issues)
- **Masters routing** — `_MA` LICENSE suffix for VBS-based transfer after prelims
- **Teammate resolver** — fuzzy matching (prefix, reversed names, middle name drop, comma strip)
- **Auto-fix reporting** — `[NOTE]` section shows all teammate name corrections
- **PDF audit** — compare results/heat-sheet PDFs against source xlsx
- **Issues report** — full list in zip download, capped at 10 in UI
- **Docker** — single container, no dependencies beyond Python

## Web UI

Upload xlsx + meet .lxf → validate or generate Lenex → download zip.

```bash
docker compose up --build -d
# Browse http://localhost:5000
```

## Ticket-Type Parser

Recognises French ticket patterns:

```
<age> <gender> <style> [<distance>]    → individual
<age> Relais Mixte <style>             → relay

<age>    ∈ { "15-18", "MA", "Open" }
<gender> ∈ { "F", "M" }
<style>  ∈ { "Corde", "Medley", "Obstacle", "Portage",
             "Remorquage", "Sauveteur d'acier" }
```

Non-race tickets (Banquet, Coach, Cosmodôme, etc.) are silently ignored.

## Masters Transfer (Phase 2)

After prelim heats, Masters athletes (marked with `_MA` LICENSE suffix) are transferred to their dedicated final events using `scripts/masters_transfer.vbs` on Windows.

See [docs/MASTERS_TRANSFER.md](docs/MASTERS_TRANSFER.md).

## Simulate Results

`scripts/simulate_results.vbs` generates random times for testing:
- Skips rows that already have a result (safe to run multiple times)
- DQ entries still get a random time (5% DQ rate)

## Running Tests

```bash
docker compose up --build -d
pip install -r tests/requirements-test.txt
pytest tests/ -v
```

## PDF Audit

Compare a SPLASH results/heat-sheet PDF against the source xlsx:

```bash
curl -sS -X POST http://localhost:5000/api/audit \
  -F pdf=@results.pdf -F xlsx=@input.xlsx
```

## Validation Rules

The tool emits warnings (`[WARNING]`) and notes (`[NOTE]`) in the issues report. These are non-blocking — the Lenex file is still generated.

### Individual Entries

| Category | Severity | Description |
|----------|----------|-------------|
| `no_dob` | WARNING | Athlete has no birthdate |
| `age_bracket_mismatch` | WARNING | Athlete age outside their registered bracket (15-18, Open 19+, Masters 25+) |
| `duplicate_athlete_key` | WARNING | Same athlete key appears more than once |
| `duplicate_entry` | WARNING | Duplicate entry for same athlete+event |

### Relay Rules

| Category | Severity | Description |
|----------|----------|-------------|
| `incomplete_relay` | WARNING | Relay team has fewer members than required (e.g., 3/4) |
| `relay_member_age` | WARNING | Member age outside the relay's age bracket |
| `relay_lower_age` | WARNING | More than 2 members below the agegroup minimum (meet setting: `RELAYTOYOUNGALLOWED=2`) |
| `relay_gender_balance` | WARNING | Quad mixed relay ("Mixte") does not have exactly 2M + 2F |
| `relay_masters_mixing` | WARNING | Masters athlete in an Open relay, or non-Masters (age < 25) in a Masters relay |
| `relay_duo_mixing` | WARNING | Duo relay members from different age categories (15-18, Open, Masters must not mix) |

### Parsing

| Category | Severity | Description |
|----------|----------|-------------|
| `missing_name` | WARNING | Row has no first/last name |
| `missing_ticket` | WARNING | Row has no ticket type |
| `unknown_ticket` | WARNING | Ticket type not recognised by the parser |
| `bad_birthdate` | WARNING | Birthdate could not be parsed |
| `bad_time` | WARNING | Entry time could not be parsed |
| `truncated_name` | WARNING | Name was truncated to fit Lenex limits |

### Auto-Fixes (informational)

| Category | Severity | Description |
|----------|----------|-------------|
| `teammate_auto_fix` | NOTE | Teammate name was fuzzy-matched or corrected automatically |

## Key Files

| File | Purpose |
|------|---------|
| `src/core.py` | Shared classes (IssueCollector, Inscription, TemplateIndex, etc.) |
| `src/load_to_lenex.py` | Main Lenex generator |
| `src/meet_parser.py` | Parse SPLASH meet export .lxf |
| `src/common.py` | Aggregation, validation, teammate resolver |
| `src/audit_pdf.py` | PDF parser for results/heat-sheets |
| `webapp/app.py` | Flask web UI |
| `scripts/` | VBS/BAT for Windows (masters_transfer, simulate_results) |

## Security

No authentication — designed for trusted LAN use. Do not expose publicly without a reverse proxy + auth.

## Licence

Private; no public licence specified.
