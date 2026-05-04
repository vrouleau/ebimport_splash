# ebimport_splash

Convert an Eventbrite/registration Excel workbook into inscriptions
loaded against an existing **SPLASH Meet Manager 11** meet database.
Two scripts are provided:

- **`load_to_mdb.py`** — writes directly into a SPLASH `.mdb` file
  (via Jackcess/UCanAccess over JDBC). Idempotent / re-runnable.
- **`load_to_lenex.py`** — emits a Lenex 3.0 `.lef` (or zipped `.lxf`)
  file that you can import into SPLASH Meet Manager or any other
  Lenex-compatible tool.

Both scripts read the **`Attendees`** sheet of a registration workbook
(one row per athlete × ticket/event).

## How the MDB loader works

The supplied `.mdb` is the **authoritative event template**.  The meet
organiser has set up the event structure in SPLASH (styles, events,
age groups, sessions, combined events) ahead of time; the loader
treats that structure as read-only and only populates:

- **CLUB** rows — one per distinct club in the xlsx
- **ATHLETE** rows — one per distinct (first name, last name, NRAN)
- **SWIMRESULT** rows — one individual entry per (athlete, event),
  with entry time in milliseconds
- **RELAY** + **RELAYPOSITION** rows — relay squads for mixed-gender
  relay events, routed by template age bracket (or by age-sum for
  Masters relays)

The loader **never** creates SWIMSTYLE, SWIMEVENT, AGEGROUP,
SWIMSESSION, or COMBINEDEVENTS rows.  If a ticket in the xlsx doesn't
resolve to an existing (UNIQUEID, gender, age-bracket) combination in
the template, it's reported as a **fatal error** and the import is
aborted with no writes performed — fix the xlsx (or add the missing
event in SPLASH) and re-run.

**First run vs. re-run** is auto-detected: if the supplied `.mdb`
contains zero SWIMRESULT/RELAY rows it's considered a first run,
otherwise additive mode kicks in.  On a re-run, only missing rows are
inserted; `ENTRYTIME` is updated only when a faster time is supplied
(never regresses).

## Age-bracket routing

| xlsx age prefix | Routed to AGEGROUP |
|---|---|
| `15-18` | the bracket `[15, 18]` on the matched SWIMEVENT |
| `Open`  | the bracket `[19, 99]` on the matched SWIMEVENT |
| `MA` (individual) | the 5-year Masters bracket containing the athlete's age at `AGE_DATE` |
| `MA Relais Mixte` | the age-sum Masters bracket containing the squad's total age |

If a Masters athlete has no birthdate, their individual entry is
warned and skipped.  A Masters relay squad where any member lacks a
DOB is skipped entirely.

---

## Running the web UI (recommended)

The supported way to run this is as a Docker container, which bundles
everything (Python runtime, Java, UCanAccess, the loaders, the baked-in
empty SPLASH template) into a single image.  See the
[Web UI (Docker)](#web-ui-docker) section near the bottom.

## Running from source (development)

If you want to hack on the loaders directly:

### Python

Python 3.10+ with:

- `openpyxl` (both scripts)
- `jaydebeapi` + `JPype1` (MDB script only)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install openpyxl jaydebeapi JPype1
```

### UCanAccess (MDB script only)

The MDB writer speaks to Access via **UCanAccess** over JDBC. Download
the bundle and unpack it somewhere, e.g.:

```bash
curl -sSL -A 'Mozilla/5.0' \
    -o /tmp/ucanaccess.zip \
    'https://downloads.sourceforge.net/project/ucanaccess/UCanAccess-5.0.1.bin.zip'
unzip /tmp/ucanaccess.zip -d /tmp/ucanaccess
export UCANACCESS_DIR=/tmp/ucanaccess/UCanAccess-5.0.1.bin
```

(SourceForge serves an HTML interstitial that needs the `-A` header
to be bypassed; if the command above still returns HTML, follow the
meta-refresh URL manually.)

Requires Java 8+ on the `PATH`.

> On Windows you can instead replace the UCanAccess connection inside
> `MDB.__init__` with a `pyodbc` call using the built-in MS Access
> ODBC driver — the rest of the code is portable.

---

## Usage

```bash
# --- MDB: fresh load or additive update ---
python load_to_mdb.py --xlsx CPLC2026FINAL.xlsx --mdb Canadien.mdb
python load_to_mdb.py --xlsx CPLC2026FINAL.xlsx --mdb Canadien.mdb --dry-run
python load_to_mdb.py --xlsx CPLC2026FINAL.xlsx --mdb Canadien.mdb --wipe

# --- Lenex: produce a .lef or a zipped .lxf ---
python load_to_lenex.py --xlsx CPLC2026FINAL.xlsx --out candien.lef
python load_to_lenex.py --xlsx CPLC2026FINAL.xlsx --out candien.lxf --zip
```

### MDB flags

| Flag | What it does |
|---|---|
| `--xlsx PATH` | Excel workbook with an `Attendees` sheet (required) |
| `--mdb PATH` | Target `.mdb` file, will be modified in place (required) |
| `--dry-run` | Parse and plan everything, then roll back — no writes |
| `--wipe` | Delete existing clubs/athletes/events/entries before loading |

---

## Re-runnable (additive) behaviour

The MDB script is designed to be **re-run safely** after the meet
director has started organising the meet inside SPLASH. On every run
it:

- detects existing rows via stable external keys
  (club norm-name, athlete first+last+license, UNIQUEID + gender +
  age bracket, …)
- inserts only what's missing
- updates an existing entry's `ENTRYTIME` only when the new time
  is **faster** (never regresses)
- fills in athlete gender / license / birthdate when they were
  missing before
- handles **AGEGROUP splits** done in SPLASH: if you split
  `15-18` into `15-16 + 17-18`, the script recognises the split and
  routes new entries to the correct sub-event based on the athlete's
  age (computed at `AGE_DATE`)
- **never overwrites** `HEATID`, `LANE`, `SWIMTIME`, session
  assignments, event numbers, custom `FEE` / `ROUNDNAME`, or the
  `BSGLOBAL.COMBINEDEVENTS` block

Each run ends with a **summary of changes** (`+3 clubs / +12 athletes
/ +47 entries / +2 entries updated (faster time) / …`) followed by an
**issues section** flagging data-quality problems (unknown ticket
types, unparseable times/birthdates, athletes out of bracket,
incomplete relays, …).

---

## Editing the ticket → UNIQUEID map

The `TICKET_UID` dict near the top of `load_to_mdb.py` maps the xlsx
ticket label + (is_relay, is_masters_obstacle) to the SWIMSTYLE
UNIQUEID expected in the template.  The values below match the
**Championnats canadiens 2026** template:

| Ticket label | UNIQUEID | Template event name |
|---|---|---|
| Corde              | 504 | 12 m Lancer de la corde / Line Throw |
| Obstacle (15-18/Open) | 501 | 200 m Nage avec obstacles / Obstacle Swim |
| Obstacle 100 m (Masters) | 541 | 100 m Nage avec obstacles / Obstacle Swim |
| Portage (100 m)    | 502 | 100 m Portage Mannequin palmes |
| Portage 50 m       | 507 | 50 m Portage du mannequin plein |
| Remorquage         | 506 | 100 m Remorquage mannequin palmes |
| Sauveteur d'acier  | 508 | 200 m Sauveteur d'acier / Super Lifesaver |
| Medley             | 531 | 100 m Sauvetage combiné / Rescue Medley |
| Relais Medley      | 544 | 4 × 50 m Relais mixte sauve combiné |
| Relais Obstacle    | 542 | 4 × 50 m Relais obstacle mixte |
| Relais Portage     | 543 | 2 × 50 m Relais mixte portage |

Editing this dict is enough to adapt the loader to a different season
or federation — the matching SWIMSTYLE UIDs in the template `.mdb`
are the authority.

---

## Ticket-type parser

The parser recognises these French ticket name patterns:

```
<age> <gender> <style> [<distance> m]         individual
<age> Relais Mixte <style>                    relay

<age>    ∈ { "15-18", "MA", "Open" }
<gender> ∈ { "F", "M" }
<style>  ∈ { "Corde", "Medley", "Obstacle", "Portage",
             "Remorquage", "Sauveteur d'acier" }
```

Non-race tickets prefixed with `Banquet`, `Coach`, `Cosmod`,
`Couloir`, `Officiel`, `Priorit`, `Sheraton` are silently ignored.
Everything else is reported as `[WARNING] unknown_ticket` with the
xlsx row number.

---

## Caveats & known limitations

- The MDB writer is tested on UCanAccess **5.0.1** / Jackcess 3.0.1
  against a SPLASH 11 MEET database template. Your SPLASH version
  may have schema additions that these scripts don't touch but also
  don't break.
- All generated events default to a single **placeholder session**
  you rename/split in SPLASH afterwards.
- Best times are stored in **milliseconds**
  (`ENTRYTIME` = total ms). Parser accepts `mm:ss.cc`, `hh:mm:ss.cc`,
  `ss.cc`, as well as Excel time/timedelta cell types.
- Lifesaving strokes require `STROKE=0`, `TECHNIQUE=0`
  (federation catalog). Other values crash SPLASH's result module
  at `TBSwLanguage.StrokeName` or `TFModulResult.ECEventTreeNodeChange`.
  The scripts also populate every boolean flag (`'F'/'T'`) SPLASH
  reads from `SWIMEVENT` / `AGEGROUP` / `SWIMRESULT` / `RELAY` to
  avoid those crashes.

---

## Licence

Private; no public licence specified yet. Do not distribute the
generated `.mdb` files (they may contain personal data).

---

## Web UI (Docker)

A minimal Flask web app in `webapp/` wraps both loaders behind a
single-page French UI.  Upload an xlsx, pick a mode (dry-run / MDB /
Lenex), see summary + issues in the page, download a zip with the
generated file + issues report.

### Build + run locally

```bash
docker build -t ebimport-splash:latest .
docker run --rm -p 5000:5000 ebimport-splash:latest
# Browse http://localhost:5000
```

### Deployment via Portainer

On the target host (e.g. `192.168.1.190`), in Portainer (`:9000`):

1. **Stacks → Add stack**
2. Name: `ebimport-splash`
3. **Web editor** (or **Repository** pointing at this repo), paste the
   content of `docker-compose.yml`.
4. **Deploy the stack**.

The container exposes port `5000` on the host.  Browse to
`http://192.168.1.190:5000` on the LAN.

No persistent volumes are needed — uploads live in `/tmp/ebimport_staging/`
inside the container and are cleaned up on download or after a 30-min
TTL.

### Security note

The app has **no authentication** and is designed for use on a trusted
LAN. Do not expose it on the public internet without a reverse proxy
with a password (e.g. nginx/Traefik with basic auth).  Uploaded xlsx
files contain PII (athlete names, birthdates, emails).
