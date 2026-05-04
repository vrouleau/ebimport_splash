# ebimport_splash

Convert an Eventbrite/registration Excel workbook into a **SPLASH Meet
Manager 11** meet. Two independent scripts are provided:

- **`load_to_mdb.py`** — writes directly into a SPLASH `.mdb` file
  (via Jackcess/UCanAccess over JDBC). Idempotent / re-runnable.
- **`load_to_lenex.py`** — emits a Lenex 3.0 `.lef` (or zipped `.lxf`)
  file that you can import into SPLASH Meet Manager or any other
  Lenex-compatible tool.

Both scripts read the **`Attendees`** sheet of a registration workbook
(one row per athlete × ticket/event) and produce a meet populated with:

- clubs, athletes
- one `SWIMEVENT` per (age bracket × gender × style) with its `AGEGROUP`
- individual entries (`SWIMRESULT` rows) with entry time in
  hundredths of a second
- relay squads (`RELAY` + `RELAYPOSITION`), chunked 4 members at a time
- `SWIMSTYLE` rows following the **Société de Sauvetage** lifesaving
  catalog (`STROKE=0`, `TECHNIQUE=0`, `UNIQUEID` 501–552, French names)
- combined events (Cumulatifs) written to
  `BSGLOBAL.COMBINEDEVENTS` with the federation's
  `pointsforplaces="20,18,16,14,13,12,11,10,8,7,6,5,4,3,2,1"`
  point schedule (MDB only)

---

## Requirements

### Python

Python 3.10+ with:

- `openpyxl` (both scripts)
- `jaydebeapi` + `JPype1` (MDB script only)

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install openpyxl jaydebeapi
```

### UCanAccess (MDB script only)

The MDB writer speaks to Access via **UCanAccess** over JDBC. Download
the bundle and unpack it somewhere, e.g.:

```bash
wget -O /tmp/ucanaccess.zip \
  "https://sourceforge.net/projects/ucanaccess/files/latest/download"
unzip /tmp/ucanaccess.zip -d /tmp/ucanaccess
export UCANACCESS_DIR=/tmp/ucanaccess/UCanAccess-5.0.1.bin
```

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

## Editing event definitions

All lifesaving events live in the `LIFESAVING_CATALOG` dict near the
top of each script. Event structure follows this identity:

```
(age bracket, gender, catalog UID) -> SWIMEVENT + AGEGROUP
```

The Société-de-Sauvetage catalog UIDs used (matches the
`30-Deux 25 octobre 2025.mdb` reference):

| UID | Event |
|---|---|
| 501 | 200 m Nage avec obstacles |
| 502 | 100 m Portage Mannequin plein avec palmes |
| 504 | 12 m Lancer de la corde |
| 506 | 100 m Remorquage du mannequin ½ plein + palmes |
| 507 | 50 m Portage du mannequin plein |
| 508 | 200 m Sauveteur d'acier |
| 538 | 4 × 50 m Relais Medley |
| 540 | 4 × 50 m Relais obstacles |
| 550 | 200 m Medley de sauvetage *(added for Canadien)* |
| 551 | 4 × 50 m Relais portage du mannequin *(added)* |
| 552 | 100 m Nage avec obstacles (Masters) *(added)* |

Cumulatifs are configured via the `CUMULATIFS` dict — one entry per
(age × gender), each listing the UIDs that contribute to the
cumulative score.

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
- Best times are stored in hundredths of a second
  (`ENTRYTIME` = total cs). Parser accepts `mm:ss.cc`, `hh:mm:ss.cc`,
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
