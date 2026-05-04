# Running ebimport_splash

This tarball contains everything needed to import a registration
workbook into a SPLASH Meet Manager 11 database, or to emit a Lenex
3.0 meet file.

## Prerequisites on the target machine

- **Python 3.10+** on `PATH` (`python3 --version`)
- **Java 8+** on `PATH` (`java -version`) — only for the MDB loader;
  the Lenex exporter doesn't need Java.

Nothing else. All Python packages (`openpyxl`, `jaydebeapi`, `JPype1`)
and the UCanAccess jars are bundled inside `vendor/`. **No network
access** is required at runtime.

## Install

Extract the tarball anywhere:

```bash
tar xzf ebimport_splash-*.tgz
cd ebimport_splash-*
```

That's it. The first invocation of either script will create a local
`.venv/` in this directory and install the bundled wheels there.

## Run the MDB loader

```bash
./scripts/run_mdb.sh \
    --xlsx CPLC2026FINAL.xlsx \
    --mdb  Canadien.mdb
```

Useful flags:

| Flag | What it does |
|---|---|
| `--dry-run`          | Parse + validate, no DB writes. |
| `--wipe`             | Delete existing clubs/athletes/events/entries before loading. |
| `--issues-full`      | List every issue (no per-category cap). |
| `--issues-out PATH`  | Write the issues section to a text file. Implies `--issues-full`. |

See the top of `load_to_mdb.py` for the full docstring.

## Run the Lenex exporter

```bash
./scripts/run_lenex.sh \
    --xlsx CPLC2026FINAL.xlsx \
    --out  candien.lef

# or zipped .lxf (what SPLASH prefers)
./scripts/run_lenex.sh \
    --xlsx CPLC2026FINAL.xlsx \
    --out  candien.lxf --zip
```

## Output

Both scripts print a **Summary** section and then an **Issues**
section that flags data-quality problems in the xlsx (unknown
tickets, bad times/DOBs, age-bracket mismatches, fuzzy-duplicate
club/athlete names, incomplete relays, ...). Share the issues
output with the xlsx owner for cleanup.

## Re-runs (additive mode)

The MDB loader is safe to re-run on the same `.mdb` after you've
started organising the meet in SPLASH. It detects existing rows via
stable keys (normalised club name, athlete name+license, etc.) and
only inserts what's missing. It never touches HEATID / LANE /
SWIMTIME / session assignments / custom FEE / ROUNDNAME. ENTRYTIME
is updated only when a faster time is supplied (never regresses).

## Troubleshooting

**Q: "Python 3 not found"**
Install Python 3.10 or newer, or point at a specific interpreter:
```bash
PYTHON=/usr/local/bin/python3.12 ./scripts/run_mdb.sh ...
```

**Q: "Java not found"**
Install OpenJDK 8 or newer and make sure `java` is on `PATH`.
On Debian/Ubuntu: `sudo apt install default-jre`.

**Q: First run fails saying a wheel is incompatible with my Python**
The bundled wheels are manylinux x86_64 for Python 3.10+. If your
Python is on a different platform (e.g. macOS arm64), rebuild the
tarball from source using `make dist` on your platform.
