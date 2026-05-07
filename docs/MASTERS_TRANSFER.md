# Masters Transfer Process

## Problem

SPLASH Meet Manager doesn't natively support a workflow where Masters athletes swim in the same prelim heats as Open/15-18 athletes and are later moved to a separate Masters final for scoring. Additionally, if Masters 5-year age brackets exist on the prelim event, SPLASH reassigns Open athletes (aged 25+) into those brackets, breaking points/medals.

## Solution

A two-phase approach using `BONUSENTRY` as a marker:

1. **Phase 1 (Import)**: All athletes (including Masters) are registered in the prelim event under the `[19-99]` bracket. Masters athletes are marked with `BONUSENTRY='T'` so they can be identified later.

2. **Phase 2 (Transfer)**: After prelim heats are run, a VBS script moves Masters athletes from the prelim to the Masters final event, creating heats and deleting the prelim rows.

## How Masters Are Marked

### MDB Loader Path (`load_to_mdb.py`)
- Sets `SWIMRESULT.BONUSENTRY = 'T'` directly for any athlete whose ticket is `MA` (Masters).
- No additional steps needed.

### Lenex Path (`load_to_lenex.py`)
- Suffixes the athlete's `LICENSE` field with `_MA` in the Lenex XML (e.g., `YETP42` → `YETP42_MA`).
- When SPLASH imports the Lenex, it preserves the LICENSE value.
- The VBS script detects the `_MA` suffix, sets `BONUSENTRY='T'`, and strips the suffix back to the clean NRAN.

### Masters-Only Events (e.g., UID 541 — 100m Obstacle Masters)
- These events have no prelim — only a timed final.
- Athletes are routed directly to the Masters final with their correct 5-year age bracket.
- They still get `BONUSENTRY='T'` but the VBS won't move them (they're already in the final).

## Template Requirement

The template MDB must NOT have Masters 5-year brackets (AGEMIN 25-99) on prelim events. Only `[15-18]` and `[19-99]` should exist on prelims. This prevents SPLASH from reassigning athletes to age-specific brackets during heat generation.

Use `cleanup_prelim_brackets.vbs` (one-time) to remove them:
```
cscript cleanup_prelim_brackets.vbs "C:\path\to\template.mdb"
```

## Workflow

### MDB Path
```
1. Upload xlsx → webapp generates meet.mdb
2. Open meet.mdb in SPLASH
3. Generate heats (all athletes in prelim together)
4. Run prelim races (or simulate_results.bat for testing)
5. Run masters_transfer.bat
   → Transfers BONUSENTRY='T' athletes to Masters finals
   → Creates heats in final events
   → Deletes prelim rows
6. Continue with Masters finals in SPLASH
```

### Lenex Path
```
1. Upload xlsx → webapp generates meet.lxf + meet.mdb
2. Import meet.lxf into SPLASH (creates entries)
3. Run masters_transfer.bat
   → Detects _MA suffix in LICENSE
   → Sets BONUSENTRY='T' and strips _MA
4. Generate heats in SPLASH (Masters stay in prelim with everyone)
5. Run prelim races (or simulate_results.bat for testing)
6. Run masters_transfer.bat AGAIN
   → Now transfers athletes with BONUSENTRY='T' and SWIMTIME > 0
   → Creates heats in final events
   → Deletes prelim rows
7. Continue with Masters finals in SPLASH
```

Note: For the Lenex path, `masters_transfer.bat` is run **twice**:
- First time: marks Masters (converts `_MA` → `BONUSENTRY`)
- Second time: does the actual transfer (after prelims are run)

### Testing with Simulated Results
```
1. Generate MDB or import Lenex
2. Generate heats in SPLASH
3. Run simulate_results.bat (writes random SWIMTIME ±5% of entry time)
4. Run masters_transfer.bat (transfers Masters to finals)
```

## VBS Scripts Included in Output ZIP

| Script | Purpose |
|--------|---------|
| `masters_transfer.vbs` | Mark Masters + transfer to finals |
| `masters_transfer.bat` | Runs masters_transfer.vbs on meet.mdb |
| `simulate_results.vbs` | Generate random swim times for testing |
| `simulate_results.bat` | Runs simulate_results.vbs on meet.mdb |

## Technical Details

### BONUSENTRY Field
- `SWIMRESULT.BONUSENTRY` is a text field ('T'/'F') in the SPLASH MDB schema.
- SPLASH does not use it for lifesaving meets — safe to repurpose as a Masters marker.
- The VBS transfer script queries: `WHERE BONUSENTRY='T' AND SWIMTIME > 0`

### _MA Suffix
- Applied to `ATHLETE.LICENSE` in the Lenex XML.
- SPLASH preserves LICENSE on import without modification.
- The VBS uses `WHERE LICENSE LIKE '%[_]MA'` to find them.
- After marking, the suffix is stripped: `UPDATE ATHLETE SET LICENSE='...' WHERE ...`

### Age-Based Fallback
- If neither `BONUSENTRY='T'` nor `_MA` suffix is found, the VBS falls back to transferring all athletes aged 25+ (original behaviour).
- This handles legacy MDBs created before the BONUSENTRY system.

### UCanAccess Limitation
- UCanAccess (Java) cannot persist DELETE operations to MDB files created by SPLASH's Lenex import.
- The VBS scripts use Microsoft ACE OLEDB (native Windows driver) which works on all MDB files.
- This is why the transfer is done via VBS on Windows, not in the Python container.
