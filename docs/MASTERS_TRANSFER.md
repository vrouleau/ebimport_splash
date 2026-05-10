# Masters Transfer Process

## Problem

SPLASH Meet Manager doesn't natively support a workflow where Masters athletes swim in the same prelim heats as Open/15-18 athletes and are later moved to a separate Masters final for scoring.

## Solution

A two-phase approach:

1. **Phase 1 (Import)**: All athletes (including Masters) are registered in the prelim event under the `[19-99]` bracket. Masters athletes are marked with `HANDICAP exception='X'` in the Lenex, which SPLASH imports as `EXCEPTIONNAGEUR='X'` on the athlete record.

2. **Phase 2 (Transfer)**: After prelim heats are run, a VBS script identifies Masters athletes by their exception code, moves them from the prelim to the Masters final event, creating heats and deleting the prelim rows.

## How Masters Are Identified

- The Lenex generator adds `<HANDICAP exception="X"/>` to any athlete with Masters entries.
- SPLASH imports this as `EXCEPTIONNAGEUR='X'` on the ATHLETE record.
- The VBS script queries `WHERE EXCEPTIONNAGEUR='X'`, marks their SWIMRESULT rows with `BONUSENTRY='T'`, then transfers them.
- The 'X' code also appears on heat sheet PDFs, making Masters athletes visually identifiable.

### Masters-Only Events (e.g., UID 541 — 100m Obstacle Masters)
- These events have no prelim — only a timed final.
- Athletes are routed directly to the Masters final with their correct 5-year age bracket.
- The VBS won't move them (they're already in the final).

## Workflow

```
1. Upload xlsx + meet .lxf → webapp generates entries .lxf
2. Import entries .lxf into SPLASH
3. Generate heats in SPLASH (Masters swim with everyone in prelim)
4. Run prelim races (or simulate_results.bat for testing)
5. Run masters_transfer.bat
   → Detects exception='X' athletes
   → Marks BONUSENTRY='T' on their SWIMRESULT rows
   → Transfers those with SWIMTIME > 0 to Masters finals
   → Creates heats in final events
   → Deletes prelim rows
6. Continue with Masters finals in SPLASH
```

### Testing with Simulated Results
```
1. Import Lenex into SPLASH
2. Generate heats
3. Run simulate_results.bat (random SWIMTIME ±5% of entry time)
4. Run masters_transfer.bat (transfers Masters to finals)
```

## Template Requirement

The template MDB must NOT have Masters 5-year brackets (AGEMIN 25-99) on prelim events. Only `[15-18]` and `[19-99]` should exist on prelims.

## Prerequisites (Windows)

The VBS scripts require the **Microsoft Access Database Engine** (ACE OLEDB provider):
- [Access Database Engine 2016 Redistributable](https://www.microsoft.com/en-us/download/details.aspx?id=54920)

## VBS Scripts

| Script | Purpose |
|--------|---------|
| `masters_transfer.vbs` | Identify Masters (exception=X) + transfer to finals |
| `masters_transfer.bat` | Runs masters_transfer.vbs on meet.mdb |
| `simulate_results.vbs` | Generate random swim times for testing |
| `simulate_results.bat` | Runs simulate_results.vbs on meet.mdb |

## Technical Details

### Exception Code
- `ATHLETE.EXCEPTIONNAGEUR = 'X'` in the SPLASH MDB schema.
- Set via `<HANDICAP exception="X"/>` in Lenex import.
- Visible on heat sheet PDFs as 'X' next to the athlete name.

### BONUSENTRY Field
- `SWIMRESULT.BONUSENTRY` ('T'/'F') is used internally by the VBS as a transfer marker.
- The VBS sets it from the exception code, then queries `WHERE BONUSENTRY='T' AND SWIMTIME > 0` for transfer.

### Age-Based Fallback
- If no exception-marked athletes are found, the VBS falls back to transferring all athletes aged 25+.
