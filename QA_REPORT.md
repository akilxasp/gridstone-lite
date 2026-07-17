# Gridstone QA Report

Date: 2026-07-16

Build tested: 0.1.0, commit `7333abc`

Mode: diagnostic only; application defects were not fixed in this round.

## Test summary

| Suite | Passed | Failed | Total |
|---|---:|---:|---:|
| Existing workbook unit tests | 4 | 0 | 4 |
| Spreadsheet interaction tests | 5 | 1 | 6 |
| Workbook/file compatibility matrix | 22 | 8 | 30 |
| UI implementation audit | 6 | 13 | 19 |

Additional checks:

- TypeScript and production renderer build: PASS
- Production dependency audit: PASS, zero production vulnerabilities
- macOS `Info.plist` and file associations: PASS
- Installed application launch: PASS
- Strict macOS bundle signature verification: FAIL

## Confirmed product defects

### High priority

#### QA-01 — Direct typing drops the first character

1. Select a cell.
2. Type `42` without double-clicking.
3. Press Enter.

Expected: the cell contains `42`.

Actual: the cell contains `2`.

The inline editor selects its seeded first character after focus. The second character replaces that selection.

#### QA-02 — Save As can produce a file with mismatched content and extension

The renderer serializes the workbook before the native Save As dialog returns the selected format. Starting with XLSX and choosing the CSV filter can therefore write XLSX bytes to a `.csv` filename.

#### QA-03 — Non-ASCII CSV text is corrupted

Round-trip value:

- Expected: `comma, quote " and ✓`
- Actual: `comma, quote " and â`

#### QA-04 — Imported data beyond row 500 or column 100 cannot be reached in the UI

The file engine preserves cells such as `A501` and `CW1`, but the renderer caps the visible grid at 500 rows and 100 columns. Saving is possible, but users cannot inspect or edit those preserved cells.

#### QA-05 — XLSX formatting is lost during round-trip

Bold and fill-color formatting did not survive export followed by import. Values, formulas, Unicode text, and multiple sheets did survive the same XLSX round-trip.

#### QA-06 — Cross-sheet ranges calculate incorrectly

`=SUM(Details!A1:A2)` returned `0` instead of `12`.

### Medium priority

#### QA-07 — AVERAGE treats blank cells as zero

`=AVERAGE(A1:A2)` with `A1=10` and `A2` blank returned `5` instead of `10`.

#### QA-08 — MIN treats blank cells as zero

`=MIN(A1:A2)` with `A1=10` and `A2` blank returned `0` instead of `10`.

#### QA-09 — Invalid formula characters are silently ignored

`=1@2` returned `1` rather than a formula error.

#### QA-10 — Date number format is selectable but not implemented

Applying Date format produces the same display as General format.

#### QA-11 — Font family and font size controls do nothing

Both selectors are visible but have no change handlers or corresponding cell-style fields.

#### QA-12 — Add Chart does not add a chart

The dialog previews a chart, but Add Chart only closes the dialog and changes the status message. No chart is stored in the workbook.

#### QA-13 — Freeze Top Row does not affect rendering

The command changes `frozenRows` in the worksheet model, but the grid never reads that state.

#### QA-14 — New Workbook is not blank

New Workbook reloads the demonstration quarterly-sales workbook.

#### QA-15 — Zoom changes text only

Row height and column width do not scale with the displayed zoom percentage.

#### QA-16 — Notes consume one undo entry per character

Every textarea change writes the entire workbook immediately. Long notes can fill the 100-entry undo history and trigger repeated grid renders.

#### QA-17 — macOS bundle signature is invalid

Strict verification reports: `code has no resources but signature indicates they must be present`. The locally installed app launches because quarantine was removed, but the bundle is unsuitable for normal external distribution without proper signing and notarization.

#### QA-18 — Windows file opening lacks single-instance forwarding

The app does not call `requestSingleInstanceLock`. Opening a workbook while Gridstone is already running may create a second Windows process instead of forwarding the file to the existing window.

### Known placeholder

#### QA-19 — Share is not implemented

Share only displays a message stating that cloud collaboration can be connected later.

## Verified working behavior

- Double-click cell editing and single commit
- Formula-bar editing and recalculation
- Bold formatting and Undo
- Adding and switching worksheets
- Find shortcut and match count
- Numeric, boolean, and text parsing
- Arithmetic, SUM, IF, comparisons, division-by-zero, and circular-reference detection
- XLSX values, formulas, Unicode text, and multiple sheets
- CSV numeric values and computed formulas
- TSV text and numeric import
- Data preservation by the file engine beyond the current UI grid limit
- Production build and production dependency audit
- macOS application metadata and XLSX/XLS/CSV/TSV associations

## Test limitation

The installed Mac application could not be controlled through macOS Accessibility because permission is not enabled for Codex, and the in-app browser test connection was unavailable. UI interactions were therefore exercised through the React DOM interaction suite. A final release candidate should still receive a manual Mac smoke test and a native Windows test after the defects above are fixed.
