# v10.15.2 — refinement of the clean light Excel interface

Based on Excel LTSC screenshots (22 October 2026), the v10.15.1 light theme works but needs easier navigation and clearer empty data states.

## Implemented

- **Dashboard (Зміст):** 3 *internal-only* Excel hyperlinks in the subtitle row lead to existing sections at B15 (Рейси), B20 (Довідники), and B25 (Фінанси). No new VBA handlers, external destinations, network requests, or changes to the 16 working actions.
- **Поточний стан:** explanatory text on the active-trip forecast row is merged across E:L instead of being clipped in one column. If there are no drivers or trips, a clean-start instruction appears above the forecast.
- **Архів рейсів:** clean-first-use guidance appears only when all statuses are empty; existing completion-month analytics and source ledger remain unchanged.
- **Курси валют:** label/instructions use a left-table/right-help layout. Manual input positions B7:B9, existing FX conversion code, source and dates are unchanged. Actions moved visually to row 14 and return home to row 18.
- Original empty 11-sheet initial XLSX, 114 reference routes, the no-overwrite Windows installer, event caching and other VBA modules are retained.

## Build/package boundaries

A complete **FULL** ZIP was distributed separately in the chat, with the original v10.15.1 clean XLSX copied **byte-for-byte**. Only the presentation module, PowerShell installer layout, status/archive VBA display, version messages and package instructions were modified.

Validation performed on the supplied ZIP: CRC, all 13 manifest files and their SHA-256 checksums, parseability of XML/RELS in the XLSX, 11 expected sheets, ASCII launcher, Unicode PowerShell BOM, 16 dashboard buttons, no `DEMO_SEED` calls, and balanced VBA procedure delimiters.

GitHub Actions checks the updated source contracts and Windows PowerShell syntax. **The full XLSM is created on Windows**: there is no claim that the Excel LTSC VBA runtime, HYPERLINK rendering, or button activation was exercised in CI. Do not install over an existing workbook.

## Installation/test

1. Extract the full ZIP to a **new folder** (not the existing v10.15.1 folder).
2. Close Excel, run `START_INSTALL.cmd` and temporarily authorize trusted VBA project access only if the installer requires it. Revoke that permission after installation.
3. On `Зміст`, click all 3 top shortcuts; verify they jump to the intended section and all 16 existing buttons still work.
4. Open `Поточний стан`: look at the empty-state message and the untruncated forecast explanation.
5. Open `Архів рейсів`, verify empty-state text and preserved header/navigation.
6. Open `Курси валют` and test **manual** USD/EUR/UAH entry using the original cells B7:B9.
7. Check that the ribbon returns to its former state on Excel workbook deactivation/close.

The established clean v10.14.8 build is unchanged.
