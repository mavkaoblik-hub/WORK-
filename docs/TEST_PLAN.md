# Transport Center 0.2 — Windows Excel acceptance test

**Environment required:** Windows 11, licensed Microsoft Excel LTSC desktop. These steps are **not yet executed**; do not treat source CI as a VBA runtime test.

## One-time setup
1. Download repository ZIP from GitHub; unzip into a writable, trusted local directory.
2. Put `Транспортний_центр_v8_робочий_Times_New_Roman.xlsx` beside `START_INSTALL.cmd`. Use the **empty** working book, not DEMO with invented trips.
3. Close every Excel instance; run `START_INSTALL.cmd`. If blocked, confirm Excel's **Trust access to the VBA project object model** is enabled, then rerun. Disable the setting again when the build is complete.
4. Expect `Транспортний_центр.xlsm`, with the original XLSX unchanged. If Excel reports "repair content", **stop** and preserve the repair log.
5. Open the new XLSM only after reviewing its origin and macro permissions.

## Acceptance cases
| No. | Action | Required result |
| --- | --- | --- |
| T01 | Open XLSM | No content repair prompt; first worksheet is **Зміст** |
| T02 | Press **Заповнити дані** | Administrator form opens, registered driver dropdown works, current date prefilled in B6/B9/B10 |
| T03 | Add driver with button, then open **Зміст** | New worksheet exists between driver boundary tabs; PІБ / vehicle appear on driver sheet **and** registry; no formulas overwritten in registry B:N |
| T04 | Enter one trial trip on form, leave payment zero | Unique ID written into one row on driver and same ID in `Нарахування`; odometer distance formula and summaries calculate; backup created |
| T05 | Enter a spending row in driver worksheet using the exact ID | Expenses flow into administrator totals; parking tariff uses effective-date history (or manual sum when unknown) |
| T06 | **Закрити рейс** | After correct/complete data, L:N in `Нарахування` contain status, timestamp and username; duplicate closure refused |
| T07 | **PDF-звіт** | `_Звіти` created; nonempty PDF generated from the report sheet |
| T08 | Disconnect internet, reopen XLSM and press F9 / Ctrl+Alt+F9 | Existing course unchanged. No WEBSERVICE / FILTERXML formulas or background website update |
| T09 | Reconnect and press **Оновити з інтернету** | Confirm request, only NBP exchange rate and matching archive date updated; parking tariffs remain unchanged |
| T10 | Close Excel, run `RESTORE_BACKUP.cmd` | Choice of snapshot; current XLSM saved as pre-restore backup; selected backup restored |
| T11 | Attempt restore with XLSM open | Restoration blocked; no overwrite of active workbook |

## Performance
Time the form-opening and save operations on a workbook with four drivers, 24 test trips and ~100 spending lines. Target: interactive behavior without noticeable freezing. No timings have been measured so far.

## Known limits
- The parking-price downloader is not connected to the XLSM button in version 0.2.
- No live VBA/COM test has run in the current authoring environment.
- The repository is **public**. Never submit real driver identities, customer information, routes or business expenses to GitHub.
- Macros require Excel macro security approval. Never enable untrusted macros.
