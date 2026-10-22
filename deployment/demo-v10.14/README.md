## v10.14.3: Excel LTSC Run-time error 91 diagnostic

During the user’s v10.14.2 installation, `TransportStatusFX.TC_RefreshBoth` failed with "Object variable or With block variable not set" at its error-rethrow handler. That highlighted line did **not** identify the original failing statement.

The new `TransportStatusFX.bas` records **26 refresh stages**, the original `Err.Number`, `Err.Source`, `Erl` and message, and writes them to both operational sheets' `N1` cells on failure. Chart cleanup is now presentation-only (its errors cannot block ledger rendering).

`INSTALL_DIAGNOSTIC.ps1` **must** be packaged as `INSTALL_OFFLINE.ps1`: the installed XLSM is saved for debugging, but a failed VBA smoke test now produces `VBA_SMOKE_TEST_FAILED.txt` and a nonzero process exit rather than a misleading SUCCESS report. The previous v10.12 installer had swallowed this failure.

This is a **diagnostic release**, not a claim the runtime VBA error has been eliminated. Once the real failing stage is reported, make a narrow root-cause fix. GitHub CI still cannot execute desktop Microsoft Excel VBA.

To locally assemble a FULL version, include the verified synthetic XLSX through `--base-workbook ... --require-base`. A CI source-only ZIP does not include this XLSX.

---
# Транспортний центр DEMO v10.14.1 — склад повної поставки

**Виправлення 22.10.2026:** перший архів v10.14 не мав фізичного `.xlsx`. Хоча запуск `CREATE_DEMO_BASE.ps1` міг її створити, це не відповідає вимозі повного ZIP. У v10.14.1 тестова база `Транспортний_центр_v8_робочий_Times_New_Roman.xlsx` включена всередину ZIP.

## Повний архів для Windows

`Transport_Center_DEMO_v10.14.1_FULL.zip` має містити початкову XLSX, `START_INSTALL.cmd`, `INSTALL_OFFLINE.ps1`, 4 VBA-модулі, код `ThisWorkbook`, маршрути і синтетичний `demo_seed.json`. Усі дані вигадані.

Поклади ZIP у нову папку, розпакуй і запусти `START_INSTALL.cmd` на Windows з Microsoft Excel. Інсталятор не перезаписує попередні `.xlsm`. Якщо потрібне створення VBA-проєкту, тимчасово дозволь Trust access to VBA project object model, а після інсталяції вимкни. На виході очікується `Транспортний_центр.xlsm`.

**У GitHub Actions автоматично публікується лише SOURCE-ONLY ZIP** — він не містить базової XLSX і не повинен називатися повним інсталятором. Повний ZIP був зібраний окремо для тестування та наданий у бесіді.

## Як відтворити FULL ZIP у локальному репозиторії

```powershell
.\deployment\demo-v10.14\BUILD_PACKAGE.ps1
```

Сценарій спочатку перевіряє/створює синтетичну базову XLSX локально в Excel, потім викликає Python 3 для пакування, завжди передаючи `--base-workbook --require-base`. `ASSEMBLE_PACKAGE.py` відмовиться позначати повним ZIP без потрібного файлу. Для окремої наявної початкової тестової книги:

```powershell
.\deployment\demo-v10.14\BUILD_PACKAGE.ps1 -BaseWorkbook "C:\\TC_TEST\\base.xlsx"
```

**Попередження:** XLSX, що генерується для DEMO, — мінімальна синтетична основа, **не оригінальна робоча версія v8**. Використовувати її для перенесення реальних даних не можна.

## Контроль

- 13 початкових аркушів, зокрема `ШАБЛОН_ВОДІЯ`, `КІНЕЦЬ_ВОДІЇВ`, `Зміст`, `Нарахування`.
- 7 вигаданих водіїв, 9 автомобілів, 75 рейсів (70 завершено, 3 активні, 2 заплановано), 259 записів витрат.
- 114 маршрутів, контрольний дохід завершених рейсів 149 987,91 USD, сума зареєстрованих витрат 48 211,73 USD.
- PDF v10.13, зведення v10.14, події стрічки Excel v10.12.

У CI перевіряються джерела VBA, структури архіву та PowerShell. **Виконання VBA, інсталювання XLSM, формування PDF й поведінка стрічки Excel LTSC на реальному Windows не перевірені** — це предмет тестування на комп'ютері користувача.
