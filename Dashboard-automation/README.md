# Validation Dashboard Automation (Node.js)

Generates the **validation sheet** for a given **site + date** and fills the
**Expected** values (column G) from the **DGR workbook**, using the
"Comments" column (E) of your validation template to know *which DGR field to read*.

## Run
```bash
npm install
node index.js \
  --dgr      "TG1_1_JULY_TOTAL_CALCULATION_DGR.xlsx" \
  --template "Module1_validation.xlsx" \
  --site     TG1.1 \
  --date     2026-07-16            # repeat / comma-separate, or use --all-dates
```
Output: `output/TG1_1_validation.xlsx` -> one sheet per date (`TG1.1_2026-07-16`) + a `Run_Log` sheet.

## Executive Dashboard
Use the **Executive Dashboard validation** option in the web UI. The validation reference `inputs/Executive-Dashboard-reference.xlsx` is selected automatically. Upload the DGR workbook and separate KPI comparison workbook. The reference sheet's `Comments` column provides cell addresses (for example, `D14`) for KPI fields and labelled cell addresses for Chart fields. The KPI comparison workbook is read from the matching report-date worksheet when available, or from its first worksheet if it has no date-named sheet.

KPI/chart rows are matched by Sub Module + Field Name in the Executive reference sheet, then the mapped cells are read from the KPI comparison workbook and written to column G. Other rows continue to use source markers in column E (for example, `Exec_DGR_AM`) to read the matching DGR sheet/column. Markers `BGA_*` and `DTA_*` resolve to the `BRIDGE_GAP_ANALYSIS` and `DIGITAL_TWIN_ANALYSIS` sheets, respectively; both use column A for the date and match the selected date. Date matching supports Excel date values and common date strings such as `10-08-2026`.

## Browser dashboard
Start the local web app with:
```bash
npm start
```
Then open `http://localhost:3000`. Choose **Dashboard validation** or **Executive Dashboard validation** in the Dashboard type selector. The validation reference is automatically selected from `inputs/` (`Dashboard_validation_sheet.xlsx` for Dashboard, `Executive-Dashboard-reference.xlsx` for Executive); no template upload is needed. Upload the DGR, choose a configured site, and select one or more dates (or choose **All dates**). Executive mode additionally asks for a KPI comparison workbook; its expected values are read using the KPI/chart cell addresses in the automatically selected Executive reference sheet's Comments column. Source markers remain the fallback for other rows. The dashboard shows filled, skipped, and unresolved row counts after generation. Download the workbook from the latest-run panel or the generated reports list; each workbook includes a `Run_Log` sheet.

The **Generated reports** list includes a Delete action for each workbook. Deletion requires confirmation and permanently removes that generated report from `report/`.

### Record dashboard interactions
`npm start` launches only the local validation website. To open Playwright Inspector and record dashboard interactions, run this in a second terminal:
```bash
npm run record:dashboard
```
Sign in and perform the station/date and dashboard interactions to record. Close the Inspector to save the generated Playwright Test to `tests/dashboard.spec.js` and refresh `auth.json`. This is a separate interactive recording step; **Run validation** uses the saved Playwright reader and does not open the recorder.

Uploaded workbooks are held in a temporary directory and removed after inspection or generation. Generated workbooks are kept in `report/`. Set the `PORT` environment variable to use a different local port. The command-line workflow remains available with `node index.js` (or `npm run cli`).

## How each row is resolved
| Validation row | Where the value comes from |
|---|---|
| `Exec_DGR_AM`, `Exec_DGR_AP (Inv Gen)` ... | `EXECUTIVE_DGR` sheet, that column letter, row matching the date |
| `WEATHER_AV`, `Weather_N` ... | `Weather` sheet, that column, row matching the date |
| `AC_Energy_Last active value` / Active Power | `AC_Energy!BA` (`AC_Power`) last value > 0 on the requested date, with time `(h.mm)` |
| GTI `Last value > 5` | `Expected Power!D` last GTI > 5 of the day + time |
| Expected Power `last > 5 value` | `Expected Power!M` at that same timestamp |
| Peak GTI | max GTI of the day (`Expected Power!D`) |
| Peak Power (no comment) | `Weather!AX` (set in `config/rules.json`) |
| Inverter Reading (INVn) rows | `Inverter_Wise` sheet: INV_GEN / INV_PR / INV_PLF / INV_YIELD |
| Inverter with INV_GEN = 0 | merged cell `Inverter Inactive` |
| PA / GA / Status | not in DGR -> left blank (marked SKIPPED in Run_Log) |

Only the text of the comment is read (anything after the column letter, e.g. "(Net Gen)", is ignored).
Rows are matched by **Sub Module + Field Name**, never by row number.

## Expected value precision
Expected values preserve the source precision from the selected DGR or KPI comparison cell. Numeric cells formatted as percentages are displayed as percentages in the report; other numeric values are not truncated or rounded. Filled rows keep the `Expected:` label.

`config/rules.json` determines which source to read and which rows to include; decimal-place formatting settings are not applied.
To add a new site, add it under `sites` (`plantName` = value of column A in `EXECUTIVE_DGR`).

## Output sheet
Same layout/styling as the template. Columns D and F (MV_DB / Application Pass-Fail)
are cleared for the tester and get a Pass/Fail dropdown with green/red colouring.
`Run_Log` lists every row: comment, resolved source, raw value, text written, status, note.
