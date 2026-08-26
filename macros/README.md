# BankMemoPivot.bas

A macro that rebuilds the **Summary** pivot table in the bank memo /
returned payment report and re-applies the report's design, so you can run it
each period and get an identical-looking sheet over new data.

Written against `bankmemo_distr_20260806_report.xlsx`.

## Installing it

The macro has to live inside the workbook, and a workbook with macros has to be
saved as `.xlsm`.

1. Open the report workbook.
2. **File > Save As** and pick **Excel Macro-Enabled Workbook (\*.xlsm)**.
   (A plain `.xlsx` silently discards macros when you save.)
3. Press **Alt+F11** to open the VBA editor.
4. **File > Import File…**, choose `BankMemoPivot.bas`, then **Alt+Q** to close
   the editor.
5. Save.

If you'd rather not import a file: Alt+F11, then **Insert > Module**, and paste
the contents of `BankMemoPivot.bas` into the blank module.

To put it on a button, go to **Developer > Insert > Button (Form Control)**,
draw it on the Summary sheet, and assign `RefreshSummaryPivot`.

If the Developer tab isn't showing: **File > Options > Customize Ribbon**, tick
**Developer**.

## Running it

**Alt+F8** lists three macros:

| Macro | When to use it |
| --- | --- |
| `RefreshSummaryPivot` | The normal monthly run, after new rows have been added to the **Data** sheet. Re-points the pivot at the current extent of Data, refreshes it, and re-applies the whole design. |
| `FormatSummaryPivot` | Design only, no refresh. Use when someone has hand-edited the sheet and you just want the look back. |
| `RebuildSummaryPivot` | Clears the Summary sheet and builds the pivot from scratch. Use if the pivot got deleted or its fields got scrambled. Asks for confirmation first. |

The normal workflow each period is: paste the period's rows onto the **Data**
sheet, then run `RefreshSummaryPivot`.

## What it produces

| | |
| --- | --- |
| Source | `Data!A1:N<last row>`, measured on every run, so the range grows with the data |
| Rows | Region › Customer › SAP No., tabular layout, labels repeated down, subtotals at the bottom of each group |
| Columns | Date grouped Years › Quarters › Months |
| Values | Count of SAP No., captioned "Bank Memo Count, Returned Payment Count" |
| Totals | Row grand total on, headed "Totals"; column grand total off |
| Style | PivotStyleLight13, no row or column stripes, last column emphasised |
| Type | Calibri 10 throughout, vertically centred. Values centred, Region and Customer left-aligned. The corner block, the top-right block and the SAP No. header are bold. |
| Shading | Value cells banded by quarter: odd quarters (Qtr1, Qtr3) light yellow, even quarters and the year/grand-total columns white. A quarter-total column takes its quarter's shade. |
| Chrome | Gridlines off, 90% zoom, frozen below the header rows, value columns 13.14 wide, header rows 20.1 / 20.1 / 13.5 tall |

### One deliberate difference from the 8-6-26 file

In that file the shading runs Qtr1 yellow, Qtr2 white, Qtr3 yellow — except for
**April**, a single month inside Qtr2 that is yellow while the rest of its
quarter is not. That looks like a leftover from hand-highlighting rather than
part of the design, so the macro shades whole quarters and leaves April white
along with the rest of Qtr2. Every other value column comes out exactly as it
is in the original.

If you'd rather have no shading at all, set `BAND_QUARTERS = False` near the top
of the module.

## Tweaking it

Everything adjustable is a `Private Const` in the first fifty lines:

- `DATA_SHEET`, `SUMMARY_SHEET`, `PIVOT_NAME` — rename these if the tabs are
  ever renamed.
- `FLD_SAP`, `FLD_REGION`, `FLD_CUSTOMER`, `FLD_DATE` — must match the header
  row on the Data sheet.
- `DATA_CAPTION`, `GRAND_TOTAL_CAP`, `COL_HEADER_CAP` — the captions.
- `PIVOT_STYLE`, `BODY_FONT`, `BODY_SIZE`, `BAND_COLOR`, `BAND_QUARTERS` — the
  look.
- `VALUE_COL_WIDTH`, `ROW1_HEIGHT`, `ROW2_HEIGHT`, `ROW3_HEIGHT`, `SHEET_ZOOM` —
  the measurements.

## Notes

- The macro sets `HasAutoFormat = False` on the pivot, so a plain
  **PivotTable Analyze > Refresh** no longer resizes the columns out from under
  the design.
- It always hands the pivot a freshly built cache. That drops the date grouping,
  which the macro then puts back from its own rules — the grouping is rebuilt
  identically every run rather than inherited from whatever state the workbook
  was in.
- Grouped date fields are found by prefix (`Years…`, `Quarters…`, `Months…`)
  because Excel names them `Months (Date)` in some versions and plain `Months`
  in others.
- Quarter shading is worked out by reading the quarter header text, which Excel
  writes as `Qtr1`, `Qtr1 Total`. On a non-English Excel, change
  `QUARTER_PREFIX`.
- If grouping fails, the usual cause is text that looks like a date sitting in
  the Data sheet's Date column. Select the column and use
  **Data > Text to Columns > Finish** to coerce it, then run the macro again.
