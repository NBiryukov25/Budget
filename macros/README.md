# BankMemoPivot.bas

A macro that rebuilds the **Summary** pivot table in the bank memo /
returned payment report and re-applies the report's design, so you can run it
each period and get an identical-looking sheet over new data.

Written against `bankmemo_distr_20260806_report.xlsx` and checked against
`bankmemo_distr_20260820_report.xlsm`.

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

**Replacing an older copy:** importing a second time creates a duplicate module
called `BankMemoPivot1` and the old code keeps running. Right-click the existing
`BankMemoPivot` module in the Project pane, choose **Remove BankMemoPivot**,
answer **No** when it offers to export, then import the new file. Or just open
the existing module, select all, and paste the new code over it.

To put it on a button, go to **Developer > Insert > Button (Form Control)**,
draw it on the Summary sheet, and assign `RefreshSummaryPivot`.

If the Developer tab isn't showing: **File > Options > Customize Ribbon**, tick
**Developer**.

## Running it

**Alt+F8** lists three macros:

| Macro | When to use it |
| --- | --- |
| `RefreshSummaryPivot` | The normal monthly run, after new rows have been added to the source sheet. Re-points the pivot at the current extent of that sheet, refreshes it, and re-applies the whole design. |
| `FormatSummaryPivot` | Design only, no refresh. Use when someone has hand-edited the sheet and you just want the look back. |
| `RebuildSummaryPivot` | Clears the Summary sheet and builds the pivot from scratch. Use if the pivot got deleted or its fields got scrambled. Asks for confirmation first. |
| `CheckSetup` | Changes nothing. Reports the module version, the tabs it can see, which pivot it found, which tab and columns it resolves each field to, and how many Region values are filled in. Run this first when something misbehaves. |

The normal workflow each period is: paste the period's rows onto the source
sheet, then run `RefreshSummaryPivot`.

## What it produces

| | |
| --- | --- |
| Source | The tab whose row 1 has SAP / Region / Customer / Date headings, measured on every run so the range grows with the data |
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

- `DATA_SHEET`, `SUMMARY_SHEET`, `PIVOT_NAME` — `DATA_SHEET` is only the first
  place the macro looks; see *Finding the source sheet* below.
- `FLD_SAP`, `FLD_REGION`, `FLD_CUSTOMER`, `FLD_DATE` — matched loosely against
  the header row, ignoring case, spaces and punctuation.
- `DATA_CAPTION`, `GRAND_TOTAL_CAP`, `COL_HEADER_CAP` — the captions.
- `PIVOT_STYLE`, `BODY_FONT`, `BODY_SIZE`, `BAND_COLOR`, `BAND_QUARTERS` — the
  look.
- `VALUE_COL_WIDTH`, `ROW1_HEIGHT`, `ROW2_HEIGHT`, `ROW3_HEIGHT`, `SHEET_ZOOM` —
  the measurements.

## Finding the source sheet and columns

The source tab is called `Data` in the 8-6-26 file and `Sheet1` in the 8-20-26
file, and the SAP column is headed `SAP No.` in one and `SAP No` in the other.
Rather than needing to be re-pointed each month, the macro:

- looks for a tab named `DATA_SHEET` first, and if that tab doesn't exist or
  doesn't have the right headers, searches every other tab (except Summary) for
  a row 1 containing SAP, Region, Customer and Date headings;
- matches column and field names with case, spaces and punctuation stripped, so
  `SAP No.` and `SAP No` are the same column.

It also nudges the pivot up to cell A1 if it was built lower down, as long as
the rows above it are empty.

## If the Region column is blank

The pivot groups Region › Customer › SAP No. If the Region column has no values,
every customer lands under a single `(blank)` heading and the report comes out
structurally unlike the 8-6-26 one however good the formatting is. The macro
says so at the end of a run when it sees this.

In the 8-6-26 file, Region was a `VLOOKUP` from the SAP number into a
SAP-number-to-region table on a hidden tab. A file without that table has
nothing to fill Region from, so it has to be supplied before the pivot can match.

## Check which version is loaded

Every message the module shows is stamped with its version — the title bar of
an error box reads `Bank Memo Pivot v4`. A message with **no version in the
title and no `Step:` line** is an older copy of the module still sitting in the
workbook.

That happens easily: each download of `BankMemoPivot.bas` lands in Downloads
alongside the last one as `BankMemoPivot (1).bas`, `BankMemoPivot (2).bas`, and
it's the original filename that gets picked in the import dialog. Check the
date on the file you're importing, and remove the old module first (see
*Replacing an older copy* above).

`CheckSetup` prints the version along with everything else it can see.

## If it reports an error

The error box names the step it got to and the Excel error number, for
example:

> **Bank Memo Pivot v4**
> Step: applying the pivot total options
> Error 438: Object doesn't support this property or method

That is enough to say which line to look at. Send the whole message on and it
can be fixed directly.

Separately, the cosmetic settings — layout, totals, captions, table style,
sorting, subtotals, moving the pivot to A1 — are now each applied on their own
through `OptSet` / `OptCall`. Excel builds differ in which of these they expose,
and error 438 means one of them isn't available in a given build. Rather than
losing the whole run to one property, anything that won't take is collected and
listed in a single message at the end:

> The pivot was rebuilt and formatted, but this copy of Excel would not accept
> these settings: ColumnHeaderCaption — Object doesn't support this property
> or method

The pivot is still built and formatted; only the named setting is missing.

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
  the source sheet's Date column. Select the column and use
  **Data > Text to Columns > Finish** to coerce it, then run the macro again.
