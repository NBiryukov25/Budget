# Weekly Cash Flow Survival Worksheet

`Weekly_Cash_Flow_Survival.xlsx` — a 13-week cash flow projection driven by recurring
transactions and weekly paychecks. Every number in it is a live Excel formula, so changing
the beginning balance or editing a recurring transaction re-drives the whole sheet.

## The five tabs

| Tab | What it does |
|---|---|
| **Instructions** | How to use the workbook, the colour legend, and every assumption made |
| **Recurring** | The editable recurring-transaction table — 30 slots, 17 pre-filled |
| **Cash Flow** | The main worksheet: 13 weekly blocks, each running from a starting balance to an ending balance |
| **Summary** | Week-by-week rollup, lowest balance, and the first week you run short |
| **Schedule** | The formula engine that expands recurring rules into dated occurrences — look, don't type |

## Using it

1. On **Cash Flow**, type your beginning balance in `E5` and a start date in `E6`. Weeks run
   **Friday through Thursday**, so use a Friday — the sheet warns you beside `E6` if you don't.
   `E7` is the cushion below which a week is flagged `TIGHT`.
2. On **Recurring**, edit any row or add one in the first empty row. Set Description, Type
   (Income/Expense), Amount as a positive number, Frequency, Next Due Date, and Active = Yes.
   It appears in the cash flow on its own.
3. Back on **Cash Flow**, record the real amount in the **Amount Paid** column as each due
   date arrives or passes. The cell turns yellow when something is due and still blank. Until
   you fill it in, the sheet projects the scheduled amount.
4. Each week block has a **Weekly Variable Expenses** section — type a description and amount
   for groceries, gas, and anything else. It is deducted from that week's income immediately.

## Supported frequencies

`Weekly` · `Every 2 Weeks` · `Monthly` · `Twice a Month` (15th and last day) · `One Time`

The **Weekend Rule** column shifts a due date that lands on a weekend: `Move Before` to the
Friday, `Move After` to the Monday, or `None`.

## Regenerating

```bash
pip install openpyxl
python3 build_cash_flow_workbook.py
```

`build_cash_flow_workbook.py` rebuilds the workbook from scratch. The recurring
transactions it seeds are in the `SEED` list near the top; the layout constants
(`N_WEEKS`, `SLOTS`, `VAR_ROWS`, `N_ITEMS`) control how big the sheet is.

Note that rebuilding overwrites anything typed into the workbook — edit the spreadsheet for
day-to-day use, and only rerun the script if you want to change the structure.

## Where the seeded data came from

The 17 pre-filled recurring transactions were transcribed from three screenshots of the
budget app's Recurring / Upcoming list for August 2026. Two of them are the weekly paychecks:
Bruno's Restaurant ($90.00, Wednesdays) and Compunnel Software ($790.00, Fridays).

Two assumptions worth checking:

- **Progressive Insurance ($143.37)** was cut off at the bottom of the third screenshot, so
  its due date is set to Aug 31, 2026.
- **Patrick Rombalski ($350.00)** is set to Twice a Month with Weekend Rule = Move Before,
  which reproduces the app's Fri Aug 14 date for the 15th-of-the-month payment.

The beginning balance ($1,250.00) and the Week 1 variable-expense rows are placeholders.
