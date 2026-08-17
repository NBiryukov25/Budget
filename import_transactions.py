#!/usr/bin/env python3
"""Append a bank CSV export into the survival workbook.

Written for Revolut/Monarch-style exports with the columns
Date, Merchant, Category, Account, Amount. Only input cells are written; every
formula is left alone.

What it does, in order:

  1. Drops anything dated before the workbook's start date - that money is
     already inside the beginning balance.
  2. Drops rows on a second account ("Cash on Hand" by default). A single-balance
     sheet counts the ATM withdrawal when it leaves the bank; counting the cash
     spending as well would double it.
  3. Matches a transaction to an already-scheduled recurring bill when the amount
     is identical, the direction agrees, and the dates are within a few days.
     Those become that row's Amount Paid / Received rather than a duplicate line.
  4. Groups whatever is left by merchant and day, then writes it into that week's
     free variable-expense rows.
  5. Skips anything already in the sheet, so re-running over an overlapping export
     does not double-post.

    python3 import_transactions.py <book.xlsx> <export.csv> [--out FILE]
                                   [--cash-account "Cash on Hand"] [--dry-run]
"""

import argparse
import csv
import datetime as dt
from collections import defaultdict

from openpyxl import load_workbook

N_WEEKS, SLOTS, FIRST_BLOCK = 13, 10, 13
DATE_F = "ddd mm/dd"
MATCH_DAYS = 4          # how far a posting may drift from its scheduled date
CENTS = 0.005


def block(w, var_rows):
    h = FIRST_BLOCK + (SLOTS + var_rows + 5) * (w - 1)
    return {"head": h, "first_slot": h + 2, "last_slot": h + 1 + SLOTS,
            "first_var": h + 3 + SLOTS, "last_var": h + 2 + SLOTS + var_rows}


def detect_var_rows(cf):
    """Work out the block height from where week 2's banner sits."""
    for vr in range(4, 41):
        h2 = FIRST_BLOCK + (SLOTS + vr + 5)
        v = cf[f"B{h2}"].value
        if isinstance(v, str) and v.startswith("="):
            return vr
    raise SystemExit("could not work out the block height - is this the right workbook?")


def read_csv(path):
    out = []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for row in csv.DictReader(fh):
            raw = (row.get("Amount") or "").strip().replace("$", "").replace(",", "")
            if not raw or not (row.get("Date") or "").strip():
                continue
            m, d, y = [int(x) for x in row["Date"].split("/")]
            out.append({"date": dt.date(y, m, d),
                        "merchant": (row.get("Merchant") or "?").strip(),
                        "category": (row.get("Category") or "").strip(),
                        "account": (row.get("Account") or "").strip(),
                        "amount": float(raw)})
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("book"); ap.add_argument("csv_path")
    ap.add_argument("--out"); ap.add_argument("--cash-account", default="Cash on Hand")
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    wb = load_workbook(a.book)
    vals = load_workbook(a.book, data_only=True)
    cf, cfv = wb["Cash Flow"], vals["Cash Flow"]
    var_rows = detect_var_rows(cf)

    start = cfv["E6"].value
    start = start.date() if isinstance(start, dt.datetime) else start
    end = start + dt.timedelta(days=N_WEEKS * 7 - 1)

    txns = read_csv(a.csv_path)
    kept, skipped = [], defaultdict(list)
    for t in txns:
        if t["date"] < start:            skipped["before start"].append(t)
        elif t["date"] > end:            skipped["past horizon"].append(t)
        elif t["account"] == a.cash_account: skipped["other account"].append(t)
        else:                            kept.append(t)

    # what the sheet already holds, so a re-run cannot double-post
    existing = set()
    for w in range(1, N_WEEKS + 1):
        b = block(w, var_rows)
        for r in range(b["first_var"], b["last_var"] + 1):
            nm = cf[f"B{r}"].value
            if nm and not str(nm).startswith("="):
                d = cfv[f"C{r}"].value
                d = d.date() if isinstance(d, dt.datetime) else d
                existing.add((str(nm).strip().lower(), d,
                              round(float(cf[f"E{r}"].value or 0), 2)))

    # scheduled recurring occurrences, so a posting can be matched to its bill
    sched = []
    for w in range(1, N_WEEKS + 1):
        b = block(w, var_rows)
        for r in range(b["first_slot"], b["last_slot"] + 1):
            nm = cfv[f"B{r}"].value
            if not nm:
                continue
            d = cfv[f"C{r}"].value
            sched.append({"row": r, "name": nm,
                          "date": d.date() if isinstance(d, dt.datetime) else d,
                          "amount": float(cfv[f"E{r}"].value or 0),
                          "kind": cfv[f"D{r}"].value,
                          "already": cf[f"F{r}"].value is not None})

    matched, leftover = [], []
    used = set()
    for t in kept:
        want_in = t["amount"] > 0
        hit = None
        for s in sched:
            if s["row"] in used or s["date"] is None:
                continue
            if (s["kind"] == "Income") != want_in:
                continue
            if abs(s["amount"] - abs(t["amount"])) > CENTS:
                continue
            if abs((s["date"] - t["date"]).days) > MATCH_DAYS:
                continue
            hit = s
            break
        if hit:
            used.add(hit["row"])
            matched.append((t, hit))
        else:
            leftover.append(t)

    # group the rest by merchant and day
    groups = defaultdict(list)
    for t in leftover:
        groups[(t["date"], t["merchant"])].append(t)

    to_write, dupes = [], []
    for (d, merch), ts in sorted(groups.items()):
        total = round(sum(abs(x["amount"]) for x in ts), 2)
        label = merch if len(ts) == 1 else f"{merch} ({len(ts)} charges)"
        if (label.strip().lower(), d, total) in existing:
            dupes.append((label, d, total)); continue
        to_write.append({"date": d, "label": label, "amount": total,
                         "week": (d - start).days // 7 + 1, "n": len(ts)})

    # place them in free variable rows, week by week
    placed, overflow = [], []
    free = {}
    for w in range(1, N_WEEKS + 1):
        b = block(w, var_rows)
        free[w] = [r for r in range(b["first_var"], b["last_var"] + 1)
                   if cf[f"B{r}"].value is None or str(cf[f"B{r}"].value).startswith("=")]
    for item in to_write:
        w = item["week"]
        if free.get(w):
            item["row"] = free[w].pop(0); placed.append(item)
        else:
            overflow.append(item)

    print(f"start {start:%a %b %d, %Y} · horizon to {end:%b %d} · {var_rows} variable rows/week")
    print(f"csv rows {len(txns)} → in scope {len(kept)}")
    for why, rows in skipped.items():
        print(f"   skipped {len(rows):>3}  {why}")
    print()
    if matched:
        print("matched to a scheduled bill (recorded as Amount Paid / Received):")
        for t, s in matched:
            note = " — already recorded, left alone" if s["already"] else ""
            print(f"   {t['date']:%b %d} {t['merchant']:24s} {abs(t['amount']):>8.2f}"
                  f"  ->  {s['name']} sched {s['date']:%b %d}{note}")
        print()
    if dupes:
        print(f"already in the sheet, not re-added ({len(dupes)}):")
        for l, d, amt in dupes: print(f"   {d:%b %d} {l:30s} {amt:>8.2f}")
        print()
    print(f"variable rows to add ({len(placed)}):")
    tot = 0.0
    for i in placed:
        tot += i["amount"]
        print(f"   W{i['week']:<2} r{i['row']:<4} {i['date']:%a %b %d}  {i['label']:32s} {i['amount']:>8.2f}")
    print(f"   {'total':>47} {tot:>8.2f}")
    if overflow:
        print(f"\n!! no free rows for {len(overflow)}: " +
              ", ".join(f"{i['label']} {i['date']:%b %d}" for i in overflow))

    if a.dry_run:
        print("\ndry run — nothing written")
        return

    for t, s in matched:
        if not s["already"]:
            cf[f"F{s['row']}"] = round(abs(t["amount"]), 2)
    for i in placed:
        cf[f"B{i['row']}"] = i["label"]
        cf[f"C{i['row']}"] = i["date"]
        cf[f"C{i['row']}"].number_format = DATE_F
        cf[f"E{i['row']}"] = i["amount"]

    out = a.out or a.book
    wb.save(out)
    print(f"\nwrote {out}")


if __name__ == "__main__":
    main()
