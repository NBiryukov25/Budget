#!/usr/bin/env python3
"""Third pass: typing aids, protection of the engine, print set-up."""
import argparse
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.properties import PageSetupProperties

SLOTS, FIRST_BLOCK, N_WEEKS = 10, 13, 13
REC_FIRST, REC_LAST = 5, 34
LOG_FIRST, LOG_LAST = 9, 308
OVR_FIRST, OVR_LAST = 9, 108
SM_LAST = 25


def detect_var_rows(cf):
    for vr in range(4, 41):
        v = cf[f"B{FIRST_BLOCK + (SLOTS + vr + 5)}"].value
        if isinstance(v, str) and v.startswith('="WEEK '):
            return vr
    raise SystemExit("cannot determine block height")


def fit(ws, area, landscape=True, repeat=None):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.print_area = area
    if repeat:
        ws.print_title_rows = repeat


def main():
    ap = argparse.ArgumentParser(); ap.add_argument("src"); ap.add_argument("out")
    a = ap.parse_args()
    wb = load_workbook(a.src)
    cf, rec, sm, plog, ovr, sch, ins = (wb["Cash Flow"], wb["Recurring"], wb["Summary"],
                                        wb["Paid Log"], wb["Overrides"], wb["Schedule"],
                                        wb["Instructions"])
    VR = detect_var_rows(cf)

    # 11 · pick a transaction from a list instead of retyping its name
    names = f"=Recurring!$B${REC_FIRST}:$B${REC_LAST}"
    for ws, col, lo, hi in ((plog, "C", LOG_FIRST, LOG_LAST),
                            (ovr, "B", OVR_FIRST, OVR_LAST)):
        dv = DataValidation(type="list", formula1=names, allow_blank=True,
                            errorStyle="warning")
        dv.errorTitle = "Not a listed transaction"
        dv.error = ("Nothing on the Recurring tab is called that, so it will not match. "
                    "Choose Yes to enter it anyway.")
        ws.add_data_validation(dv)
        dv.add(f"{col}{lo}:{col}{hi}")

    # 12 · only the calculation engine is actually locked, and without a password
    sch.protection.sheet = True          # no password: Review > Unprotect clears it
    sch.protection.enable()
    sch.sheet_state = "visible"

    # 10 · group the engine's working columns so they fold away
    for col in "HIJKLMN":
        sch.column_dimensions[col].outlineLevel = 1
    sch.sheet_properties.outlinePr.summaryRight = True

    last_tot = FIRST_BLOCK + (SLOTS + VR + 5) * (N_WEEKS - 1) + 3 + SLOTS + VR
    fit(cf, f"A1:N{last_tot}")
    fit(rec, f"A1:K{REC_LAST + 4}", repeat="4:4")
    fit(sm, f"A1:L{SM_LAST + 3}")
    fit(plog, f"A1:F{LOG_FIRST + 60}")
    fit(ovr, f"A1:F{OVR_FIRST + 40}")

    wb.save(a.out)
    print("stage 3 saved ->", a.out)


if __name__ == "__main__":
    main()
