"""Build two small workbooks: one clean, one with the six defects audit.py looks for.
Run: python make_samples.py
"""
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent
MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun"]
SALES = [100, 110, 121, 133, 146, 161]


def build(broken):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "PnL"
    ws["A1"] = "Six-month P&L (illustrative figures)"
    ws.merge_cells("A1:H1")
    ws.append(["", *MONTHS, "Total"])
    ws.append(["Sales", *SALES, "=SUM(B3:G3)"])
    ws.append(["Cost of sales", *[f"=-{c}3*Inputs!$B$2" for c in "BCDEFG"], "=SUM(B4:G4)"])
    ws.append(["Gross margin", *[f"={c}3+{c}4" for c in "BCDEFG"], "=SUM(B5:G5)"])
    ws.append(["Margin %", *[f"={c}5/{c}3" for c in "BCDEFGH"]])
    for c in "BCDEFGH":
        ws[f"{c}6"].number_format = "0.0%"
        ws[f"{c}5"].number_format = '#,##0;(#,##0);"–"'

    inp = wb.create_sheet("Inputs")
    inp.append(["Input", "Value"])
    inp.append(["Cost of sales, % of sales", 0.62])
    inp.append(["Fixed costs per month", 25])
    ws["A8"] = "Fixed costs"
    for c in "BCDEFG":
        ws[f"{c}8"] = "=-Inputs!$B$3"
    ws["A9"] = "Operating result"
    for c in "BCDEFG":
        ws[f"{c}9"] = f"={c}5+{c}8"

    if broken:
        ws["D4"] = -70                                   # 1 typed number inside a row of formulas
        ws["A11"] = "= margin after fixed costs"         # 2 note saved as a formula
        ws.merge_cells("A1:I1")                          # 3 second merge on top of A1:H1
        inp["C3"] = 25
        inp.merge_cells("B3:C3")                         #   the merge wipes C3
        ws["B8"] = "=-Inputs!$C$3"                       # 4 points to a cell wiped by a merge
        ws["H9"] = "=H5/Inputs!$D$9"                     # 5 division by an empty cell
        ws["H5"].number_format = "#,##0;(#,##0);–"       # 6 unquoted dash in a number format
    name = "sample_broken.xlsx" if broken else "sample_clean.xlsx"
    wb.save(HERE / name)
    print("saved", name)


if __name__ == "__main__":
    build(False)
    build(True)
