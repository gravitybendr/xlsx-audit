"""Build the two demo workbooks used in the README and the images.

    python scripts/demo_workbooks.py [folder]      (default: docs/demo)

model_v1.xlsx is a small sales model with one of each common mistake.
model_v2.xlsx is the next version: a row inserted, a formula pasted over
with a value, the total fixed and a sheet renamed. All numbers are made up.

The files are written from raw SpreadsheetML by tests/xlsx_builder.py, the
same builder the tests use, because openpyxl can't write the cached results,
error values and external links these examples need.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tests"))

from xlsx_builder import F, build_xlsx  # noqa: E402

RATE = 1.08  # Inputs!B1
UNITS = [120, 85, 64, 150, 42, 97]  # rows 2-7 of v1
PRICES = [4.5, 6.25, 8.0, 3.75, 12.0, 5.5]
LINK = "file:///C:/Finance/Prices.xlsx"


def _round(x: float) -> float:
    return round(x, 2)


def v1_sheets() -> dict:
    sales: dict[str, object] = {
        "A1": "Units", "B1": "Price", "C1": "Revenue", "D1": "Adjusted", "E1": "List price", "F1": "Lookup",
    }
    revenue = []
    for i, (u, p) in enumerate(zip(UNITS, PRICES)):
        r = i + 2
        sales[f"A{r}"] = u
        sales[f"B{r}"] = p
        value = _round(u * p * RATE)
        sales[f"C{r}"] = F(f"A{r}*B{r}*Inputs!$B$1", value)
        revenue.append(value)
    # Someone typed a number over the formula in C4...
    sales["C4"] = 535
    # ...and C5 picks up the price from the row below.
    sales["C5"] = F("A5*B6*Inputs!$B$1", _round(UNITS[3] * PRICES[4] * RATE))
    # The total includes its own cell.
    sales["C8"] = F("SUM(C2:C8)", 0)
    # A column that points at deleted cells, and a cell that inherits the error.
    sales["D2"] = F("#REF!*2", "#REF!")
    sales["D3"] = F("D2+C3", "#REF!")
    # A price read from another workbook, with a typed mark-up.
    sales["E2"] = F("[1]Prices!A1*1.15", 5.18)
    # A volatile lookup.
    sales["F2"] = F('INDIRECT("Inputs!B1")', RATE)
    return {
        "Sales": sales,
        "Inputs": {"A1": "FX rate", "B1": RATE},
        "Notes": {"A1": "Old assumptions, do not use"},
    }


def v2_sheets() -> dict:
    v1 = v1_sheets()
    old = v1["Sales"]
    sales: dict[str, object] = {}
    # Insert a row under the header: every cell from row 2 down moves one row.
    for ref, value in old.items():
        col = ref.rstrip("0123456789")
        row = int(ref[len(col):])
        if row == 1:
            sales[ref] = value
            continue
        if isinstance(value, F):
            value = F(_shift(value.text), value.cached)
        sales[f"{col}{row + 1}"] = value
    sales["A2"] = 40  # the new row: units typed, price and formula not filled in yet
    # A formula pasted over with its value.
    sales["C4"] = 42.8
    # The total, fixed so it no longer includes itself.
    sales["C9"] = F("SUM(C3:C8)", 1990.5)
    return {
        "Sales": sales,
        "Inputs": v1["Inputs"],
        "Internal notes": v1["Notes"],
    }


def _shift(formula: str) -> str:
    """Move same-sheet row references down by one, as Excel does when a row is inserted at row 2."""
    import re

    def bump(m: re.Match) -> str:
        sheet, col, dollar, row = m.group(1) or "", m.group(2), m.group(3), int(m.group(4))
        if sheet or dollar or row < 2:
            return m.group(0)
        return f"{col}{row + 1}"

    return re.sub(r"((?:\[\d+\])?[A-Za-z]+!)?(\$?[A-Z]{1,3})(\$?)(\d+)", bump, formula)


def build(folder: Path) -> tuple[Path, Path]:
    folder.mkdir(parents=True, exist_ok=True)
    common = {"states": None, "names": {"OldRange": "#REF!"}, "external_links": [LINK]}
    v1 = build_xlsx(folder / "model_v1.xlsx", v1_sheets(), **{**common, "states": {"Notes": "veryHidden"}})
    v2 = build_xlsx(folder / "model_v2.xlsx", v2_sheets(), **{**common, "states": {"Internal notes": "veryHidden"}})
    return v1, v2


if __name__ == "__main__":
    out = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "docs" / "demo"
    for p in build(out):
        print(p)
