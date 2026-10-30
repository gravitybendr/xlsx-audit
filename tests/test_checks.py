import pytest
from conftest import at, codes
from xlsx_builder import F

from xlsx_audit import lint


def column(formula_rows, col="C", start=2, cached=1, overrides=None):
    """Cells for a column of copied formulas =A{r}*B{r}, with optional overrides."""
    cells = {}
    for r in range(start, start + formula_rows):
        cells[f"A{r}"] = r
        cells[f"B{r}"] = 10
        cells[f"{col}{r}"] = F(f"A{r}*B{r}", cached)
    cells.update(overrides or {})
    return cells


# XA001 error-value

def test_error_values_by_kind(wb):
    book = wb({"S": {"A1": F("1/0", "#DIV/0!"), "A2": F("NA()", "#N/A"), "A3": F("FOO(1)", "#NAME?"), "A4": "#VALUE!"}})
    f = lint(book)
    assert at(f, "XA001", "A1")[0].severity == "warning"
    assert at(f, "XA001", "A2")[0].severity == "info"
    assert at(f, "XA001", "A3")[0].severity == "error"
    assert "typed or pasted" in at(f, "XA001", "A4")[0].message


def test_inherited_errors_point_at_their_source(wb):
    book = wb({"S": {"A1": F("1/0", "#DIV/0!"), "B1": F("A1*2", "#DIV/0!"), "C1": F("Other!A1", "#DIV/0!")}, "Other": {"A1": F("1/0", "#DIV/0!")}})
    f = lint(book)
    b1 = at(f, "XA001", "B1")[0]
    assert b1.severity == "info" and b1.detail["source"] == "A1"
    assert at(f, "XA001", "C1")[0].detail["source"] == "Other!A1"
    assert at(f, "XA001", "A1", "S")[0].severity == "warning"


# XA002 broken-reference

def test_broken_reference_even_when_hidden_by_iferror(wb):
    book = wb({"S": {"A1": F("IFERROR(#REF!*2,0)", 0), "A2": F("#REF!+1", "#REF!")}})
    f = lint(book)
    assert at(f, "XA002", "A1") and at(f, "XA002", "A2")
    assert not at(f, "XA001", "A2"), "#REF! from a broken formula is reported once, as XA002"


# XA003 external-link

def test_external_links(wb):
    book = wb({"S": {"A1": F("[1]Prices!A1*2", 4), "A2": F("'C:\\Data\\[Old.xlsx]Sheet1'!B2", 1)}},
              names={"Ext": "[1]Prices!$A$1"}, external_links=["file:///C:/Prices.xlsx"])
    f = at(lint(book), "XA003")
    messages = " | ".join(x.message for x in f)
    assert "Workbook links to another file: file:///C:/Prices.xlsx" in messages
    assert at(f, "XA003", "A1")[0].detail["targets"] == ["file:///C:/Prices.xlsx"]
    assert at(f, "XA003", "A2")[0].detail["targets"] == ["C:\\Data\\Old.xlsx"]
    assert any("Defined name Ext" in x.message for x in f)


# XA004 hardcoded-in-formula-range

def test_typed_value_inside_a_run(wb):
    f = lint(wb({"S": column(6, overrides={"C4": 535})}))
    hit = at(f, "XA004", "C4")[0]
    assert hit.severity == "warning"
    assert hit.detail["expected"] == "=A4*B4"


def test_typed_value_between_two_formulas(wb):
    f = lint(wb({"S": {"A1": F("B1*2", 2), "A2": 7, "A3": F("B3*2", 2)}}))
    assert at(f, "XA004", "A2")[0].detail["expected"] == "=B2*2"


def test_typed_value_ending_a_run_is_info(wb):
    f = lint(wb({"S": column(4, overrides={"C6": 99, "A6": 6, "B6": 10})}))
    assert at(f, "XA004", "C6")[0].severity == "info"


def test_adjacent_problems_are_both_found(wb):
    f = lint(wb({"S": column(6, overrides={"C4": 535, "C5": F("A5*B6", 1)})}))
    assert at(f, "XA004", "C4") and at(f, "XA005", "C5")


def test_year_headers_above_formulas_are_not_flagged(wb):
    cells = {"B1": 2024, "C1": 2025, "D1": 2026}
    for col in "BCD":
        for r in range(2, 6):
            cells[f"{col}{r}"] = F(f"{col}{r - 1}*1", 1) if r > 2 else 5
    f = lint(wb({"S": cells}))
    assert not at(f, "XA004")


# XA005 inconsistent-formula

def test_rare_variant_in_a_run(wb):
    f = lint(wb({"S": column(8, overrides={"C6": F("A6*B7", 1)})}))
    hit = at(f, "XA005", "C6")[0]
    assert hit.detail["expected"] == "=A6*B6"


def test_two_deliberate_blocks_are_not_flagged(wb):
    cells = column(4)  # rows 2-5: =A*B
    for r in range(6, 10):
        cells[f"A{r}"], cells[f"B{r}"] = r, 10
        cells[f"C{r}"] = F(f"C{r - 1}*1.1", 1)
    assert not at(lint(wb({"S": cells})), "XA005")


def test_running_total_first_row_is_not_flagged(wb):
    cells = {"B2": 1, "C2": F("B2", 1)}
    for r in range(3, 8):
        cells[f"B{r}"] = 1
        cells[f"C{r}"] = F(f"C{r - 1}+B{r}", 1)
    assert not at(lint(wb({"S": cells})), "XA005")


def test_formula_across_a_row(wb):
    cells = {"A1": 100}
    for i, col in enumerate("BCDEFG"):
        prev = "ABCDEF"[i]
        cells[f"{col}1"] = F(f"{prev}1*1.05", 1)
    cells["E1"] = F("D1*1.5", 1)
    assert at(lint(wb({"S": cells})), "XA005", "E1")


# XA006, XA007, XA008, XA009, XA010

def test_volatile_functions(wb):
    f = lint(wb({"S": {"A1": F('INDIRECT("B1")', 1), "A2": F("TODAY()", 45000), "A3": F("SUM(B1:B2)", 0)}}))
    assert "hides what it reads" in at(f, "XA006", "A1")[0].message
    assert at(f, "XA006", "A2") and not at(f, "XA006", "A3")


def test_hidden_sheets(wb):
    f = lint(wb({"A": {"A1": 1}, "B": {"A1": 1}, "C": {"A1": 1}}, states={"B": "hidden", "C": "veryHidden"}))
    assert at(f, "XA007", sheet="B")[0].severity == "info"
    assert at(f, "XA007", sheet="C")[0].severity == "warning"
    assert not at(f, "XA007", sheet="A")


def test_broken_names(wb):
    f = lint(wb({"S": {"A1": 1}}, names={"Good": "S!$A$1", "Bad": "#REF!", "_xlnm._FilterDatabase": "S!#REF!"}, sheet_names={"Local": ("S", "S!#REF!")}))
    names = sorted(x.detail["name"] for x in at(f, "XA008"))
    assert names == ["Bad", "Local"]


def test_typed_numbers_in_formulas(wb):
    f = lint(wb({"S": {"A1": F("B1*1.07", 1), "A2": F("B2/12", 1), "A3": F("B3*100", 1)}}))
    assert at(f, "XA009", "A1")[0].detail["numbers"] == ["1.07"]
    assert not at(f, "XA009", "A2") and not at(f, "XA009", "A3")


def test_self_reference(wb):
    f = lint(wb({"S": {"C8": F("SUM(C2:C8)", 0), "A5": F("SUM(A:A)", 0), "B1": F("Other!B1", 0)}, "Other": {"B1": 1}}))
    assert at(f, "XA010", "C8") and at(f, "XA010", "A5")
    assert not at(f, "XA010", "B1"), "same address on another sheet isn't circular"


# Selecting rules

def test_select_and_ignore_by_code_or_name(wb):
    book = wb({"S": {"A1": F('INDIRECT("B1")', 1), "A2": F("B1*1.07", 1)}})
    assert set(codes(lint(book, select={"XA006"}))) == {"XA006"}
    assert set(codes(lint(book, select={"hardcoded-number"}))) == {"XA009"}
    assert "XA006" not in codes(lint(book, ignore={"volatile-function"}))
    with pytest.raises(ValueError, match="unknown rule"):
        lint(book, select={"XA999"})


def test_clean_workbook_has_no_findings(wb):
    assert lint(wb({"S": column(5)})) == []
