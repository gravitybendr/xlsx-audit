import pytest

from xlsx_audit.refs import (
    MAX_COL,
    MAX_ROW,
    canonical,
    cell_name,
    external_book,
    functions,
    has_ref_error,
    is_benign_number,
    normalize,
    numbers,
    parse_cell,
    ref_to_r1c1,
    references,
    remap,
    split_sheet,
    unquote_sheet,
)


def test_cell_names_round_trip():
    assert parse_cell("C7") == (7, 3)
    assert parse_cell("xfd1048576") == (1048576, 16384)
    assert cell_name(7, 3) == "C7"
    with pytest.raises(ValueError):
        parse_cell("Rate")


@pytest.mark.parametrize(
    "operand,expected",
    [
        ("A1", ("", "A1")),
        ("Data!A1", ("Data!", "A1")),
        ("'My Sheet'!A1:B2", ("'My Sheet'!", "A1:B2")),
        ("'It''s!'!B3", ("'It''s!'!", "B3")),
        ("[1]Prices!A1", ("[1]Prices!", "A1")),
        ("'C:\\Data\\[Prices.xlsx]Sheet1'!$A$1", ("'C:\\Data\\[Prices.xlsx]Sheet1'!", "$A$1")),
    ],
)
def test_split_sheet(operand, expected):
    assert split_sheet(operand) == expected


def test_unquote_and_external_book():
    assert unquote_sheet("'It''s'!") == "It's"
    assert unquote_sheet("Data!") == "Data"
    assert external_book("[1]Prices!") == "1"
    assert external_book("'C:\\Data\\[Prices.xlsx]Sheet1'!") == "C:\\Data\\Prices.xlsx"
    assert external_book("Data!") is None
    assert external_book("'My Sheet'!") is None


@pytest.mark.parametrize(
    "ref,row,col,expected",
    [
        ("A1", 1, 1, "RC"),
        ("B3", 2, 1, "R[1]C[1]"),
        ("$A$1", 5, 5, "R1C1"),
        ("A$1", 5, 2, "R1C[-1]"),
        ("$A1", 5, 2, "R[-4]C1"),
        ("A1:B2", 1, 1, "RC:R[1]C[1]"),
        ("A:A", 5, 3, "C[-2]"),
        ("$A:$B", 5, 3, "C1:C2"),
        ("3:3", 5, 1, "R[-2]"),
        ("Rate", 1, 1, None),
        ("Table1[Col]", 1, 1, None),
    ],
)
def test_ref_to_r1c1(ref, row, col, expected):
    if ":" in ref and expected and ":" not in expected:
        expected = f"{expected}:{expected}"
    assert ref_to_r1c1(ref, row, col) == expected


def test_normalize_matches_copied_formulas():
    assert normalize("A2*B2", 2, 3) == normalize("A3*B3", 3, 3)
    assert normalize("A2*B2", 2, 3) != normalize("A3*B4", 3, 3)
    assert normalize("SUM($B$2:B2)", 2, 3) == normalize("SUM($B$2:B9)", 9, 3)


def test_normalize_ignores_spacing_case_and_quoting():
    assert normalize("sum( a1 , b1 )", 1, 3) == normalize("SUM(A1,B1)", 1, 3)
    assert normalize("Data!A1", 1, 1) == normalize("'Data'!A1", 1, 1)
    assert normalize('"a"&A1', 1, 2) != normalize('"A"&A1', 1, 2), "text in quotes keeps its case"


def test_normalize_sheet_alias_and_own_sheet_mode():
    assert normalize("Old!A1", 1, 1, {"Old": "New"}) == normalize("New!A1", 1, 1)
    # Moving the formula's cell doesn't change references to other sheets.
    assert normalize("Inputs!B2", 2, 1, own_sheet="Sales") == normalize("Inputs!B2", 3, 1, own_sheet="Sales")
    assert normalize("Inputs!B2", 2, 1) != normalize("Inputs!B2", 3, 1)
    # References to the formula's own sheet are still relative.
    assert normalize("Sales!B2", 2, 1, own_sheet="Sales") == normalize("Sales!B3", 3, 1, own_sheet="Sales")


def test_references():
    refs = references("SUM(A1:B3)+'My Sheet'!C4+[1]Prices!A1+Rate+A:A+2:3")
    by_text = {r.text: r for r in refs}
    assert by_text["A1:B3"].area == (1, 1, 3, 2)
    assert by_text["A1:B3"].sheet is None
    assert by_text["'My Sheet'!C4"].sheet == "My Sheet"
    assert by_text["'My Sheet'!C4"].area == (4, 3, 4, 3)
    assert by_text["[1]Prices!A1"].book == "1"
    assert by_text["[1]Prices!A1"].sheet is None
    assert by_text["Rate"].area is None
    assert by_text["A:A"].area == (1, 1, MAX_ROW, 1)
    assert by_text["2:3"].area == (2, 1, 3, MAX_COL)


def test_functions_numbers_and_ref_errors():
    assert functions("_xlfn.XLOOKUP(A1,B:B,C:C)+SUM(1)") == ["XLOOKUP", "SUM"]
    assert numbers("A1*1.07+12-3.5%") == ["1.07", "12", "3.5"]
    assert has_ref_error("#REF!*2")
    assert has_ref_error("SUM(Data!#REF!)")
    assert not has_ref_error('IF(A1="#REF!",1,2)'), "#REF! inside text is fine"


def test_benign_numbers():
    assert is_benign_number("12")
    assert is_benign_number("100")
    assert not is_benign_number("1.07")
    assert not is_benign_number("17")


def test_remap_moves_references_through_the_row_map():
    maps = {"sales": ("Sales", 0, {1: 1, 2: 3, 3: 4, 8: 9})}
    # A row was inserted at 2: the range grows with it.
    assert remap("SUM(C2:C8)", "Sales", maps) == canonical("SUM(C3:C9)")
    # References from another sheet move too.
    assert remap("Sales!C2*2", "Summary", maps) == canonical("Sales!C3*2")
    # Absolute markers are kept; external references untouched.
    assert remap("$C$2+[1]Prices!A2", "Sales", maps) == canonical("$C$3+[1]Prices!A2")
    # Rows past the end of the data move with the last row: Excel shifts them too.
    assert remap("SUM(C2:C100)", "Sales", maps) == canonical("SUM(C3:C101)")
    assert remap("SUM(Sales!20:30)", "Summary", maps) == canonical("SUM(Sales!21:31)")


def test_remap_columns_and_renamed_sheets():
    maps = {"notes": ("Internal notes", 1, {1: 2, 2: 3})}
    assert remap("Notes!A1+SUM(Notes!A:B)", "Sales", maps, {"Notes": "Internal notes"}) == canonical(
        "'Internal notes'!B1+SUM('Internal notes'!B:C)"
    )
