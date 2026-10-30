from conftest import at, codes
from xlsx_builder import F

from xlsx_audit import diff


def table(rows, start=2, extra=None, total=True):
    """Sales-like table: A qty, B price, C =A*B, and a total under it."""
    cells = {"A1": "Qty", "B1": "Price", "C1": "Revenue"}
    r = start
    for qty in rows:
        cells[f"A{r}"], cells[f"B{r}"] = qty, 10
        cells[f"C{r}"] = F(f"A{r}*B{r}", qty * 10)
        r += 1
    if total:
        cells[f"C{r}"] = F(f"SUM(C{start}:C{r - 1})", 1)
    cells.update(extra or {})
    return cells


def changes(result, *exclude):
    return [f for f in result.findings if f.code not in exclude]


def test_identical_versions(wb):
    a = wb({"S": table([1, 2, 3])})
    b = wb({"S": table([1, 2, 3])})
    assert diff(a, b).findings == []


def test_inserted_row_is_one_change(wb):
    old = wb({"S": table([1, 2, 3, 4])})
    new_cells = table([1, 2, 3, 4], start=3)
    new_cells["A2"] = "Q1"  # the inserted row has a label
    new = wb({"S": new_cells})
    result = diff(old, new)
    assert codes(result.findings) == ["XD010"], [f.message for f in result.findings]
    assert result.findings[0].message == "Row 2 inserted"


def test_row_inserted_inside_a_summed_range(wb):
    old = wb({"S": table([1, 2, 3, 4])})  # total in C6 = SUM(C2:C5)
    new_cells = table([1, 2, 3, 4])
    # Excel moves rows 4+ down and grows the total's range.
    new_cells = {k: v for k, v in new_cells.items() if not k.endswith(("4", "5", "6"))}
    new_cells.update({"A4": 9, "B4": 10, "C4": F("A4*B4", 90)})
    for r_old, r_new, qty in ((4, 5, 3), (5, 6, 4)):
        new_cells.update({f"A{r_new}": qty, f"B{r_new}": 10, f"C{r_new}": F(f"A{r_new}*B{r_new}", qty * 10)})
    new_cells["C7"] = F("SUM(C2:C6)", 1)
    result = diff(old, wb({"S": new_cells}))
    assert codes(result.findings) == ["XD010"], [f.message for f in result.findings]


def test_inserted_column(wb):
    old = wb({"S": {"A1": 1, "B1": F("A1*2", 2), "A2": 3, "B2": F("A2*2", 6)}})
    new = wb({"S": {"A1": 1, "B1": "note", "C1": F("A1*2", 2), "A2": 3, "C2": F("A2*2", 6)}})
    result = diff(old, new)
    assert codes(result.findings) == ["XD012"], [f.message for f in result.findings]
    assert result.findings[0].message == "Column B inserted"


def test_reference_from_another_sheet_follows_inserted_rows(wb):
    old = wb({"Sales": table([1, 2, 3]), "Summary": {"B1": F("Sales!C5", 60)}})
    new_sales = table([1, 2, 3], start=3)
    new_sales["A2"] = "Q1"
    new = wb({"Sales": new_sales, "Summary": {"B1": F("Sales!C6", 60)}})
    result = diff(old, new)
    assert not at(result.findings, "XD020"), [f.message for f in result.findings]


def test_formula_overwritten_with_a_value(wb):
    old = wb({"S": table([1, 2, 3])})
    new = wb({"S": table([1, 2, 3], extra={"C3": 20})})
    result = diff(old, new)
    hit = at(result.findings, "XD021", "C3")[0]
    assert hit.severity == "warning"
    assert hit.old == "=A3*B3" and hit.new == "20"


def test_overwrite_after_a_move_names_the_old_cell(wb):
    old = wb({"S": table([1, 2, 3, 4])})
    new_cells = table([1, 2, 3, 4], start=3, extra={"C4": 20})
    new_cells["A2"] = "Q1"
    hit = at(diff(old, wb({"S": new_cells})).findings, "XD021")[0]
    assert hit.cell == "C4" and hit.old_cell == "C3"
    assert "was C3 in the old version" in hit.message


def test_formula_and_value_changes(wb):
    old = wb({"S": {"A1": 1, "A2": 2, "A3": 3, "B1": F("A1*2", 2), "B2": F("A2*2", 4), "B3": 7}})
    new = wb({"S": {"A1": 5, "A3": 3, "A4": 9, "B1": F("A1*3", 15), "B3": F("A3*2", 6)}})
    result = diff(old, new)
    kinds = {(f.code, f.cell) for f in result.findings if f.code.startswith("XD02")}
    assert kinds == {
        ("XD024", "A1"),  # value changed
        ("XD026", "A2"),  # value deleted
        ("XD025", "A4"),  # value added
        ("XD020", "B1"),  # formula changed
        ("XD023", "B2"),  # formula deleted
        ("XD022", "B3"),  # formula added over a value
    }
    no_values = diff(old, new, include_values=False)
    assert not {"XD024", "XD025", "XD026"} & set(codes(no_values.findings))


def test_sheets_added_removed_and_renamed(wb):
    old = wb({"Main": {"A1": F("Notes!A1", 1)}, "Notes": {"A1": 1, "A2": 2}, "Old": {"A1": "x"}})
    new = wb({"Main": {"A1": F("'Read me'!A1", 1)}, "Read me": {"A1": 1, "A2": 2}, "Fresh": {"A1": "y"}})
    result = diff(old, new)
    assert at(result.findings, "XD003")[0].message == "Sheet Notes renamed to Read me"
    assert at(result.findings, "XD002")[0].sheet == "Old"
    assert at(result.findings, "XD001")[0].sheet == "Fresh"
    assert not at(result.findings, "XD020"), "a formula pointing at the renamed sheet is unchanged"
    assert result.renamed == {"Notes": "Read me"}


def test_names_and_links(wb):
    old = wb({"S": {"A1": 1}}, names={"Rate": "S!$A$1", "Gone": "S!$A$2"})
    new = wb({"S": {"A1": 1}}, names={"Rate": "S!$A$3", "New": "S!$A$4"}, external_links=["file:///C:/x.xlsx"])
    messages = [f.message for f in at(diff(old, new).findings, "XD030")]
    assert "Defined name Rate changed: S!$A$1 -> S!$A$3" in messages
    assert any(m.startswith("Defined name New added") for m in messages)
    assert any(m.startswith("Defined name Gone removed") for m in messages)
    assert at(diff(old, new).findings, "XD031")[0].severity == "warning"


def test_new_and_resolved_issues(wb):
    old = wb({"S": table([1, 2, 3, 4], extra={"D1": F("#REF!+1", "#REF!")})})
    new = wb({"S": table([1, 2, 3, 4], extra={"C3": 99})})
    result = diff(old, new)
    new_issue = at(result.findings, "XD040", "C3")[0]
    assert new_issue.detail["rule"] == "XA004"
    resolved = at(result.findings, "XD041")
    assert any(f.detail["rule"] == "XA002" for f in resolved)


def test_issue_on_a_moved_cell_is_not_new(wb):
    old = wb({"S": table([1, 2, 3, 4], extra={"C3": 99})})
    new_cells = table([1, 2, 3, 4], start=3, extra={"C4": 99})
    new_cells["A2"] = "Q1"
    result = diff(old, wb({"S": new_cells}))
    assert not at(result.findings, "XD040") and not at(result.findings, "XD041")
