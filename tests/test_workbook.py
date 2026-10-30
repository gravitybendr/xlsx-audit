import zipfile

import openpyxl
import pytest
from xlsx_builder import F

from xlsx_audit import WorkbookError, load


def test_reads_formulas_values_and_package_parts(wb):
    book = wb(
        {
            "Data": {
                "A1": 2,
                "B1": F("A1*2", 4),
                "C1": F("1/0", "#DIV/0!"),
                "D1": F("SUM(A1:A3*2)", 4, array_ref="D1"),
                "E1": "#N/A",
                "A2": F("A1+1", 3, shared=("A2:A3", 0)),
                "A3": F(cached=4, shared_child=0),
            },
            "Secret": {"A1": "x"},
        },
        states={"Secret": "veryHidden"},
        names={"Rate": "Data!$A$1"},
        sheet_names={"Local": ("Data", "Data!$B$1")},
        external_links=["file:///C:/Prices.xlsx", "Book2.xlsx"],
    )
    assert [s.name for s in book.sheets] == ["Data", "Secret"]
    assert book.sheet("secret").state == "veryHidden"
    data = book.sheet("Data")
    assert data.cells[(1, 2)].formula == "A1*2"
    assert data.cells[(1, 2)].cached == 4
    assert data.cells[(1, 3)].error == "#DIV/0!"
    assert data.cells[(1, 4)].array is True
    assert data.cells[(1, 5)].error == "#N/A" and data.cells[(1, 5)].formula is None
    assert data.cells[(3, 1)].formula == "A2+1", "shared formulas are expanded per cell"
    assert {(n.name, n.scope) for n in book.names} == {("Rate", None), ("Local", "Data")}
    assert book.external_links == ["file:///C:/Prices.xlsx", "Book2.xlsx"]
    assert book.stats == {"sheets": 2, "cells": 8, "formulas": 5}


def test_workbook_saved_by_openpyxl_has_no_cached_results(tmp_path):
    path = tmp_path / "plain.xlsx"
    w = openpyxl.Workbook()
    ws = w.active
    ws["A1"], ws["A2"], ws["A3"] = 1, 2, "=A1+A2"
    w.save(path)
    book = load(path)
    cell = book.sheets[0].cells[(3, 1)]
    assert cell.formula == "A1+A2" and cell.cached is None and cell.error is None


def test_missing_file(tmp_path):
    with pytest.raises(WorkbookError, match="not found"):
        load(tmp_path / "nope.xlsx")


def test_folder(tmp_path):
    with pytest.raises(WorkbookError, match="folder"):
        load(tmp_path)


def test_not_a_workbook(tmp_path):
    p = tmp_path / "notes.xlsx"
    p.write_text("hello")
    with pytest.raises(WorkbookError, match="not an .xlsx"):
        load(p)


def test_encrypted_or_legacy_file(tmp_path):
    p = tmp_path / "locked.xlsx"
    p.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\0" * 512)
    with pytest.raises(WorkbookError, match="password-protected"):
        load(p)


def test_xlsb_is_refused(tmp_path):
    p = tmp_path / "book.xlsb"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
    with pytest.raises(WorkbookError, match="xlsb"):
        load(p)


def test_broken_zip_contents(tmp_path):
    p = tmp_path / "broken.xlsx"
    with zipfile.ZipFile(p, "w") as z:
        z.writestr("hello.txt", "not a workbook")
    with pytest.raises(WorkbookError, match="can't read"):
        load(p)
