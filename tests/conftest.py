import itertools

import pytest
from xlsx_builder import build_xlsx

from xlsx_audit import load

_counter = itertools.count()


@pytest.fixture
def make(tmp_path):
    """make(sheets, **opts) -> path of a new .xlsx built by xlsx_builder."""

    def _make(sheets, name=None, **opts):
        path = tmp_path / (name or f"book{next(_counter)}.xlsx")
        return build_xlsx(path, sheets, **opts)

    return _make


@pytest.fixture
def wb(make):
    """wb(sheets, **opts) -> loaded Workbook."""

    def _wb(sheets, **opts):
        return load(make(sheets, **opts))

    return _wb


def codes(findings):
    return [f.code for f in findings]


def at(findings, code, cell=None, sheet=None):
    return [f for f in findings if f.code == code and (cell is None or f.cell == cell) and (sheet is None or f.sheet == sheet)]
