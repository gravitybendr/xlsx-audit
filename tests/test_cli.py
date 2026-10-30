import json
import subprocess
import sys

import pytest
from xlsx_builder import F

from xlsx_audit import __version__
from xlsx_audit.cli import main


@pytest.fixture
def clean(make):
    return str(make({"S": {"A1": 1, "B1": F("A1*2", 2)}}, name="clean.xlsx"))


@pytest.fixture
def broken(make):
    return str(make({"S": {"A1": F("#REF!*2", "#REF!"), "A2": F("B2*1.07", 1)}}, name="broken.xlsx"))


def test_clean_workbook_exits_0(clean, capsys):
    assert main(["lint", clean]) == 0
    assert "No findings." in capsys.readouterr().out


def test_findings_exit_1_and_fail_on(broken, capsys):
    assert main(["lint", broken]) == 1
    out = capsys.readouterr().out
    assert "XA002 broken-reference" in out and "S!A1" in out
    assert main(["lint", broken, "--fail-on", "never"]) == 0
    assert main(["lint", broken, "--select", "XA009"]) == 0, "info alone doesn't fail at the default level"
    assert main(["lint", broken, "--select", "XA009", "--fail-on", "info"]) == 1


def test_json_output(broken, capsys):
    main(["lint", broken, "--format", "json"])
    data = json.loads(capsys.readouterr().out)
    rep = data["reports"][0]
    assert data["version"] == __version__
    assert rep["counts"]["error"] == 1
    assert {f["rule"] for f in rep["findings"]} == {"broken-reference", "hardcoded-number"}


def test_sarif_output(broken, capsys):
    main(["lint", broken, "--format", "sarif"])
    sarif = json.loads(capsys.readouterr().out)
    run = sarif["runs"][0]
    assert sarif["version"] == "2.1.0"
    assert {r["id"] for r in run["tool"]["driver"]["rules"]} == {"XA002", "XA009"}
    levels = {r["ruleId"]: r["level"] for r in run["results"]}
    assert levels == {"XA002": "error", "XA009": "note"}
    loc = run["results"][0]["locations"][0]
    assert loc["physicalLocation"]["artifactLocation"]["uri"].endswith("broken.xlsx")
    assert loc["logicalLocations"][0]["fullyQualifiedName"] == "S!A1"


def test_markdown_and_output_file(broken, tmp_path):
    out = tmp_path / "report.md"
    main(["lint", broken, "--format", "markdown", "--output", str(out)])
    text = out.read_text(encoding="utf-8")
    assert text.startswith("### xlsx-audit:")
    assert "| error | XA002 broken-reference | `S!A1` |" in text


def test_several_files(clean, broken, capsys):
    assert main(["lint", clean, broken]) == 1
    out = capsys.readouterr().out
    assert out.count("xlsx-audit 0.1.0") == 2


def test_wildcards_are_expanded(clean, broken, tmp_path, capsys):
    assert main(["lint", str(tmp_path / "*.xlsx")]) == 1
    out = capsys.readouterr().out
    assert "broken.xlsx" in out and "clean.xlsx" in out
    assert main(["lint", str(tmp_path / "*.xlsm")]) == 2
    assert "no files match" in capsys.readouterr().err


def test_max_per_rule(make, capsys):
    cells = {f"A{r}": F(f"B{r}*1.07", 1) for r in range(1, 8)}
    path = str(make({"S": cells}))
    main(["lint", path, "--max-per-rule", "3"])
    assert "... and 4 more" in capsys.readouterr().out
    main(["lint", path, "--max-per-rule", "0"])
    assert "more" not in capsys.readouterr().out


def test_diff_command(make, capsys):
    old = str(make({"S": {"A1": 1, "B1": F("A1*2", 2)}}, name="v1.xlsx"))
    new = str(make({"S": {"A1": 1, "B1": 2}}, name="v2.xlsx"))
    assert main(["diff", old, new]) == 0
    out = capsys.readouterr().out
    assert "XD021 formula-overwritten" in out
    assert main(["diff", old, new, "--fail-on", "warning"]) == 1


def test_rules_command(capsys):
    assert main(["rules"]) == 0
    out = capsys.readouterr().out
    assert "XA001" in out and "XD040" in out


def test_errors_exit_2(tmp_path, clean, capsys):
    assert main(["lint", str(tmp_path / "missing.xlsx")]) == 2
    assert "file not found" in capsys.readouterr().err
    assert main(["lint", clean, "--select", "nonsense"]) == 2
    assert "unknown rule" in capsys.readouterr().err


def test_module_entry_point(clean):
    proc = subprocess.run([sys.executable, "-m", "xlsx_audit", "--version"], capture_output=True, text=True)
    assert proc.returncode == 0 and proc.stdout.strip() == f"xlsx-audit {__version__}"
