# xlsx-audit

**Find the mistakes hiding in Excel workbooks, and see exactly what changed between two versions.** From the command line, in CI, or from Python. Nothing is uploaded: your files are read on your own machine.

```
pip install xlsx-audit
xlsx-audit lint model.xlsx
xlsx-audit diff model_v1.xlsx model_v2.xlsx
```

It finds:

- **Numbers typed over formulas**: one cell in a column of `=A*B` that someone overwrote with `535`.
- **Formulas that break the pattern**: `=A5*B6` in a column where every other row is `=A*B` of its own row.
- **Broken references**: `#REF!` in formulas and defined names, even when `IFERROR` hides it.
- **Error values**, traced to the cell they come from, so one `#REF!` doesn't drown you in 300 copies.
- **Links to other workbooks** that break when a file moves.
- **Circular references** where a total includes its own cell.
- **Very hidden sheets**, volatile functions such as `INDIRECT` and `OFFSET`, and magic numbers inside formulas.

And between two versions of a workbook, it tells you **which formulas changed, which were pasted over with values, which rows were inserted**, and which problems are new. Inserting a row doesn't make every cell below it look different.

## Who it's for

- **Analysts and model reviewers** checking a financial model before it goes to a client or a board.
- **Teams that keep workbooks in Git** and want a pull-request check, like a linter for code.
- **Developers** whose apps generate or consume `.xlsx` files and need to test them.

## What it looks like

```
$ xlsx-audit lint model.xlsx
xlsx-audit 0.1.0  model.xlsx
3 sheets, 30 cells, 10 formulas
11 findings: 3 errors, 5 warnings, 3 info

ERROR    XA002 broken-reference (1)
  Sales!D2  Formula refers to deleted cells: =#REF!*2

ERROR    XA008 broken-name (1)
  (workbook)  Defined name OldRange points to #REF!

ERROR    XA010 self-reference (1)
  Sales!C8  Formula includes its own cell (C2:C8), a circular reference: =SUM(C2:C8)

WARNING  XA003 external-link (2)
  (workbook)  Workbook links to another file: file:///C:/Finance/Prices.xlsx
  Sales!E2    Formula uses another workbook (file:///C:/Finance/Prices.xlsx): =[1]Prices!A1*1.15

WARNING  XA004 hardcoded-in-formula-range (1)
  Sales!C4  Typed value 535 inside a run of 4 matching formulas (expected =A4*B4*Inputs!$B$1)

WARNING  XA005 inconsistent-formula (1)
  Sales!C5  =A5*B6*Inputs!$B$1 breaks the pattern of the 4 matching formulas around it (expected =A5*B5*Inputs!$B$1)

WARNING  XA007 hidden-sheet (1)
  Notes  Sheet Notes is very hidden: it can only be shown with VBA

INFO     XA001 error-value (1)
  Sales!D3  Shows #REF! because it uses D2, which also shows an error

INFO     XA006 volatile-function (1)
  Sales!F2  Uses INDIRECT, which recalculates on every change and hides what it reads

INFO     XA009 hardcoded-number (1)
  Sales!E2  Typed number 1.15 inside =[1]Prices!A1*1.15
```

Comparing the next version, where someone inserted a row, pasted a value over a formula, changed the total and renamed a sheet:

```
$ xlsx-audit diff model_v1.xlsx model_v2.xlsx
xlsx-audit 0.1.0  model_v1.xlsx -> model_v2.xlsx
old: 3 sheets, 30 cells, 10 formulas
new: 3 sheets, 31 cells, 9 formulas
6 changes: 2 warnings, 4 info

WARNING  XD021 formula-overwritten (1)
  Sales!C4  Formula =A3*B3*Inputs!$B$1 was replaced by the value 42.8 (was C3 in the old version)

WARNING  XD040 new-issue (1)
  Sales!C4  New XA004 hardcoded-in-formula-range: Typed value 42.8 inside a run of 3 matching formulas (expected =A4*B4*Inputs!$B$1)

INFO     XD003 sheet-renamed (1)
  'Internal notes'  Sheet Notes renamed to Internal notes

INFO     XD010 rows-inserted (1)
  Sales  Row 2 inserted

INFO     XD020 formula-changed (1)
  Sales!C9  Formula changed: =SUM(C2:C8) -> =SUM(C3:C8) (was C8 in the old version)

INFO     XD041 resolved-issue (1)
  Sales!C8  Resolved XA010 self-reference: Formula includes its own cell (C2:C8), a circular reference: =SUM(C2:C8)
```

Every formula below the inserted row moved down one row, and none of them is reported. Only the six real changes are.

## Install

Python 3.9 or newer, on Windows, macOS or Linux.

```
pip install xlsx-audit
```

Or run it without installing, using [uv](https://docs.astral.sh/uv/) or [pipx](https://pipx.pypa.io/):

```
uvx xlsx-audit lint model.xlsx
pipx run xlsx-audit lint model.xlsx
```

## Usage

### Check workbooks

```
xlsx-audit lint model.xlsx                      # one file
xlsx-audit lint reports/*.xlsx                  # several
xlsx-audit lint model.xlsx --select XA004,XA005 # only some rules
xlsx-audit lint model.xlsx --ignore hardcoded-number,volatile-function
```

Rules can be named by code (`XA004`) or name (`hardcoded-in-formula-range`). `xlsx-audit rules` lists them all.

### Compare two versions

```
xlsx-audit diff model_v1.xlsx model_v2.xlsx
xlsx-audit diff old.xlsx new.xlsx --no-values   # only structure and formulas, not input changes
```

The diff lines up rows, or columns when that explains the changes better, before comparing cells. Formulas are compared after moving every reference to where its cells ended up, on any sheet. So `=SUM(C2:C8)` that Excel turned into `=SUM(C3:C9)` when you inserted a row is unchanged, and so is a formula on another sheet that points into the moved rows. Renamed sheets are matched by their content.

It also lints both versions and reports lint findings that are **new** in the second one (XD040) or **resolved** (XD041).

### Output formats

| `--format` | Use it for |
|---|---|
| `text` (default) | Reading in a terminal. Long lists are shortened; `--max-per-rule 0` shows everything. |
| `json` | Scripts and dashboards. Every finding, with sheet, cell, rule, severity and details. |
| `sarif` | GitHub code scanning and other SARIF viewers. |
| `markdown` | Pull-request comments and CI job summaries. |

`--output report.json` writes to a file instead of the screen.

### Exit codes, for CI

| Code | Meaning |
|---|---|
| 0 | Nothing at or above the `--fail-on` level (default: `error`) |
| 1 | At least one finding at or above `--fail-on` |
| 2 | The file couldn't be read, or an option was wrong |

`--fail-on warning` makes warnings fail the build too; `--fail-on never` always exits 0.

## Use it in GitHub Actions

Check every workbook in a pull request and show the results in the code scanning tab:

```yaml
name: Workbooks
on: [pull_request]
jobs:
  lint:
    runs-on: ubuntu-latest
    permissions:
      security-events: write
    steps:
      - uses: actions/checkout@v7
      - uses: actions/setup-python@v7
        with:
          python-version: "3.12"
      - run: pip install xlsx-audit
      - run: xlsx-audit lint models/*.xlsx --format sarif --output xlsx-audit.sarif --fail-on never
      - uses: github/codeql-action/upload-sarif@v4
        with:
          sarif_file: xlsx-audit.sarif
      - run: xlsx-audit lint models/*.xlsx --format markdown >> "$GITHUB_STEP_SUMMARY"
```

## Use it with pre-commit

```yaml
repos:
  - repo: https://github.com/gravitybendr/xlsx-audit
    rev: v0.1.0
    hooks:
      - id: xlsx-audit
```

## Use it from Python

```python
import xlsx_audit

for f in xlsx_audit.audit("model.xlsx"):
    print(f.severity, f.code, f.location, f.message)

for change in xlsx_audit.compare("model_v1.xlsx", "model_v2.xlsx"):
    print(change.code, change.location, change.old, "->", change.new)
```

Each `Finding` has `code`, `slug`, `severity`, `sheet`, `cell`, `message` and `detail` (for example the expected formula). For a diff it also has `old`, `new` and `old_cell`. `xlsx_audit.load()`, `lint()` and `diff()` give you the lower-level steps.

## Rules

### Checks on one workbook

| Code | Name | Severity | What it finds |
|---|---|---|---|
| XA001 | error-value | error, warning or info | A cell shows an Excel error. `#REF!`, `#NAME?` and `#NULL!` are errors; `#DIV/0!`, `#VALUE!` and `#NUM!` are warnings; `#N/A` is info. An error that only comes from another error cell is info and names its source. |
| XA002 | broken-reference | error | A formula contains `#REF!`: it points at cells that were deleted, even if `IFERROR` hides the result. |
| XA003 | external-link | warning | The workbook, a formula or a defined name depends on another workbook file. |
| XA004 | hardcoded-in-formula-range | warning | A typed number inside a run of cells that all use the same formula. Info when it sits at the end of the run. |
| XA005 | inconsistent-formula | warning | A formula that breaks the pattern of the identical formulas around it. The message shows the formula it was expected to have. |
| XA006 | volatile-function | info | `INDIRECT`, `OFFSET`, `NOW`, `TODAY`, `RAND`, `RANDBETWEEN`, `RANDARRAY`, `CELL` or `INFO`. They recalculate on every change, and `INDIRECT` and `OFFSET` hide what they read. |
| XA007 | hidden-sheet | info, or warning if very hidden | A hidden sheet. Very hidden sheets can only be shown through VBA. |
| XA008 | broken-name | error | A defined name that points to `#REF!`. |
| XA009 | hardcoded-number | info | A typed number inside a formula, such as `=B2*1.07`. Common counts like 12, 100 and 365 are ignored. |
| XA010 | self-reference | error | A formula whose own cell is inside a range it reads, such as a total that includes itself. |

### Changes between two versions

| Code | Name | Severity |
|---|---|---|
| XD001, XD002, XD003 | sheet added, removed, renamed | info, warning, info |
| XD010, XD011 | rows inserted, deleted | info |
| XD012, XD013 | columns inserted, deleted | info |
| XD020 | formula changed | info |
| XD021 | formula overwritten with a value | warning |
| XD022, XD023 | formula added, removed | info, warning |
| XD024, XD025, XD026 | value changed, added, removed | info |
| XD030 | defined name changed | info |
| XD031, XD032 | link to another workbook added, removed | warning, info |
| XD040, XD041 | lint finding new, resolved | the finding's own severity, info |

## How the pattern checks work

Formulas are compared in relative (R1C1) form: `=A2*B2` in C2 and `=A3*B3` in C3 are the same formula copied down. In each unbroken run of cells down a column or across a row, the formula most cells share (at least three copies) is the pattern. Inside the stretch it covers, a typed number is flagged (XA004), and so is a formula that appears at most twice (XA005). A rare variant is likely a slip; a variant that repeats is likely a second block on purpose. Cells at the edges of a run, such as a header or a first row that starts a running total, aren't flagged.

## Limits

- **It reads what Excel saved.** Error values (XA001) come from the results Excel stored when the file was last saved. Files written by other programs, including `openpyxl`, often have no stored results; every other check still works on them.
- **`.xlsx` and `.xlsm` only.** Save `.xls` and `.xlsb` files as `.xlsx` first. Password-protected workbooks can't be read.
- **Not checked:** VBA code, Power Query, pivot tables, charts and conditional formatting.
- **Circular references** are found when a formula's own cell is inside a range it reads, not across chains of cells.
- **The diff** lines up rows or columns for each sheet, not both. If one edit inserted rows and columns in the same sheet, some moved cells show up as changed.

## Privacy

Everything runs locally. xlsx-audit makes no network requests and sends nothing anywhere.

## License

MIT
