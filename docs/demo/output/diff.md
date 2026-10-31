### xlsx-audit: `model_v1.xlsx` -> `model_v2.xlsx`

**0** error | **2** warning | **4** info

| Severity | Rule | Where | What |
|---|---|---|---|
| warning | XD021 formula-overwritten | `Sales!C4` | Formula =A3*B3*Inputs!$B$1 was replaced by the value 42.8 (was C3 in the old version) |
| warning | XD040 new-issue | `Sales!C4` | New XA004 hardcoded-in-formula-range: Typed value 42.8 inside a run of 3 matching formulas (expected =A4*B4*Inputs!$B$1) |
| info | XD010 rows-inserted | `Sales` | Row 2 inserted |
| info | XD041 resolved-issue | `Sales!C8` | Resolved XA010 self-reference: Formula includes its own cell (C2:C8), a circular reference: =SUM(C2:C8) |
| info | XD020 formula-changed | `Sales!C9` | Formula changed: =SUM(C2:C8) -> =SUM(C3:C8) (was C8 in the old version) |
| info | XD003 sheet-renamed | `'Internal notes'` | Sheet Notes renamed to Internal notes |
