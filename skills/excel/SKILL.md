---
name: excel
description: Create, inspect, edit and verify Excel workbooks and CSV data with formulas, formatting and native charts using the standalone document-skills toolkit.
---
# Excel

Use the installed `document-skills` CLI or `document_skills.spreadsheet` Python API. Check `document-skills --help` and `doctor` first. Installation: `pip install quickerwrite-documents-skills` only when a published release exists; from this repository use `pip install .`.

Read [workflows](references/workflows.md) for the JSON schema and commands. Inspect existing sheets and selected ranges before edits; preserve the source and the user's formatting. Use native numeric values with number formats; keep identifiers as strings. Formulas must use explicit `{"formula":"=SUM(B2:B4)"}` objects. Never invent financial assumptions or disguise missing data as zero.

Verify totals against source data and at least one independently computed result. Inspection distinguishes formula text from cached results; openpyxl does not calculate. Recalculate a copy with LibreOffice when available, then inspect all relevant output cells for actual error types and missing caches. A populated cache can still be stale. Do not claim calculation verification if no calculation engine ran.

Use readable widths, wrapping, frozen headers and restrained number formats. Charts must reference the intended data. For printed output, fit width while allowing multiple pages vertically. Render and inspect each relevant page when layout matters; parser success does not prove visual quality.

Editing/recalculation can change unsupported features such as slicers, external links or newer Excel formulas. Reject macros and embedded executable objects. For complex existing workbooks, identify fidelity needs before conversion and validate in the target spreadsheet application.
