# Excel commands and schema

Run from a checkout after `pip install .` (the package does not require the application server).

```sh
document-skills create budget.json budget.xlsx
document-skills inspect budget.xlsx --sheet Budget --range A1:C5
document-skills edit budget.xlsx revised.xlsx --changes changes.json
document-skills recalculate revised.xlsx calculated.xlsx
document-skills render calculated.xlsx calculated.pdf
document-skills preview calculated.pdf page.png --page 1
document-skills csv-import input.csv imported.xlsx
document-skills csv-export calculated.xlsx output.csv --sheet Budget --range A1:C5
```

Create input:
```json
{"sheets":[{"name":"Budget","rows":[["Item","Cost"],["Books",120],["Travel",80],["Total",{"formula":"=SUM(B2:B3)"}]],"freeze_panes":"A2","auto_filter":true,"widths":{"A":24,"B":18},"styles":[{"range":"A1:B1","bold":true,"fill":"17365D","font_color":"FFFFFF"},{"range":"B2:B4","number_format":"#,##0.00"}],"charts":[{"type":"bar","title":"Costs","data":"B1:B3","categories":"A2:A3","anchor":"D2"}]}]}
```

Edit input:
```json
[{"sheet":"Budget","cell":"B2","value":150},{"sheet":"Budget","style":{"range":"B2:B4","number_format":"#,##0.00"}}]
```

Styles: `bold`, `font_color`, `fill` (RGB), `number_format`, `wrap_text`. Charts: `bar` or `line`, data first row is series heading. Ranges must be bounded A1 ranges, not whole columns. Workbook creation supports up to 200,000 cells and 100 sheets; each inspected/style range is bounded to 200,000 cells.

CSV import preserves strings and leading zeros. CSV export uses cached values, prefixes potentially executable textual values with an apostrophe, and does not preserve formatting. Recalculate before exporting formulas; missing formula caches otherwise become empty fields. A1:T50 is the default export/inspection range; always select the complete intended range.

Errors are reported from Excel's cell type, not by searching ordinary text for an error-looking phrase. Inspection is range-limited: examine every relevant sheet/range for complete coverage. There is no built-in financial formula evaluator or guarantee of Excel/LibreOffice parity.
