"""Bounded XLSX workflows. Strings are literal; formulas are always explicit."""

from io import BytesIO, StringIO
import csv
import math
from copy import copy
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import PatternFill
from openpyxl.utils.cell import range_boundaries
from .render import _guard_ooxml, _check_formula, _bounded_bytes

MAX_CELLS = 200_000


def bounds(reference):
    values = range_boundaries(reference)
    a, b, c, d = values
    if None in values or not (1 <= a <= c <= 16384 and 1 <= b <= d <= 1048576):
        raise ValueError("Expected a bounded A1 range")
    if (c - a + 1) * (d - b + 1) > MAX_CELLS:
        raise ValueError("Range exceeds 200000 cells")
    return values


def set_cell(cell, value):
    if isinstance(value, dict):
        if set(value) != {"formula"}:
            raise ValueError("Formula objects require exactly one formula field")
        _check_formula(value["formula"])
        cell.value = value["formula"]
    else:
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("Nonfinite numbers cannot be stored")
        if isinstance(value, str) and len(value) > 32767:
            raise ValueError("Cell text exceeds Excel limit")
        cell.value = value
        if isinstance(value, str):
            cell.data_type = "s"


def save(workbook):
    output = BytesIO()
    workbook.save(output)
    result = _bounded_bytes(output.getvalue())
    _guard_ooxml(result)
    return result


def create_xlsx(spec):
    """Create sheets with rows, styles, filters, freeze panes and native charts."""
    sheets = spec.get("sheets", [])
    if not 1 <= len(sheets) <= 100:
        raise ValueError("Expected 1 to 100 sheets")
    workbook = Workbook()
    workbook.remove(workbook.active)
    count = 0
    for sheet in sheets:
        name = sheet["name"]
        if name in workbook.sheetnames or not name or len(name) > 31:
            raise ValueError("Sheet names must be unique and at most 31 characters")
        ws = workbook.create_sheet(name)
        rows = sheet.get("rows", [])
        for r, row in enumerate(rows, 1):
            count += len(row)
            if count > MAX_CELLS or r > 1048576 or len(row) > 16384:
                raise ValueError("Workbook exceeds cell limit")
            for c, value in enumerate(row, 1):
                set_cell(ws.cell(r, c), value)
        ws.freeze_panes = sheet.get("freeze_panes")
        if sheet.get("auto_filter"):
            ws.auto_filter.ref = ws.dimensions
        for col, width in sheet.get("widths", {}).items():
            if not isinstance(width, (float, int)) or not 1 <= width <= 255:
                raise ValueError("Column width must be 1 to 255")
            bounds(col + "1")
            ws.column_dimensions[col].width = width
        for item in sheet.get("styles", []):
            apply_style(ws, item)
        for item in sheet.get("charts", []):
            chart_type = item.get("type", "bar")
            if chart_type not in {"bar", "line"}:
                raise ValueError("Charts support bar or line")
            chart = BarChart() if chart_type == "bar" else LineChart()
            a, b, c, d = bounds(item["data"])
            chart.add_data(
                Reference(ws, min_col=a, min_row=b, max_col=c, max_row=d),
                titles_from_data=True,
            )
            if item.get("categories"):
                a, b, c, d = bounds(item["categories"])
                chart.set_categories(
                    Reference(ws, min_col=a, min_row=b, max_col=c, max_row=d)
                )
            chart.title = item.get("title", "")
            bounds(item.get("anchor", "E2"))
            ws.add_chart(chart, item.get("anchor", "E2"))
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    return save(workbook)


def apply_style(ws, item):
    allowed = {"range", "bold", "font_color", "fill", "number_format", "wrap_text"}
    if set(item) - allowed:
        raise ValueError("Unsupported style property")
    a, b, c, d = bounds(item["range"])
    for row in ws.iter_rows(min_row=b, max_row=d, min_col=a, max_col=c):
        for cell in row:
            if "bold" in item or "font_color" in item:
                font = copy(cell.font)
                if "bold" in item:
                    font.bold = item["bold"]
                if "font_color" in item:
                    font.color = item["font_color"].lstrip("#")
                cell.font = font
            if "fill" in item:
                cell.fill = PatternFill("solid", fgColor=item["fill"].lstrip("#"))
            if "number_format" in item:
                cell.number_format = item["number_format"]
            if "wrap_text" in item:
                alignment = copy(cell.alignment)
                alignment.wrap_text = item["wrap_text"]
                cell.alignment = alignment


def inspect_xlsx(data, sheet=None, cell_range="A1:T50"):
    """Return formulas and cached values separately; caches may be stale."""
    _guard_ooxml(data)
    a, b, c, d = bounds(cell_range)
    wb = load_workbook(BytesIO(data), read_only=True, data_only=False)
    cached = load_workbook(BytesIO(data), read_only=True, data_only=True)
    try:
        result = {
            "sheets": [],
            "cache_notice": "Cached values may be absent or stale; inspection does not calculate formulas.",
        }
        for name in [sheet] if sheet else wb.sheetnames:
            rows = wb[name].iter_rows(min_row=b, max_row=d, min_col=a, max_col=c)
            values = cached[name].iter_rows(min_row=b, max_row=d, min_col=a, max_col=c)
            cells = []
            for row, cache in zip(rows, values):
                for cell, value in zip(row, cache):
                    if cell.value is None:
                        continue
                    cells.append(
                        {
                            "cell": cell.coordinate,
                            "value": cell.value,
                            "type": cell.data_type,
                            "cached": value.value if cell.data_type == "f" else None,
                            "error": value.data_type == "e" or cell.data_type == "e",
                        }
                    )
            result["sheets"].append(
                {
                    "name": name,
                    "rows": wb[name].max_row,
                    "columns": wb[name].max_column,
                    "cells": cells,
                }
            )
        return result
    finally:
        wb.close()
        cached.close()


def edit_xlsx(data, changes):
    """Edit explicit cells/styles. openpyxl may drop unsupported Excel features."""
    _guard_ooxml(data)
    if not isinstance(changes, list) or not 1 <= len(changes) <= 1000:
        raise ValueError("Expected 1 to 1000 changes")
    wb = load_workbook(BytesIO(data))
    for change in changes:
        ws = wb[change["sheet"]]
        if "cell" in change:
            a, b, c, d = bounds(change["cell"])
            if (a, b) != (c, d):
                raise ValueError("Expected one cell")
            set_cell(ws.cell(b, a), change["value"])
        elif "style" in change:
            apply_style(ws, change["style"])
        else:
            raise ValueError("Expected cell/value or style")
    return save(wb)


def csv_to_xlsx(text, sheet="Data"):
    if len(text.encode("utf-8")) > 20 * 1024 * 1024:
        raise ValueError("CSV exceeds 20 MiB")
    # CSV is data, never executable formulas; retain identifiers and leading zeros.
    return create_xlsx(
        {"sheets": [{"name": sheet, "rows": list(csv.reader(StringIO(text)))}]}
    )


def xlsx_to_csv(data, sheet, cell_range="A1:T50"):
    _guard_ooxml(data)
    a, b, c, d = bounds(cell_range)
    wb = load_workbook(BytesIO(data), read_only=True, data_only=True)
    output = StringIO(newline="")
    try:
        writer = csv.writer(output)
        for row in wb[sheet].iter_rows(
            min_row=b, max_row=d, min_col=a, max_col=c, values_only=True
        ):
            # Prefix dangerous textual values for spreadsheet consumers.
            writer.writerow(
                [
                    (
                        "'" + v
                        if isinstance(v, str)
                        and v.lstrip().startswith(("=", "+", "-", "@"))
                        else v
                    )
                    for v in row
                ]
            )
        return output.getvalue()
    finally:
        wb.close()
