"""
Parametric parity test suite.

Every test runs against BOTH openpyxl and oxlsx with the same fixture.
A failure in oxlsx only → divergence from openpyxl.
A failure in both → bad fixture or broken sample file.
"""

import datetime
import os
import tempfile
import zipfile

import openpyxl
import pytest

import oxlsx


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def _openpyxl_load(path):
    return openpyxl.load_workbook(path)


def _oxlsx_load(path):
    return oxlsx.load_workbook(path)


@pytest.fixture(params=["openpyxl", "oxlsx"])
def load_wb(request):
    if request.param == "openpyxl":
        return _openpyxl_load
    return _oxlsx_load


@pytest.fixture(params=["openpyxl", "oxlsx"])
def new_wb(request):
    """Returns a (workbook, save_fn) tuple for the write path."""
    if request.param == "openpyxl":
        wb = openpyxl.Workbook()
        return wb, wb.save
    wb = oxlsx.Workbook()
    return wb, wb.save


# ---------------------------------------------------------------------------
# Read: workbook navigation
# ---------------------------------------------------------------------------


def test_load_workbook(load_wb):
    wb = load_wb("samples/test.xlsx")
    assert wb is not None


def test_sheetnames(load_wb):
    wb = load_wb("samples/test.xlsx")
    assert "Hoja1" in wb.sheetnames


def test_active_title(load_wb):
    wb = load_wb("samples/test.xlsx")
    assert wb.active.title == "Hoja1"


def test_getitem_found(load_wb):
    wb = load_wb("samples/test.xlsx")
    assert wb["Hoja1"].title == "Hoja1"


def test_getitem_missing(load_wb):
    wb = load_wb("samples/test.xlsx")
    with pytest.raises(KeyError):
        _ = wb["DoesNotExist"]


# ---------------------------------------------------------------------------
# Read: worksheet / cell access
# ---------------------------------------------------------------------------


def test_ws_getitem(load_wb):
    wb = load_wb("samples/test.xlsx")
    cell = wb.active["A1"]
    assert cell is not None


def test_ws_cell_method(load_wb):
    wb = load_wb("samples/test.xlsx")
    ws = wb.active
    assert ws.cell(row=1, column=1).value == ws["A1"].value


# ---------------------------------------------------------------------------
# Read: cell values
# ---------------------------------------------------------------------------


def test_cell_value_str(load_wb):
    wb = load_wb("samples/test.xlsx")
    v = wb.active["A1"].value
    assert isinstance(v, str)


def test_cell_value_int(load_wb):
    # C1 of test.xlsx → whole number 42
    wb = load_wb("samples/test.xlsx")
    v = wb.active["C1"].value
    assert isinstance(v, int) and not isinstance(v, bool)
    assert v == 42


def test_cell_value_date(load_wb):
    # B1 of multi_sheet.xlsx "Empleados" → date 2025-01-21
    # openpyxl returns datetime.datetime; oxlsx returns datetime.date.
    # Both are subclasses of datetime.date — so isinstance passes for both.
    # We compare only the date part to avoid divergence on time component.
    wb = load_wb("samples/multi_sheet.xlsx")
    v = wb["Empleados"]["B1"].value
    assert isinstance(v, datetime.date)
    assert v.year == 2025 and v.month == 1 and v.day == 21


def test_cell_value_none(load_wb):
    # A reference that holds no data → None
    wb = load_wb("samples/test.xlsx")
    cell = wb.active["Z99"]
    value = cell.value if cell is not None else None
    assert value is None


# ---------------------------------------------------------------------------
# Read: formatting
# ---------------------------------------------------------------------------


def test_cell_font_bold(load_wb):
    # A1 of test.xlsx has bold font
    wb = load_wb("samples/test.xlsx")
    font = wb.active["A1"].font
    assert font is not None
    assert isinstance(font.bold, bool)
    assert font.bold is True


def test_cell_fill_fg_color(load_wb):
    # A1 of test.xlsx has a solid fill
    wb = load_wb("samples/test.xlsx")
    fill = wb.active["A1"].fill
    assert fill is not None


# ---------------------------------------------------------------------------
# Write: round-trip and setitem
# ---------------------------------------------------------------------------


def test_write_round_trip(new_wb, load_wb):
    wb, save = new_wb
    ws = wb.active
    ws["A1"] = "hello"
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        save(path)
        wb2 = load_wb(path)
        assert wb2.active["A1"].value == "hello"
    finally:
        os.unlink(path)


def test_ws_setitem(new_wb):
    wb, _ = new_wb
    ws = wb.active
    ws["A1"] = "world"
    assert ws["A1"].value == "world"


def test_rmw_modify_cell(load_wb):
    wb = load_wb("samples/test.xlsx")
    wb.active["A1"] = "RMW"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        wb.save(path)
        wb2 = load_wb(path)
        assert wb2.active["A1"].value == "RMW"
    finally:
        os.unlink(path)


def test_rmw_preserves_other_sheets(load_wb):
    wb = load_wb("samples/multi_sheet.xlsx")
    wb["Empleados"]["A1"] = "RMW"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        wb.save(path)
        wb2 = load_wb(path)
        assert wb2["Productos"]["A1"].value == "Laptop"
    finally:
        os.unlink(path)


def test_rmw_preserves_font(load_wb):
    wb = load_wb("samples/test.xlsx")
    wb.active["C1"] = 99

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        wb.save(path)
        wb2 = load_wb(path)
        assert wb2.active["A1"].font.bold is True
    finally:
        os.unlink(path)


def test_write_formula(new_wb, load_wb):
    wb, save = new_wb
    ws = wb.active
    ws["A1"] = "=SUM(B1:B3)"

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        save(path)
        wb2 = load_wb(path)
        value = wb2.active["A1"].value
        assert isinstance(value, str)
        assert value == "=SUM(B1:B3)"
    finally:
        os.unlink(path)


def test_write_date_renders_as_date(new_wb, load_wb):
    wb, save = new_wb
    ws = wb.active
    ws["A1"] = datetime.date(2025, 1, 21)

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        save(path)
        wb2 = load_wb(path)
        value = wb2.active["A1"].value
        assert isinstance(value, datetime.date)
        assert value.year == 2025 and value.month == 1 and value.day == 21
    finally:
        os.unlink(path)


def test_write_merge_cells(new_wb):
    wb, save = new_wb
    ws = wb.active
    ws.merge_cells("A1:C1")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name
    try:
        save(path)
        with zipfile.ZipFile(path, "r") as zf:
            sheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert '<mergeCell ref="A1:C1"' in sheet_xml
    finally:
        os.unlink(path)


# ---------------------------------------------------------------------------
# Phase 9 — Iteration API
# ---------------------------------------------------------------------------


def test_max_row(load_wb):
    wb = load_wb("samples/test.xlsx")
    assert wb.active.max_row == 1


def test_max_column(load_wb):
    wb = load_wb("samples/test.xlsx")
    assert wb.active.max_column == 3


def test_min_row(load_wb):
    wb = load_wb("samples/test.xlsx")
    assert wb.active.min_row == 1


def test_min_column(load_wb):
    wb = load_wb("samples/test.xlsx")
    assert wb.active.min_column == 1


def test_iter_rows_values_only(load_wb):
    wb = load_wb("samples/test.xlsx")
    rows = list(wb.active.iter_rows(min_row=1, max_row=1, values_only=True))
    assert len(rows) == 1
    assert rows[0][0] == "Hola"
    assert rows[0][2] == 42


def test_iter_rows_all_values_only(load_wb):
    # No bounds → all data rows
    wb = load_wb("samples/multi_sheet.xlsx")
    rows = list(wb["Empleados"].iter_rows(values_only=True))
    assert len(rows) == 1
    assert rows[0][0] == "Juan"


def test_iter_rows_returns_cells(load_wb):
    wb = load_wb("samples/test.xlsx")
    rows = list(wb.active.iter_rows(min_row=1, max_row=1, values_only=False))
    assert len(rows) == 1
    # each element is a cell-like object with .value
    assert rows[0][0].value == "Hola"


def test_iter_cols_values_only(load_wb):
    wb = load_wb("samples/test.xlsx")
    cols = list(wb.active.iter_cols(min_col=1, max_col=1, values_only=True))
    assert len(cols) == 1
    assert cols[0][0] == "Hola"


def test_iter_cols_all_values_only(load_wb):
    wb = load_wb("samples/multi_sheet.xlsx")
    cols = list(wb["Empleados"].iter_cols(values_only=True))
    assert len(cols) == 2
    assert cols[0][0] == "Juan"


def test_ws_values(load_wb):
    # ws.values → generator of row tuples (values only)
    wb = load_wb("samples/test.xlsx")
    rows = list(wb.active.values)
    assert len(rows) == 1
    assert rows[0][0] == "Hola"
    assert rows[0][2] == 42


def test_ws_rows_property(load_wb):
    # ws.rows → generator of cell tuples
    wb = load_wb("samples/test.xlsx")
    rows = list(wb.active.rows)
    assert len(rows) == 1
    assert rows[0][0].value == "Hola"


def test_number_format_round_trip(load_wb, new_wb):
    wb, save = new_wb
    ws = wb.active
    ws["A1"] = 42

    if type(wb).__module__.startswith("openpyxl"):
        ws["A1"].number_format = "#,##0.00"
    else:
        ws.set_cell_number_format("A1", "#,##0.00")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        save(path)
        wb2 = load_wb(path)
        assert wb2.active["A1"].number_format == "#,##0.00"
    finally:
        os.unlink(path)


def test_alignment_round_trip(load_wb, new_wb):
    wb, save = new_wb
    ws = wb.active
    ws["A1"] = "hola"

    if type(wb).__module__.startswith("openpyxl"):
        ws["A1"].alignment = openpyxl.styles.Alignment(
            horizontal="center", vertical="top", wrap_text=True, indent=2
        )
    else:
        ws.set_cell_alignment(
            "A1",
            oxlsx.Alignment(horizontal="center", vertical="top", wrap_text=True, indent=2),
        )

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        save(path)
        wb2 = load_wb(path)
        alignment = wb2.active["A1"].alignment
        assert alignment.horizontal == "center"
        assert alignment.vertical == "top"
        assert alignment.wrap_text is True
        assert alignment.indent == 2
    finally:
        os.unlink(path)


def test_border_round_trip(load_wb, new_wb):
    wb, save = new_wb
    ws = wb.active
    ws["A1"] = "hola"

    if type(wb).__module__.startswith("openpyxl"):
        ws["A1"].border = openpyxl.styles.Border(
            left=openpyxl.styles.Side(style="thin", color=openpyxl.styles.Color(theme=3)),
            right=openpyxl.styles.Side(style="double", color=openpyxl.styles.Color(indexed=6)),
        )
    else:
        ws.set_cell_border(
            "A1",
            oxlsx.Border(
                left=oxlsx.Side(style="thin", color=oxlsx.Color(theme=3)),
                right=oxlsx.Side(style="double", color=oxlsx.Color(indexed=6)),
            ),
        )

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        save(path)
        wb2 = load_wb(path)
        border = wb2.active["A1"].border
        assert border.left.style == "thin"
        assert border.right.style == "double"
        assert border.left.color.theme == 3
        assert border.right.color.indexed == 6
    finally:
        os.unlink(path)


def test_protection_round_trip(load_wb, new_wb):
    wb, save = new_wb
    ws = wb.active
    ws["A1"] = "hola"

    if type(wb).__module__.startswith("openpyxl"):
        ws["A1"].protection = openpyxl.styles.Protection(locked=False, hidden=True)
    else:
        ws.set_cell_protection("A1", oxlsx.Protection(locked=False, hidden=True))

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        save(path)
        wb2 = load_wb(path)
        protection = wb2.active["A1"].protection
        assert protection.locked is False
        assert protection.hidden is True
    finally:
        os.unlink(path)


def test_style_round_trip_preserves_existing_ids(load_wb):
    wb = load_wb("samples/test.xlsx")
    original_style = wb.active["A1"]._style if type(wb).__module__.startswith("openpyxl") else None

    if type(wb).__module__.startswith("openpyxl"):
        wb.active["C1"].number_format = "0.000"
    else:
        wb.active.set_cell_number_format("C1", "0.000")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as f:
        path = f.name

    try:
        wb.save(path)
        wb2 = load_wb(path)
        if type(wb2).__module__.startswith("openpyxl"):
            assert wb2.active["A1"]._style == original_style
        else:
            assert wb2.active["A1"].font.bold is True
            assert wb2.active["A1"].fill.fg_color.indexed == 2
    finally:
        os.unlink(path)
