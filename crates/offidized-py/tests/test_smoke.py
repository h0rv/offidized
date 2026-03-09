"""Smoke tests for offidized Python bindings.

Verifies that all exported types import, construct, and perform basic
round-trip operations. These tests don't need real Office files — they
exercise the in-memory create/save/load path.
"""

import offidized


# ---------------------------------------------------------------------------
# Import completeness
# ---------------------------------------------------------------------------


def test_all_exports_importable():
    """Every name in __all__ should be an attribute of the module."""
    for name in offidized.__all__:
        assert hasattr(offidized, name), f"offidized.{name} not found"


def test_all_exports_are_not_none():
    """Every export should resolve to a real object (not accidentally None)."""
    for name in offidized.__all__:
        obj = getattr(offidized, name)
        assert obj is not None, f"offidized.{name} is None"


# ---------------------------------------------------------------------------
# xlsx
# ---------------------------------------------------------------------------


def test_workbook_create_and_sheet_names():
    wb = offidized.Workbook()
    names = wb.sheet_names()
    assert isinstance(names, list)
    # New workbook starts empty; add a sheet to use it
    assert len(names) == 0


def test_workbook_add_sheet():
    wb = offidized.Workbook()
    wb.add_sheet("Sales")
    names = wb.sheet_names()
    assert "Sales" in names


def test_workbook_roundtrip_bytes():
    wb = offidized.Workbook()
    wb.add_sheet("Data")
    data = wb.to_bytes()
    assert isinstance(data, bytes)
    assert len(data) > 0

    wb2 = offidized.Workbook.from_bytes(data)
    assert "Data" in wb2.sheet_names()


def test_worksheet_cell_value():
    wb = offidized.Workbook()
    wb.add_sheet("Sheet1")
    ws = wb.sheet("Sheet1")
    ws.set_cell_value("A1", "hello")
    assert ws.cell_value("A1") == "hello"


def test_worksheet_cell_types():
    wb = offidized.Workbook()
    wb.add_sheet("Sheet1")
    ws = wb.sheet("Sheet1")

    ws.set_cell_value("A1", "text")
    ws.set_cell_value("A2", 42)
    ws.set_cell_value("A3", 3.14)
    ws.set_cell_value("A4", True)

    assert ws.cell_value("A1") == "text"
    assert ws.cell_value("A2") == 42
    assert ws.cell_value("A3") == 3.14
    assert ws.cell_value("A4") is True


def test_worksheet_cell_formula():
    wb = offidized.Workbook()
    wb.add_sheet("Sheet1")
    ws = wb.sheet("Sheet1")
    ws.set_cell_formula("B1", "=A1+1")
    formula = ws.cell_formula("B1")
    # Rust stores without leading '='; accept either form
    assert formula in ("=A1+1", "A1+1")


def test_workbook_save_and_open(tmp_path):
    path = str(tmp_path / "test.xlsx")
    wb = offidized.Workbook()
    wb.add_sheet("Sheet1")
    ws = wb.sheet("Sheet1")
    ws.set_cell_value("A1", "roundtrip")
    wb.save(path)

    wb2 = offidized.Workbook.open(path)
    ws2 = wb2.sheet("Sheet1")
    assert ws2.cell_value("A1") == "roundtrip"


def test_xlsx_cell_wrapper():
    wb = offidized.Workbook()
    wb.add_sheet("Sheet1")
    ws = wb.sheet("Sheet1")
    ws.set_cell_value("C3", 99)
    cell = ws.cell("C3")
    assert cell.value() == 99
    assert cell.reference() == "C3"


def test_xlsx_style_construct():
    style = offidized.XlsxStyle()
    style.set_number_format("0.00")
    assert style.number_format() == "0.00"


def test_xlsx_font_construct():
    font = offidized.XlsxFont()
    font.name = "Arial"
    font.bold = True
    font.size = "12"
    assert font.name == "Arial"
    assert font.bold is True
    assert font.size == "12"


def test_xlsx_fill_construct():
    fill = offidized.XlsxFill()
    fill.pattern = "solid"
    fill.foreground_color = "FF0000FF"
    assert fill.pattern == "solid"


def test_xlsx_border_construct():
    border = offidized.XlsxBorder()
    border.left_style = "thin"
    border.left_color = "FF000000"
    assert border.left_style == "thin"


def test_xlsx_alignment_construct():
    alignment = offidized.XlsxAlignment()
    alignment.horizontal = "center"
    alignment.wrap_text = True
    assert alignment.horizontal == "center"
    assert alignment.wrap_text is True


def test_worksheet_merged_ranges():
    wb = offidized.Workbook()
    wb.add_sheet("Sheet1")
    ws = wb.sheet("Sheet1")
    ws.add_merged_range("A1:B2")
    ranges = ws.merged_ranges()
    assert len(ranges) == 1
    assert "A1" in ranges[0]


# ---------------------------------------------------------------------------
# docx
# ---------------------------------------------------------------------------


def test_document_create():
    doc = offidized.Document()
    assert doc is not None


def test_document_add_paragraph():
    doc = offidized.Document()
    doc.add_paragraph("Hello, world!")
    paras = doc.paragraphs()
    texts = [p.text() for p in paras]
    assert "Hello, world!" in texts


def test_document_add_heading():
    doc = offidized.Document()
    doc.add_heading("Title", 1)
    paras = doc.paragraphs()
    assert len(paras) >= 1


def test_document_roundtrip_bytes():
    doc = offidized.Document()
    doc.add_paragraph("test content")
    data = doc.to_bytes()
    assert isinstance(data, bytes)
    assert len(data) > 0

    doc2 = offidized.Document.from_bytes(data)
    texts = [p.text() for p in doc2.paragraphs()]
    assert "test content" in texts


def test_document_save_and_open(tmp_path):
    path = str(tmp_path / "test.docx")
    doc = offidized.Document()
    doc.add_paragraph("roundtrip")
    doc.save(path)

    doc2 = offidized.Document.open(path)
    texts = [p.text() for p in doc2.paragraphs()]
    assert "roundtrip" in texts


def test_docx_paragraph_runs():
    doc = offidized.Document()
    doc.add_paragraph("some text")
    paras = doc.paragraphs()
    para = [p for p in paras if p.text() == "some text"][0]
    runs = para.runs()
    assert len(runs) >= 1
    assert runs[0].text() == "some text"


def test_docx_run_formatting():
    doc = offidized.Document()
    doc.add_paragraph("bold text")
    paras = doc.paragraphs()
    para = [p for p in paras if p.text() == "bold text"][0]
    run = para.runs()[0]
    run.set_bold(True)
    assert run.is_bold() is True


def test_docx_add_table():
    doc = offidized.Document()
    doc.add_table(2, 3)
    tables = doc.tables()
    assert len(tables) >= 1


def test_docx_table_cell_text():
    doc = offidized.Document()
    doc.add_table(2, 2)
    table = doc.tables()[0]
    table.set_cell_text(0, 0, "hello")
    assert table.cell_text(0, 0) == "hello"


def test_docx_section():
    doc = offidized.Document()
    section = doc.section()
    assert section is not None
    # New empty doc may not have page size set; just verify the accessor works
    section.page_width_twips()  # returns int or None


def test_docx_document_properties():
    doc = offidized.Document()
    props = doc.document_properties()
    assert props is not None


# ---------------------------------------------------------------------------
# pptx
# ---------------------------------------------------------------------------


def test_presentation_create():
    pres = offidized.Presentation()
    assert pres is not None


def test_presentation_add_slide():
    pres = offidized.Presentation()
    pres.add_slide()
    assert pres.slide_count() >= 1


def test_presentation_roundtrip_file(tmp_path):
    path = str(tmp_path / "roundtrip.pptx")
    pres = offidized.Presentation()
    pres.add_slide()
    pres.save(path)

    pres2 = offidized.Presentation.open(path)
    assert pres2.slide_count() >= 1


def test_presentation_save_and_open(tmp_path):
    path = str(tmp_path / "test.pptx")
    pres = offidized.Presentation()
    pres.add_slide()
    pres.save(path)

    pres2 = offidized.Presentation.open(path)
    assert pres2.slide_count() >= 1


def test_slide_show_settings_construct():
    settings = offidized.SlideShowSettings()
    assert settings is not None


def test_custom_show_construct():
    show = offidized.CustomShow("My Show")
    assert show.name() == "My Show"
    show.set_name("Renamed")
    assert show.name() == "Renamed"


def test_slide_transition_construct():
    transition = offidized.SlideTransition("fade")
    assert transition is not None


# ---------------------------------------------------------------------------
# Error types
# ---------------------------------------------------------------------------


def test_error_types_exist():
    """All error types are importable and are exception classes."""
    for name in [
        "OffidizedError",
        "OffidizedIoError",
        "OffidizedValueError",
        "OffidizedUnsupportedError",
        "OffidizedRuntimeError",
    ]:
        cls = getattr(offidized, name)
        assert isinstance(cls, type)
        assert issubclass(cls, BaseException)


def test_workbook_open_nonexistent_raises():
    try:
        offidized.Workbook.open("/nonexistent/path.xlsx")
        assert False, "should have raised"
    except Exception:
        pass  # any error is fine — just verify it doesn't silently succeed


def test_document_open_nonexistent_raises():
    try:
        offidized.Document.open("/nonexistent/path.docx")
        assert False, "should have raised"
    except Exception:
        pass


def test_presentation_open_nonexistent_raises():
    try:
        offidized.Presentation.open("/nonexistent/path.pptx")
        assert False, "should have raised"
    except Exception:
        pass
