from io import BytesIO
import json
from pathlib import Path
import shutil
import zipfile
import pytest
from docx import Document
from openpyxl import load_workbook
from document_skills.cli import main
from document_skills.spreadsheet import (
    create_xlsx,
    inspect_xlsx,
    edit_xlsx,
    csv_to_xlsx,
    xlsx_to_csv,
)
from document_skills.render import (
    _guard_ooxml,
    _check_formula,
    _libreoffice_pdf,
    recalculate_xlsx,
    _pdf_document,
    _page_png,
)
from document_skills.render_word import render_word
from document_skills.render_ppt import render_ppt
from document_skills.word import edit_docx, inspect_docx
from document_skills.pptx_patch import inspect_pptx, patch_pptx

ROOT = Path(__file__).resolve().parents[1]


def budget():
    return create_xlsx(json.loads((ROOT / "examples/budget.json").read_text()))


def test_excel_native_features_and_cache():
    data = budget()
    wb = load_workbook(BytesIO(data))
    ws = wb["Budget"]
    assert ws["B4"].value == "=SUM(B2:B3)"
    assert ws["A1"].font.bold and ws.freeze_panes == "A2"
    assert len(ws._charts) == 1
    report = inspect_xlsx(data, "Budget", "A1:B4")
    total = report["sheets"][0]["cells"][-1]
    assert total["type"] == "f" and total["cached"] is None
    modified = edit_xlsx(data, [{"sheet": "Budget", "cell": "B2", "value": 150}])
    assert load_workbook(BytesIO(modified))["Budget"]["B2"].value == 150
    assert len(load_workbook(BytesIO(modified))["Budget"]._charts) == 1


def test_literals_and_csv_injection():
    data = csv_to_xlsx("id,text\n001,=1+1\n002,@SUM(A1)\n")
    wb = load_workbook(BytesIO(data))
    assert wb.active["A2"].value == "001"
    assert wb.active["B2"].data_type == "s"
    result = xlsx_to_csv(data, "Data", "A1:B3")
    assert "'=1+1" in result and "'@SUM" in result


@pytest.mark.parametrize(
    "formula",
    [
        '=WEBSERVICE("https://example.com")',
        "=cmd|x!A1",
        "=SUM([x.xlsx]Sheet!A1)",
        '=HYPERLINK("file:///tmp/a")',
        '=REGISTER.ID("x")',
    ],
)
def test_reject_external_formula(formula):
    with pytest.raises(ValueError):
        _check_formula(formula)


@pytest.mark.parametrize("reference", ["A:A", "A1:XFD1048576", "XFE1", "A0"])
def test_range_limits(reference):
    with pytest.raises(ValueError):
        inspect_xlsx(budget(), cell_range=reference)


def test_zip_security():
    for name, value in [
        ("../escape", b"x"),
        ("xl/vbaProject.bin", b"x"),
        ("xl/embeddings/a.bin", b"x"),
        (
            "xl/workbook.xml",
            b'<!DOCTYPE x [<!ENTITY x SYSTEM "file:///etc/passwd">]><x>&x;</x>',
        ),
    ]:
        output = BytesIO()
        with zipfile.ZipFile(output, "w") as z:
            z.writestr("[Content_Types].xml", "<Types/>")
            z.writestr(name, value)
        with pytest.raises(Exception):
            _guard_ooxml(output.getvalue())


def test_named_formula_guard():
    output = BytesIO()
    with zipfile.ZipFile(output, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr(
            "xl/workbook.xml",
            '<workbook><definedName>WEBSERVICE("https://example.com")</definedName></workbook>',
        )
    with pytest.raises(ValueError):
        _guard_ooxml(output.getvalue())


def test_word_runs_preserved():
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("first").bold = True
    p.add_run("second").italic = True
    out = BytesIO()
    doc.save(out)
    edited = edit_docx(out.getvalue(), [{"paragraph": 0, "run": 0, "text": "changed"}])
    result = Document(BytesIO(edited)).paragraphs[0]
    assert result.text == "changedsecond"
    assert result.runs[0].bold and result.runs[1].italic
    assert inspect_docx(edited)["paragraphs"][0]["text"] == "changedsecond"


def test_ppt_patch_preserves_other_parts():
    data = render_ppt(
        "Overview",
        [{"kind": "slide", "text": "Title", "meta": {"bullets": ["Body"]}}],
        {},
    )
    report = inspect_pptx(data)
    element = next(x for x in report["slides"][0]["elements"] if x["text"])
    edited = patch_pptx(
        data,
        [
            {
                "slide_number": 1,
                "element_id": element["element_id"],
                "changes": {"text": "Changed"},
            }
        ],
    )
    with zipfile.ZipFile(BytesIO(data)) as a, zipfile.ZipFile(BytesIO(edited)) as b:
        differing = [n for n in a.namelist() if a.read(n) != b.read(n)]
    assert differing == ["ppt/slides/slide1.xml"]


def test_cli_atomic_non_overwrite(tmp_path):
    target = tmp_path / "budget.xlsx"
    assert main(["create", str(ROOT / "examples/budget.json"), str(target)]) == 0
    original = target.read_bytes()
    assert main(["create", str(ROOT / "examples/budget.json"), str(target)]) == 1
    assert target.read_bytes() == original
    assert main(["validate", str(target)]) == 0
    assert not list(tmp_path.glob(".documents-*"))


@pytest.mark.parametrize("kind", ["docx", "pptx", "xlsx"])
def test_examples(tmp_path, kind):
    source = {"docx": "report", "pptx": "deck", "xlsx": "budget"}[kind]
    target = tmp_path / f"example.{kind}"
    assert main(["create", str(ROOT / f"examples/{source}.json"), str(target)]) == 0
    assert main(["inspect", str(target)]) == 0


@pytest.mark.skipif(
    not (shutil.which("libreoffice") or shutil.which("soffice")),
    reason="LibreOffice unavailable",
)
def test_real_conversion_and_calculation():
    calculated = recalculate_xlsx(budget())
    wb = load_workbook(BytesIO(calculated), data_only=True)
    assert wb["Budget"]["B4"].value == 200
    for data, extension in [
        (calculated, ".xlsx"),
        (render_word("Report", [{"kind": "paragraph", "text": "Hello"}], {}), ".docx"),
        (render_ppt("Deck", [{"kind": "slide", "text": "Hello"}], {}), ".pptx"),
    ]:
        pdf = _libreoffice_pdf(data, extension)
        document = _pdf_document(pdf)
        try:
            assert len(document) >= 1
            assert _page_png(document, 1)[0].startswith(b"\x89PNG")
            with pytest.raises(ValueError):
                _page_png(document, 0)
        finally:
            document.close()


def test_fallback_slide_does_not_duplicate_or_indent_peer_facts():
    from pptx import Presentation

    facts = ["Books: 120", "Travel: 80", "Total: 200"]
    data = render_ppt(
        "Budget", [{"kind": "slide", "text": "Budget", "meta": {"bullets": facts}}], {}
    )
    slide = Presentation(BytesIO(data)).slides[0]
    texts = [
        p.text
        for shape in slide.shapes
        if shape.has_text_frame
        for p in shape.text_frame.paragraphs
    ]
    assert all(texts.count(fact) == 1 for fact in facts)
    paragraphs = [
        p
        for shape in slide.shapes
        if shape.has_text_frame
        for p in shape.text_frame.paragraphs
        if p.text in facts
    ]
    assert {p.level for p in paragraphs} == {0}


@pytest.mark.parametrize(
    "part",
    [
        "xl/macrosheets/sheet1.xml",
        "xl/connections.xml",
        "xl/queryTables/queryTable1.xml",
    ],
)
def test_legacy_macros_and_data_connections_rejected(part):
    out = BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        z.writestr("[Content_Types].xml", "<Types/>")
        z.writestr(part, "<root/>")
    with pytest.raises(ValueError):
        _guard_ooxml(out.getvalue())


def test_macro_enabled_content_type_rejected():
    out = BytesIO()
    with zipfile.ZipFile(out, "w") as z:
        z.writestr(
            "[Content_Types].xml",
            '<Types><Override ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/></Types>',
        )
    with pytest.raises(ValueError):
        _guard_ooxml(out.getvalue())
