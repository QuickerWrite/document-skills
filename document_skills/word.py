"""Word inspection and conservative run edits without rebuilding paragraphs."""

from io import BytesIO
from docx import Document
from .render import _guard_ooxml


def inspect_docx(data):
    _guard_ooxml(data)
    doc = Document(BytesIO(data))
    return {
        "paragraphs": [
            {
                "index": i,
                "text": p.text,
                "style": p.style.name,
                "runs": [
                    {"index": j, "text": r.text, "bold": r.bold, "italic": r.italic}
                    for j, r in enumerate(p.runs)
                ],
            }
            for i, p in enumerate(doc.paragraphs)
        ],
        "tables": [
            [[cell.text for cell in row.cells] for row in table.rows]
            for table in doc.tables
        ],
        "sections": len(doc.sections),
    }


def edit_docx(data, changes):
    """Replace named body runs. Complex fields, hyperlinks and revisions need XML tooling."""
    _guard_ooxml(data)
    if not isinstance(changes, list) or not 1 <= len(changes) <= 1000:
        raise ValueError("Expected 1 to 1000 run changes")
    doc = Document(BytesIO(data))
    for change in changes:
        p, r = change["paragraph"], change["run"]
        if type(p) is not int or not 0 <= p < len(doc.paragraphs):
            raise ValueError("Invalid paragraph index; inspect again")
        runs = doc.paragraphs[p].runs
        if type(r) is not int or not 0 <= r < len(runs):
            raise ValueError("Invalid run index; inspect again")
        if not isinstance(change["text"], str) or len(change["text"]) > 100000:
            raise ValueError("Invalid replacement text")
        # Do not destroy embedded fields, drawings, tabs or bookmarks inside a run.
        allowed = {"rPr", "t", "br", "cr", "tab"}
        if any(child.tag.rsplit("}", 1)[-1] not in allowed for child in runs[r]._r):
            raise ValueError("Complex run requires an OOXML-aware editor")
        runs[r].text = change["text"]
    out = BytesIO()
    doc.save(out)
    result = out.getvalue()
    _guard_ooxml(result)
    return result
