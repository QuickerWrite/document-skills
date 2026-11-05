"""Portable command line interface; payloads are JSON files, never shell programs."""

import argparse
import json
import shutil
import sys
from pathlib import Path
from .io import read, write
from .render import _guard_ooxml, _libreoffice_pdf, _pdf_document, _page_png
from .spreadsheet import create_xlsx, inspect_xlsx, edit_xlsx, csv_to_xlsx, xlsx_to_csv
from .word import inspect_docx, edit_docx
from .pptx_patch import inspect_pptx, patch_pptx
from .render_word import render_word
from .render_ppt import render_ppt


def main(argv=None):
    parser = argparse.ArgumentParser(prog="document-skills")
    parser.add_argument("--version", action="version", version="0.2.0")
    sub = parser.add_subparsers(dest="command", required=True)
    sub.add_parser("doctor", help="Report optional conversion dependency")
    for cmd in (
        "create",
        "inspect",
        "edit",
        "validate",
        "render",
        "preview",
        "recalculate",
        "csv-import",
        "csv-export",
    ):
        p = sub.add_parser(cmd)
        p.add_argument("input", type=Path)
        if cmd in (
            "create",
            "edit",
            "render",
            "preview",
            "recalculate",
            "csv-import",
            "csv-export",
        ):
            p.add_argument("output", type=Path)
            p.add_argument("--overwrite", action="store_true")
        if cmd == "edit":
            p.add_argument("--changes", type=Path, required=True)
        if cmd in ("inspect", "csv-export"):
            p.add_argument("--sheet")
            p.add_argument("--range", default="A1:T50")
        if cmd == "preview":
            p.add_argument("--page", type=int, default=1)
            p.add_argument("--width", type=int, default=1000)
    args = parser.parse_args(argv)
    try:
        if args.command == "doctor":
            import os

            binary = (
                os.environ.get("LIBREOFFICE_BIN")
                or shutil.which("libreoffice")
                or shutil.which("soffice")
            )
            print(
                json.dumps(
                    {
                        "libreoffice": binary,
                        "python": sys.version.split()[0],
                        "native_creation": True,
                    }
                )
            )
            return 0
        data = read(args.input)
        ext = args.input.suffix.lower()
        if args.command == "create":
            spec = json.loads(data)
            kind = args.output.suffix.lower()
            if kind == ".xlsx":
                result = create_xlsx(spec)
            elif kind == ".docx":
                result = render_word(
                    spec.get("title", ""),
                    spec["content"],
                    spec.get("styles", {}),
                    spec.get("options", {}),
                )
            elif kind == ".pptx":
                result = render_ppt(
                    spec.get("title", ""),
                    spec["content"],
                    spec.get("styles", {}),
                    spec.get("options", {}),
                )
            else:
                raise ValueError("Output must be .docx, .pptx or .xlsx")
            _guard_ooxml(result)
        elif args.command == "inspect":
            if ext == ".xlsx":
                result = inspect_xlsx(data, args.sheet, args.range)
            elif ext == ".docx":
                result = inspect_docx(data)
            elif ext == ".pptx":
                result = inspect_pptx(data)
            else:
                raise ValueError("Inspection supports docx, pptx and xlsx")
            print(json.dumps(result, ensure_ascii=False, default=str))
            return 0
        elif args.command == "validate":
            _guard_ooxml(data)
            if ext == ".xlsx":
                inspect_xlsx(data)
            elif ext == ".docx":
                inspect_docx(data)
            elif ext == ".pptx":
                inspect_pptx(data)
            else:
                raise ValueError("Validation supports docx, pptx and xlsx")
            print(
                json.dumps(
                    {
                        "valid": True,
                        "scope": "package safety and parser readability; not full OOXML conformance or visual QA",
                    }
                )
            )
            return 0
        elif args.command == "edit":
            changes = json.loads(read(args.changes))
            editors = {".xlsx": edit_xlsx, ".docx": edit_docx, ".pptx": patch_pptx}
            if ext not in editors:
                raise ValueError("Unsupported edit format")
            result = editors[ext](data, changes)
        elif args.command == "render":
            result = _libreoffice_pdf(data, ext)
        elif args.command == "recalculate":
            from .render import recalculate_xlsx

            if ext != ".xlsx":
                raise ValueError("Recalculation requires .xlsx")
            result = recalculate_xlsx(data)
        elif args.command == "preview":
            pdf = data if ext == ".pdf" else _libreoffice_pdf(data, ext)
            document = _pdf_document(pdf)
            try:
                result = _page_png(document, args.page, args.width)[0]
            finally:
                document.close()
        elif args.command == "csv-import":
            result = csv_to_xlsx(data.decode("utf-8-sig"))
        elif args.command == "csv-export":
            if not args.sheet:
                raise ValueError("--sheet is required")
            result = xlsx_to_csv(data, args.sheet, args.range).encode("utf-8-sig")
        write(args.output, result, args.overwrite)
        return 0
    except Exception as exc:
        print(f"{type(exc).__name__}: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
