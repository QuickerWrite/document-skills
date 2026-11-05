"""Word 文档渲染。"""

import io
from typing import Any

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT

from .defaults import get_default_section
from .locale import infer_text_language, localized_text
from .render_utils import parse_hex_color, strip_markdown


def render_word(
    title: str,
    content: list[dict],
    styles: dict,
    options: dict[str, Any] | None = None,
) -> bytes:
    doc = Document()
    style_map = styles if isinstance(styles, dict) else {}
    render_options = options if isinstance(options, dict) else {}
    configured_defaults = render_options.get("document_defaults")
    doc_defaults = (
        configured_defaults
        if isinstance(configured_defaults, dict)
        else get_default_section("document")
    )
    sample_text = (
        title
        + "\n"
        + "\n".join(
            str(node.get("text") or (node.get("meta") or {}).get("text") or "")
            for node in content[:20]
            if isinstance(node, dict)
        )
    )
    document_language = infer_text_language(
        sample_text, render_options.get("language") or "zh-CN"
    )

    def _alignment(value: Any, default=WD_ALIGN_PARAGRAPH.LEFT):
        return {
            "left": WD_ALIGN_PARAGRAPH.LEFT,
            "center": WD_ALIGN_PARAGRAPH.CENTER,
            "right": WD_ALIGN_PARAGRAPH.RIGHT,
            "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
            "distributed": WD_ALIGN_PARAGRAPH.DISTRIBUTE,
        }.get(str(value or "").strip().lower(), default)

    section = doc.sections[0]
    page_sizes = {
        "a4": (8.2677, 11.6929),
        "letter": (8.5, 11.0),
        "b5": (6.9291, 9.8425),
    }
    page_width, page_height = page_sizes.get(
        str(render_options.get("page_size") or "a4").lower(), page_sizes["a4"]
    )
    if str(render_options.get("orientation") or "portrait").lower() == "landscape":
        section.orientation = WD_ORIENT.LANDSCAPE
        page_width, page_height = page_height, page_width
    section.page_width = Inches(page_width)
    section.page_height = Inches(page_height)
    margins = (
        render_options.get("margins_inches")
        if isinstance(render_options.get("margins_inches"), dict)
        else {}
    )
    for attr in ("top", "right", "bottom", "left"):
        try:
            value = max(0.25, min(float(margins.get(attr, 1.0)), 3.0))
        except (TypeError, ValueError):
            value = 1.0
        setattr(section, f"{attr}_margin", Inches(value))

    def _style_value(section: str, key: str, default: Any) -> Any:
        node = style_map.get(section)
        if isinstance(node, dict) and node.get(key) not in (None, ""):
            return node.get(key)
        return default

    heading_font_name = str(
        _style_value(
            "heading",
            "font_name",
            doc_defaults.get("document_default_font_name", "宋体"),
        )
    )
    heading_font_size = float(
        _style_value(
            "heading", "font_size", doc_defaults.get("document_heading_font_size", "16")
        )
    )
    heading_color_rgb = parse_hex_color(
        str(
            _style_value(
                "heading",
                "color",
                doc_defaults.get("document_default_text_color", "#000000"),
            )
        ),
        (0, 0, 0),
    )
    title_font_name = str(_style_value("title", "font_name", heading_font_name))
    title_font_size = float(
        _style_value("title", "font_size", max(heading_font_size + 8, 24))
    )
    title_color_rgb = parse_hex_color(
        str(_style_value("title", "color", "#%02x%02x%02x" % heading_color_rgb)),
        heading_color_rgb,
    )

    body_font_name = str(
        _style_value(
            "body", "font_name", doc_defaults.get("document_default_font_name", "宋体")
        )
    )
    body_font_size = float(
        _style_value(
            "body", "font_size", doc_defaults.get("document_body_font_size", "12")
        )
    )
    body_color_rgb = parse_hex_color(
        str(
            _style_value(
                "body",
                "color",
                doc_defaults.get("document_default_text_color", "#000000"),
            )
        ),
        (0, 0, 0),
    )
    body_options = (
        render_options.get("body")
        if isinstance(render_options.get("body"), dict)
        else {}
    )
    body_font_name = str(body_options.get("font_name") or body_font_name)
    body_font_size = float(body_options.get("font_size") or body_font_size)
    body_color_rgb = parse_hex_color(
        str(body_options.get("color") or "#%02x%02x%02x" % body_color_rgb),
        body_color_rgb,
    )

    def _apply_run_style(
        run,
        font_name: str,
        font_size: float,
        color_rgb: tuple[int, int, int],
        bold: bool,
        italic: bool,
        *,
        underline: bool = False,
        strike: bool = False,
        superscript: bool = False,
        subscript: bool = False,
        highlight: str = "",
        all_caps: bool = False,
        character_spacing_pt: Any = None,
    ) -> None:
        run.bold = bool(bold)
        run.italic = bool(italic)
        run.font.name = str(font_name)
        run.font.size = Pt(float(font_size))
        run.font.color.rgb = RGBColor(*color_rgb)
        run.underline = bool(underline)
        run.font.strike = bool(strike)
        run.font.superscript = bool(superscript)
        run.font.subscript = bool(subscript)
        run.font.all_caps = bool(all_caps)
        highlight_map = {
            "yellow": WD_COLOR_INDEX.YELLOW,
            "green": WD_COLOR_INDEX.BRIGHT_GREEN,
            "cyan": WD_COLOR_INDEX.TURQUOISE,
            "pink": WD_COLOR_INDEX.PINK,
            "blue": WD_COLOR_INDEX.BLUE,
            "red": WD_COLOR_INDEX.RED,
            "gray": WD_COLOR_INDEX.GRAY_25,
        }
        if str(highlight or "").lower() in highlight_map:
            run.font.highlight_color = highlight_map[str(highlight).lower()]
        if run._element.rPr is not None:
            run._element.rPr.rFonts.set(qn("w:eastAsia"), str(font_name))
            run._element.rPr.rFonts.set(qn("w:ascii"), str(font_name))
            run._element.rPr.rFonts.set(qn("w:hAnsi"), str(font_name))
            if character_spacing_pt not in (None, ""):
                try:
                    spacing = OxmlElement("w:spacing")
                    spacing.set(qn("w:val"), str(int(float(character_spacing_pt) * 20)))
                    run._element.rPr.append(spacing)
                except (TypeError, ValueError):
                    pass

    def _apply_paragraph_format(
        paragraph,
        *,
        body: bool = False,
        alignment: Any = None,
        formatting: dict[str, Any] | None = None,
    ) -> None:
        fmt = formatting if isinstance(formatting, dict) else {}
        selected_alignment = alignment
        if selected_alignment is None:
            selected_alignment = fmt.get("alignment") or (
                body_options.get("alignment") if body else "left"
            )
        paragraph.alignment = _alignment(
            selected_alignment,
            WD_ALIGN_PARAGRAPH.JUSTIFY if body else WD_ALIGN_PARAGRAPH.LEFT,
        )
        if body:
            try:
                paragraph.paragraph_format.line_spacing = max(
                    0.5,
                    min(
                        float(
                            fmt.get("line_spacing")
                            or render_options.get("line_spacing")
                            or 1.5
                        ),
                        5.0,
                    ),
                )
            except (TypeError, ValueError):
                paragraph.paragraph_format.line_spacing = 1.5
            try:
                paragraph.paragraph_format.space_after = Pt(
                    max(
                        0,
                        min(
                            float(
                                fmt.get(
                                    "space_after_pt",
                                    render_options.get("paragraph_spacing_after_pt")
                                    or 0,
                                )
                            ),
                            144,
                        ),
                    )
                )
            except (TypeError, ValueError):
                pass
            try:
                indent_chars = max(
                    0,
                    min(
                        float(
                            fmt.get(
                                "first_line_indent_chars",
                                render_options.get("first_line_indent_chars") or 0,
                            )
                        ),
                        20,
                    ),
                )
                paragraph.paragraph_format.first_line_indent = Pt(
                    body_font_size * indent_chars
                )
            except (TypeError, ValueError):
                pass
        if fmt.get("first_line_indent_inches") not in (None, ""):
            try:
                paragraph.paragraph_format.first_line_indent = Inches(
                    float(fmt["first_line_indent_inches"])
                )
            except (TypeError, ValueError):
                pass
        for key, attr in (
            ("left_indent_inches", "left_indent"),
            ("right_indent_inches", "right_indent"),
        ):
            if fmt.get(key) not in (None, ""):
                try:
                    setattr(paragraph.paragraph_format, attr, Inches(float(fmt[key])))
                except (TypeError, ValueError):
                    pass
        for key, attr in (
            ("space_before_pt", "space_before"),
            ("space_after_pt", "space_after"),
        ):
            if fmt.get(key) not in (None, ""):
                try:
                    setattr(paragraph.paragraph_format, attr, Pt(float(fmt[key])))
                except (TypeError, ValueError):
                    pass
        paragraph.paragraph_format.keep_with_next = bool(
            fmt.get("keep_with_next", False)
        )
        paragraph.paragraph_format.keep_together = bool(fmt.get("keep_together", False))
        paragraph.paragraph_format.page_break_before = bool(
            fmt.get("page_break_before", False)
        )

    def _add_text_run(
        paragraph, text: str, style: dict[str, Any], defaults: dict[str, Any]
    ) -> None:
        if not text:
            return
        style = dict(style)
        url = str(style.get("url") or "").strip()
        if url.startswith("https://") or url.startswith("http://"):
            style.setdefault("color", "#0563C1")
            style.setdefault("underline", True)
        run = paragraph.add_run(text)
        color = parse_hex_color(
            str(style.get("color") or defaults.get("color") or "#000000"),
            body_color_rgb,
        )
        _apply_run_style(
            run,
            str(style.get("font_name") or defaults.get("font_name") or body_font_name),
            float(
                style.get("font_size") or defaults.get("font_size") or body_font_size
            ),
            color,
            bool(style.get("bold", defaults.get("bold", False))),
            bool(style.get("italic", defaults.get("italic", False))),
            underline=bool(style.get("underline", defaults.get("underline", False))),
            strike=bool(style.get("strike", defaults.get("strike", False))),
            superscript=bool(style.get("superscript", False)),
            subscript=bool(style.get("subscript", False)),
            highlight=str(style.get("highlight") or ""),
            all_caps=bool(style.get("all_caps", False)),
            character_spacing_pt=style.get("character_spacing_pt"),
        )
        if url.startswith("https://") or url.startswith("http://"):
            relationship_id = paragraph.part.relate_to(
                url, RT.HYPERLINK, is_external=True
            )
            hyperlink = OxmlElement("w:hyperlink")
            hyperlink.set(qn("r:id"), relationship_id)
            paragraph._p.remove(run._r)
            hyperlink.append(run._r)
            paragraph._p.append(hyperlink)

    def _append_page_number(paragraph) -> None:
        run = paragraph.add_run()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        instruction = OxmlElement("w:instrText")
        instruction.set(qn("xml:space"), "preserve")
        instruction.text = " PAGE "
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        run._r.extend((begin, instruction, end))

    header_text = str(render_options.get("header_text") or "").strip()
    if header_text:
        header_options = (
            render_options.get("header")
            if isinstance(render_options.get("header"), dict)
            else {}
        )
        header_p = section.header.paragraphs[0]
        header_p.alignment = _alignment(header_options.get("alignment", "center"))
        _apply_run_style(
            header_p.add_run(header_text),
            str(header_options.get("font_name") or body_font_name),
            float(header_options.get("font_size") or max(body_font_size - 2, 8)),
            parse_hex_color(
                str(header_options.get("color") or "#000000"), body_color_rgb
            ),
            bool(header_options.get("bold", False)),
            bool(header_options.get("italic", False)),
            underline=bool(header_options.get("underline", False)),
        )
    footer_text = str(render_options.get("footer_text") or "").strip()
    if footer_text or bool(render_options.get("show_page_numbers", False)):
        footer_options = (
            render_options.get("footer")
            if isinstance(render_options.get("footer"), dict)
            else {}
        )
        footer_p = section.footer.paragraphs[0]
        footer_p.alignment = _alignment(footer_options.get("alignment", "center"))
        if footer_text:
            _apply_run_style(
                footer_p.add_run(footer_text),
                str(footer_options.get("font_name") or body_font_name),
                float(footer_options.get("font_size") or max(body_font_size - 2, 8)),
                parse_hex_color(
                    str(footer_options.get("color") or "#000000"), body_color_rgb
                ),
                bool(footer_options.get("bold", False)),
                bool(footer_options.get("italic", False)),
                underline=bool(footer_options.get("underline", False)),
            )
        if footer_text and bool(render_options.get("show_page_numbers", False)):
            footer_p.add_run(" · ")
        if bool(render_options.get("show_page_numbers", False)):
            _append_page_number(footer_p)

    title_options = (
        render_options.get("title")
        if isinstance(render_options.get("title"), dict)
        else {}
    )
    title_color_rgb = parse_hex_color(
        str(title_options.get("color") or "#%02x%02x%02x" % title_color_rgb),
        title_color_rgb,
    )
    if bool(render_options.get("include_title", False)) and title.strip():
        title_p = doc.add_paragraph()
        _apply_paragraph_format(
            title_p, alignment=title_options.get("alignment", "center")
        )
        title_p.paragraph_format.space_after = Pt(18)
        title_run = title_p.add_run(title.strip())
        _apply_run_style(
            title_run,
            str(title_options.get("font_name") or title_font_name),
            float(title_options.get("font_size") or title_font_size),
            title_color_rgb,
            bool(title_options.get("bold", True)),
            bool(title_options.get("italic", False)),
            underline=bool(title_options.get("underline", False)),
        )
        cover_lines = [
            str(render_options.get(key) or "").strip()
            for key in ("subtitle", "author", "organization", "date")
            if str(render_options.get(key) or "").strip()
        ]
        if cover_lines:
            for line in cover_lines:
                p = doc.add_paragraph()
                _apply_paragraph_format(p, alignment="center")
                _apply_run_style(
                    p.add_run(line),
                    body_font_name,
                    body_font_size,
                    body_color_rgb,
                    False,
                    False,
                )
            doc.add_page_break()

    if bool(render_options.get("include_toc", False)):
        toc_title = doc.add_paragraph()
        _apply_paragraph_format(toc_title, alignment="center")
        toc_label = localized_text(
            document_language,
            **{
                "zh-CN": "目录",
                "zh-TW": "目錄",
                "en-US": "Contents",
                "ja-JP": "目次",
                "ko-KR": "목차",
                "ru-RU": "Содержание",
                "fr-FR": "Sommaire",
                "de-DE": "Inhalt",
                "es-ES": "Índice",
            },
        )
        _apply_run_style(
            toc_title.add_run(toc_label),
            heading_font_name,
            heading_font_size,
            heading_color_rgb,
            True,
            False,
        )
        toc_p = doc.add_paragraph()
        field = OxmlElement("w:fldSimple")
        field.set(qn("w:instr"), 'TOC \\o "1-3" \\h \\z \\u')
        toc_p._p.append(field)
        doc.add_page_break()

    def _collect_text_segments(node: dict, fallback_text: str) -> list[dict[str, Any]]:
        raw_segments: Any = node.get("segments")
        if raw_segments is None:
            meta = node.get("meta")
            if isinstance(meta, dict):
                raw_segments = meta.get("segments")

        segments: list[dict[str, Any]] = []
        if isinstance(raw_segments, list):
            for item in raw_segments:
                if isinstance(item, dict):
                    segment_text = str(item.get("text") or "")
                    if segment_text == "":
                        continue
                    segment = dict(item)
                    segment["text"] = segment_text
                    segments.append(segment)
                elif isinstance(item, str) and item != "":
                    segments.append({"text": item, "bold": False, "italic": False})

        if segments:
            return segments

        cleaned = strip_markdown(fallback_text)
        if cleaned:
            return [{"text": cleaned, "bold": False, "italic": False}]
        return []

    for node in content:
        kind = node.get("kind", "paragraph")
        raw_text = str(node.get("text") or "")
        meta = node.get("meta") or {}

        if kind == "paragraph":
            segments = _collect_text_segments(node, raw_text)
            if not segments:
                continue

            level = meta.get("heading_level")
            is_title = str(level) == "1"
            font_name = str(
                node.get("font_name")
                or (
                    title_font_name
                    if is_title
                    else heading_font_name
                    if level
                    else body_font_name
                )
            )
            font_size = float(
                node.get("font_size")
                or (
                    title_font_size
                    if is_title
                    else heading_font_size
                    if level
                    else body_font_size
                )
            )
            color_rgb = parse_hex_color(
                node.get("color")
                or (
                    "#%02x%02x%02x"
                    % (
                        title_color_rgb
                        if is_title
                        else heading_color_rgb
                        if level
                        else body_color_rgb
                    )
                )
            )
            bold = node.get("bold", body_options.get("bold", bool(level)))
            italic = node.get("italic", body_options.get("italic", False))
            p = doc.add_paragraph()
            _apply_paragraph_format(
                p,
                body=not bool(level),
                alignment=node.get("alignment"),
                formatting=node,
            )
            defaults = {
                "font_name": font_name,
                "font_size": font_size,
                "color": "#%02x%02x%02x" % color_rgb,
                "bold": bool(bold),
                "italic": bool(italic),
                "underline": bool(
                    node.get("underline", body_options.get("underline", False))
                ),
                "strike": bool(node.get("strike", body_options.get("strike", False))),
            }
            for segment in segments:
                segment_text = str(segment.get("text") or "")
                if segment_text == "":
                    continue
                _add_text_run(p, segment_text, segment, defaults)

        elif kind == "heading":
            text = str(meta.get("text") or raw_text or "").strip()
            level = int(meta.get("level") or 1)
            level = max(1, min(level, 9))
            if not text:
                continue
            h = doc.add_heading(text, level=level)
            heading_options_all = (
                render_options.get("headings")
                if isinstance(render_options.get("headings"), dict)
                else {}
            )
            heading_options = (
                heading_options_all.get(str(level))
                if isinstance(heading_options_all.get(str(level)), dict)
                else {}
            )
            node_heading_options = {
                **heading_options,
                **{k: v for k, v in node.items() if v not in (None, "")},
            }
            _apply_paragraph_format(
                h,
                alignment=node_heading_options.get("alignment", "left"),
                formatting=node_heading_options,
            )
            h.paragraph_format.keep_with_next = True
            for run in h.runs:
                selected_font = str(
                    node_heading_options.get("font_name") or heading_font_name
                )
                selected_size = float(
                    node_heading_options.get("font_size") or heading_font_size
                )
                run.font.name = selected_font
                run.font.size = Pt(selected_size)
                default_heading_color = (
                    title_color_rgb if level == 1 else heading_color_rgb
                )
                run.font.color.rgb = RGBColor(
                    *parse_hex_color(
                        str(
                            node_heading_options.get("color")
                            or "#%02x%02x%02x" % default_heading_color
                        ),
                        default_heading_color,
                    )
                )
                run.bold = bool(node_heading_options.get("bold", True))
                run.italic = bool(node_heading_options.get("italic", False))
                run.underline = bool(node_heading_options.get("underline", False))
                if run._element.rPr is not None:
                    run._element.rPr.rFonts.set(qn("w:eastAsia"), selected_font)
                    run._element.rPr.rFonts.set(qn("w:ascii"), selected_font)
                    run._element.rPr.rFonts.set(qn("w:hAnsi"), selected_font)

        elif kind == "list":
            items = node.get("items") or meta.get("items") or []
            list_type = str(meta.get("list_type") or "bullet").lower()
            para_style = "List Number" if list_type == "ordered" else "List Bullet"
            if isinstance(items, list):
                for item in items:
                    item_config = item if isinstance(item, dict) else {}
                    item_text = strip_markdown(
                        str(item_config.get("text") if item_config else item)
                    )
                    if not item_text:
                        continue
                    p = doc.add_paragraph(item_text, style=para_style)
                    _apply_paragraph_format(
                        p,
                        body=True,
                        alignment=node.get("alignment"),
                        formatting={**node, **item_config},
                    )
                    p.paragraph_format.first_line_indent = None
                    if item_config.get("level") not in (None, ""):
                        try:
                            p.paragraph_format.left_indent = Inches(
                                0.25 * max(0, min(int(item_config["level"]), 8))
                            )
                        except (TypeError, ValueError):
                            pass
                    for run in p.runs:
                        _apply_run_style(
                            run,
                            str(
                                item_config.get("font_name")
                                or node.get("font_name")
                                or body_font_name
                            ),
                            float(
                                item_config.get("font_size")
                                or node.get("font_size")
                                or body_font_size
                            ),
                            parse_hex_color(
                                str(
                                    item_config.get("color")
                                    or node.get("color")
                                    or "#000000"
                                ),
                                body_color_rgb,
                            ),
                            bool(item_config.get("bold", node.get("bold", False))),
                            bool(item_config.get("italic", node.get("italic", False))),
                            underline=bool(
                                item_config.get(
                                    "underline", node.get("underline", False)
                                )
                            ),
                        )

        elif kind in {"quote", "code"}:
            text = str(node.get("text") or meta.get("text") or "")
            if not text:
                continue
            p = doc.add_paragraph()
            block_format = dict(node)
            block_format.setdefault("left_indent_inches", 0.35)
            block_format.setdefault("right_indent_inches", 0.15)
            block_format.setdefault("space_before_pt", 6)
            block_format.setdefault("space_after_pt", 6)
            _apply_paragraph_format(
                p,
                body=True,
                alignment=node.get("alignment", "left"),
                formatting=block_format,
            )
            p_pr = p._p.get_or_add_pPr()
            shading = OxmlElement("w:shd")
            shading.set(
                qn("w:fill"),
                str(
                    node.get("background_color")
                    or ("F3F4F6" if kind == "code" else "F8FAFC")
                ).lstrip("#"),
            )
            p_pr.append(shading)
            if kind == "quote":
                borders = OxmlElement("w:pBdr")
                left = OxmlElement("w:left")
                left.set(qn("w:val"), "single")
                left.set(qn("w:sz"), "16")
                left.set(
                    qn("w:color"), str(node.get("border_color") or "94A3B8").lstrip("#")
                )
                borders.append(left)
                p_pr.append(borders)
            _add_text_run(
                p,
                text,
                node,
                {
                    "font_name": "Courier New" if kind == "code" else body_font_name,
                    "font_size": node.get("font_size") or body_font_size,
                    "color": node.get("color") or "#334155",
                    "italic": kind == "quote",
                },
            )

        elif kind == "table":
            raw_rows = meta.get("rows", [])
            # 归一化行：LLM 有时输出 [{"cells":[...]}, ...] 而非 [[...]]
            rows_data: list[list] = []
            for row in raw_rows:
                if isinstance(row, list):
                    rows_data.append(row)
                elif isinstance(row, dict):
                    # 尝试常见键：cells、columns、values；回退到 dict 的值
                    cells = row.get("cells") or row.get("columns") or row.get("values")
                    if isinstance(cells, list):
                        rows_data.append(cells)
                    elif row:
                        rows_data.append(list(row.values()))
            if rows_data:
                caption = str(meta.get("caption") or "").strip()
                if caption:
                    caption_p = doc.add_paragraph()
                    _apply_paragraph_format(caption_p, alignment="center")
                    caption_options = (
                        render_options.get("caption")
                        if isinstance(render_options.get("caption"), dict)
                        else {}
                    )
                    _apply_run_style(
                        caption_p.add_run(caption),
                        str(caption_options.get("font_name") or body_font_name),
                        float(
                            caption_options.get("font_size")
                            or max(body_font_size - 1, 8)
                        ),
                        parse_hex_color(
                            str(caption_options.get("color") or "#000000"),
                            body_color_rgb,
                        ),
                        False,
                        False,
                    )

                def _cell_span(value: Any) -> int:
                    if not isinstance(value, dict):
                        return 1
                    try:
                        return max(1, min(int(value.get("colspan") or 1), 30))
                    except (TypeError, ValueError):
                        return 1

                cols = max(
                    (sum(_cell_span(cell) for cell in row) for row in rows_data),
                    default=1,
                )
                table = doc.add_table(rows=len(rows_data), cols=cols)
                table_options = (
                    render_options.get("table")
                    if isinstance(render_options.get("table"), dict)
                    else {}
                )
                table_alignment = str(
                    meta.get("alignment") or table_options.get("alignment") or "center"
                )
                table.alignment = {
                    "left": WD_TABLE_ALIGNMENT.LEFT,
                    "right": WD_TABLE_ALIGNMENT.RIGHT,
                }.get(table_alignment, WD_TABLE_ALIGNMENT.CENTER)
                table.autofit = bool(meta.get("autofit", True))
                table_style = str(
                    meta.get("style") or table_options.get("style") or "full_border"
                )
                if table_style == "full_border":
                    table.style = "Table Grid"
                column_widths = (
                    meta.get("column_widths_inches")
                    if isinstance(meta.get("column_widths_inches"), list)
                    else []
                )
                for r_idx, row in enumerate(rows_data):
                    if r_idx == 0 and bool(table_options.get("repeat_header", True)):
                        tr_pr = table.rows[r_idx]._tr.get_or_add_trPr()
                        repeat = OxmlElement("w:tblHeader")
                        repeat.set(qn("w:val"), "true")
                        tr_pr.append(repeat)
                    grid_index = 0
                    for cell_value in row:
                        if grid_index >= cols:
                            break
                        cell_config = cell_value if isinstance(cell_value, dict) else {}
                        span = min(_cell_span(cell_value), cols - grid_index)
                        cell = table.cell(r_idx, grid_index)
                        if span > 1:
                            cell = cell.merge(table.cell(r_idx, grid_index + span - 1))
                        vertical = str(
                            cell_config.get("vertical_alignment") or "center"
                        )
                        cell.vertical_alignment = {
                            "top": WD_CELL_VERTICAL_ALIGNMENT.TOP,
                            "bottom": WD_CELL_VERTICAL_ALIGNMENT.BOTTOM,
                        }.get(vertical, WD_CELL_VERTICAL_ALIGNMENT.CENTER)
                        if grid_index < len(column_widths):
                            try:
                                cell.width = Inches(float(column_widths[grid_index]))
                            except (TypeError, ValueError):
                                pass
                        background = (
                            str(cell_config.get("background_color") or "")
                            .strip()
                            .lstrip("#")
                        )
                        if background:
                            tc_pr = cell._tc.get_or_add_tcPr()
                            shading = tc_pr.find(qn("w:shd"))
                            if shading is None:
                                shading = OxmlElement("w:shd")
                                tc_pr.append(shading)
                            shading.set(qn("w:fill"), background)

                        paragraph = cell.paragraphs[0]
                        paragraph.clear()
                        cell_alignment = cell_config.get(
                            "alignment"
                        ) or table_options.get(
                            "header_alignment" if r_idx == 0 else "body_alignment",
                            "center",
                        )
                        _apply_paragraph_format(
                            paragraph, alignment=cell_alignment, formatting=cell_config
                        )
                        cell_defaults = {
                            "font_name": cell_config.get("font_name")
                            or table_options.get("font_name")
                            or body_font_name,
                            "font_size": cell_config.get("font_size")
                            or table_options.get("font_size")
                            or body_font_size,
                            "color": cell_config.get("color") or "#000000",
                            "bold": bool(cell_config.get("bold", r_idx == 0)),
                            "italic": bool(cell_config.get("italic", False)),
                        }
                        segments = (
                            cell_config.get("segments")
                            if isinstance(cell_config.get("segments"), list)
                            else None
                        )
                        if segments:
                            for segment in segments:
                                if isinstance(segment, dict):
                                    _add_text_run(
                                        paragraph,
                                        str(segment.get("text") or ""),
                                        segment,
                                        cell_defaults,
                                    )
                                elif isinstance(segment, str):
                                    _add_text_run(paragraph, segment, {}, cell_defaults)
                        else:
                            cell_text = (
                                cell_config.get("text") if cell_config else cell_value
                            )
                            _add_text_run(
                                paragraph,
                                strip_markdown(str(cell_text or "")),
                                cell_config,
                                cell_defaults,
                            )
                        if table_style in {"three_line", "no_vertical"}:
                            tc_pr = cell._tc.get_or_add_tcPr()
                            borders = tc_pr.first_child_found_in("w:tcBorders")
                            if borders is None:
                                borders = OxmlElement("w:tcBorders")
                                tc_pr.append(borders)
                            for edge in (
                                "top",
                                "left",
                                "bottom",
                                "right",
                                "insideH",
                                "insideV",
                            ):
                                existing = borders.find(qn(f"w:{edge}"))
                                if existing is not None:
                                    borders.remove(existing)
                                border = OxmlElement(f"w:{edge}")
                                show = False
                                size = "4"
                                if table_style == "three_line":
                                    show = (edge == "top" and r_idx == 0) or (
                                        edge == "bottom"
                                        and (r_idx == 0 or r_idx == len(rows_data) - 1)
                                    )
                                    size = (
                                        "12"
                                        if (
                                            edge == "top" or r_idx == len(rows_data) - 1
                                        )
                                        else "6"
                                    )
                                elif table_style == "no_vertical":
                                    show = (
                                        edge == "top" and r_idx == 0
                                    ) or edge == "bottom"
                                border.set(qn("w:val"), "single" if show else "nil")
                                if show:
                                    border.set(qn("w:sz"), size)
                                    border.set(qn("w:color"), "000000")
                                borders.append(border)
                        grid_index += span
                note = str(meta.get("note") or "").strip()
                if note:
                    note_p = doc.add_paragraph()
                    _apply_paragraph_format(note_p, alignment="left")
                    _apply_run_style(
                        note_p.add_run(note),
                        body_font_name,
                        max(body_font_size - 2, 8),
                        body_color_rgb,
                        False,
                        False,
                    )

        elif kind == "image":
            image_bytes = meta.get("image_bytes")
            if not isinstance(image_bytes, bytes) or not image_bytes:
                raise ValueError("Word image node is missing validated image bytes")
            try:
                requested_width = float(meta.get("width_inches") or 5.8)
            except (TypeError, ValueError):
                requested_width = 5.8
            section = doc.sections[-1]
            available_width = (
                float(section.page_width - section.left_margin - section.right_margin)
                / 914400
            )
            width_inches = max(0.5, min(requested_width, available_width))

            paragraph = doc.add_paragraph()
            _apply_paragraph_format(
                paragraph,
                alignment=meta.get("alignment", "center"),
                formatting={**node, "keep_with_next": bool(meta.get("caption"))},
            )
            run = paragraph.add_run()
            try:
                inline_shape = run.add_picture(
                    io.BytesIO(image_bytes), width=Inches(width_inches)
                )
                alt_text = str(meta.get("alt") or "").strip()
                if alt_text:
                    inline_shape._inline.docPr.set("descr", alt_text)
            except Exception as exc:
                raise ValueError("Word 图片无法解码或格式不受支持") from exc

            caption = str(meta.get("caption") or "").strip()
            if caption:
                caption_paragraph = doc.add_paragraph()
                _apply_paragraph_format(caption_paragraph, alignment="center")
                caption_run = caption_paragraph.add_run(caption)
                caption_options = (
                    render_options.get("caption")
                    if isinstance(render_options.get("caption"), dict)
                    else {}
                )
                _apply_run_style(
                    caption_run,
                    str(caption_options.get("font_name") or body_font_name),
                    float(
                        caption_options.get("font_size") or max(body_font_size - 1, 8)
                    ),
                    parse_hex_color(
                        str(caption_options.get("color") or "#000000"), body_color_rgb
                    ),
                    False,
                    False,
                )

        elif kind == "page_break":
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)

        elif kind == "horizontal_rule":
            paragraph = doc.add_paragraph()
            p_pr = paragraph._p.get_or_add_pPr()
            borders = OxmlElement("w:pBdr")
            bottom = OxmlElement("w:bottom")
            bottom.set(qn("w:val"), "single")
            bottom.set(qn("w:sz"), str(max(2, min(int(node.get("size") or 6), 48))))
            bottom.set(qn("w:color"), str(node.get("color") or "B8B8B8").lstrip("#"))
            borders.append(bottom)
            p_pr.append(borders)

        elif kind == "spacer":
            paragraph = doc.add_paragraph()
            try:
                paragraph.paragraph_format.space_after = Pt(
                    max(0, min(float(node.get("height_pt") or 12), 144))
                )
            except (TypeError, ValueError):
                paragraph.paragraph_format.space_after = Pt(12)

    if bool(render_options.get("include_toc", False)):
        settings_element = doc.settings._element
        update_fields = settings_element.find(qn("w:updateFields"))
        if update_fields is None:
            update_fields = OxmlElement("w:updateFields")
            settings_element.append(update_fields)
        update_fields.set(qn("w:val"), "true")

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
