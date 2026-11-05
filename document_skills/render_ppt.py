"""支持基本主题的 PPT 渲染。"""

import base64
import io
import logging
from typing import Any, cast

from pptx import Presentation
from pptx.util import Inches, Pt as PptPt
from pptx.dml.color import RGBColor as PptRGB
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.enum.text import PP_ALIGN
from pptx.parts.image import Image as PptxImage

from .defaults import get_default_section
from .render_utils import parse_hex_color, strip_markdown

logger = logging.getLogger(__name__)


def _decode_image_value(raw: object) -> bytes | None:
    """解码可能是原始 bytes、bytearray 或 base64 字符串的图片值。

    支持纯 base64 以及 data URI 前缀，例如::

        data:image/png;base64,<data>
    """
    if isinstance(raw, (bytes, bytearray)):
        return bytes(raw)
    if isinstance(raw, str):
        s = raw.strip()
        # 存在 data URI 前缀时将其去除
        if "base64," in s:
            s = s.split("base64,", 1)[1].strip()
        if not s:
            return None
        try:
            return base64.b64decode(s)
        except Exception:
            logger.debug("invalid base64 image payload", exc_info=True)
            return None
    return None


def _resolve_ppt_ratio(ppt_options: dict | None) -> str:
    options = ppt_options if isinstance(ppt_options, dict) else {}
    value = str(options.get("aspect_ratio") or options.get("ratio") or "").strip()
    if value in {"4:3", "4x3"}:
        return "4:3"
    return "16:9"


def _resolve_background_color(ppt_options: dict | None) -> tuple[int, int, int]:
    options = ppt_options if isinstance(ppt_options, dict) else {}
    raw_theme = options.get("theme")
    theme = raw_theme if isinstance(raw_theme, dict) else {}
    value = theme.get("background_color") or options.get("background_color")
    return parse_hex_color(value, default=(244, 241, 236))


def _apply_slide_background(slide, color_rgb: tuple[int, int, int] | None) -> None:
    if color_rgb is None:
        return  # 保留源版式/母版继承的背景。
    try:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = PptRGB(*color_rgb)
    except Exception:
        logger.debug("slide background color could not be applied", exc_info=True)


def _apply_slide_background_image(
    slide, image_bytes: bytes, slide_width: int, slide_height: int
) -> None:
    try:
        pic = _add_picture_keep_ratio(
            slide,
            image_bytes=image_bytes,
            left=0,
            top=0,
            width=slide_width,
            height=slide_height,
            mode="cover",
        )
        if pic is None:
            return
        sp_tree = slide.shapes._spTree
        sp_tree.remove(pic._element)
        sp_tree.insert(2, pic._element)
    except Exception:
        logger.debug("slide background image could not be applied", exc_info=True)


def _get_image_size_px(image_bytes: bytes) -> tuple[int, int] | None:
    try:
        parsed = PptxImage.from_blob(image_bytes)
        size = getattr(parsed, "size", None)
        if not isinstance(size, tuple) or len(size) != 2:
            return None
        width_px = int(size[0])
        height_px = int(size[1])
        if width_px <= 0 or height_px <= 0:
            return None
        return width_px, height_px
    except Exception:
        return None


def _add_picture_keep_ratio(
    slide,
    image_bytes: bytes,
    left: int,
    top: int,
    width: int,
    height: int,
    mode: str = "contain",
):
    try:
        size_px = _get_image_size_px(image_bytes)
        if size_px is None:
            return slide.shapes.add_picture(
                io.BytesIO(image_bytes), left, top, width=width, height=height
            )

        image_width_px, image_height_px = size_px
        if image_width_px <= 0 or image_height_px <= 0:
            return slide.shapes.add_picture(
                io.BytesIO(image_bytes), left, top, width=width, height=height
            )

        width_ratio = float(width) / float(image_width_px)
        height_ratio = float(height) / float(image_height_px)
        if mode == "cover":
            scale = max(width_ratio, height_ratio)
        else:
            scale = min(width_ratio, height_ratio)

        draw_width = max(1, int(image_width_px * scale))
        draw_height = max(1, int(image_height_px * scale))
        draw_left = int(left + (width - draw_width) / 2)
        draw_top = int(top + (height - draw_height) / 2)
        return slide.shapes.add_picture(
            io.BytesIO(image_bytes),
            draw_left,
            draw_top,
            width=draw_width,
            height=draw_height,
        )
    except Exception:
        return None


def _apply_slide_transition(slide, transition: str = "fade") -> None:
    try:
        sld = slide._element
        if sld.find(qn("p:transition")) is not None:
            return
        valid_transitions = {
            "fade",
            "push",
            "wipe",
            "split",
            "cover",
            "uncover",
            "dissolve",
        }
        trans = transition.strip().lower() if transition else "fade"
        if trans not in valid_transitions:
            trans = "fade"
        transition_xml = parse_xml(
            f"<p:transition {nsdecls('p')}><p:{trans}/></p:transition>"
        )
        sld.insert(1, transition_xml)
    except Exception:
        logger.debug("could not apply PPT slide transition", exc_info=True)


def _apply_element_animation(slide, shape, animation: str | None) -> None:
    """通过 Open XML 为形状应用进入动画。"""
    if not animation:
        return
    anim = str(animation).strip().lower()
    anim_map = {
        "fadein": "appear",
        "flyinleft": "fly",
        "flyinright": "fly",
        "flyinbottom": "fly",
        "zoomin": "appear",
        "wiperight": "wipe",
    }
    if anim not in anim_map:
        return
    try:
        sp_id = shape.shape_id
        if anim in {"flyinleft", "flyinright", "flyinbottom"}:
            dir_map = {"flyinleft": "l", "flyinright": "r", "flyinbottom": "b"}
            d = dir_map.get(anim, "l")
            anim_xml = (
                f'<p:anim {nsdecls("p", "a")} calcmode="lin" valueType="num">'
                f'<p:cBhvr><p:cTn id="1" dur="500" fill="hold"/>'
                f'<p:tgtEl><p:spTgt spid="{sp_id}"/></p:tgtEl>'
                f"<p:attrNameLst><p:attrName>ppt_x</p:attrName></p:attrNameLst>"
                f"</p:cBhvr></p:anim>"
            )
            # 简化处理：目前只设置过渡，完整动画需要时序树
            _ = anim_xml, d  # 预留给未来可能用途
    except Exception:
        logger.debug("could not apply PPT element animation", exc_info=True)


def _apply_shape_fill(shape, bg_color: str | None, border_color: str | None) -> None:
    """为形状应用背景填充和边框。"""
    if bg_color:
        color = parse_hex_color(bg_color, default=None)
        if color is not None:
            try:
                fill = shape.fill
                fill.solid()
                fill.fore_color.rgb = PptRGB(*color)
            except Exception:
                logger.debug("could not apply PPT shape fill", exc_info=True)
    if border_color:
        color = parse_hex_color(border_color, default=None)
        if color is not None:
            try:
                line = shape.line
                line.color.rgb = PptRGB(*color)
                line.width = PptPt(1.5)
            except Exception:
                logger.debug("could not apply PPT shape border", exc_info=True)


def _as_text_shape(shape) -> Any | None:
    if shape is None:
        return None
    if getattr(shape, "has_text_frame", False):
        return cast(Any, shape)
    return None


def _safe_float(value: object, default: float) -> float:
    try:
        return float(cast(Any, value))
    except (TypeError, ValueError):
        return default


def _percent_to_emu(value: object, total: int, default_percent: float) -> int:
    pct = _safe_float(value, default_percent)
    if 0.0 <= pct <= 1.0:
        pct *= 100.0
    pct = max(0.0, min(100.0, pct))
    return int((pct / 100.0) * total)


def _resolve_align(value: object) -> PP_ALIGN | None:
    if value is None:
        return None
    text = str(value).strip().lower()
    if text == "left":
        return PP_ALIGN.LEFT
    if text in {"center", "middle"}:
        return PP_ALIGN.CENTER
    if text == "right":
        return PP_ALIGN.RIGHT
    if text == "justify":
        return PP_ALIGN.JUSTIFY
    return None


def _extract_layout_elements(meta: dict | None) -> list[dict]:
    if not isinstance(meta, dict):
        return []
    layout = meta.get("layout")
    if isinstance(layout, dict):
        elements = layout.get("elements")
        if isinstance(elements, list):
            return [item for item in elements if isinstance(item, dict)]
    raw_elements = meta.get("elements")
    if isinstance(raw_elements, list):
        return [item for item in raw_elements if isinstance(item, dict)]
    return []


def _ensure_foreground_image_slot(
    elements: list[dict], *, slide_index: int, enabled: bool
) -> list[dict]:
    """保留模型坐标；仅在缺失时添加一个多样化的安全槽位。"""
    result = [dict(item) for item in elements]
    if not enabled or any(
        str(item.get("kind") or item.get("type") or "").strip().lower() == "image"
        for item in result
    ):
        return result
    fallback_slots = (
        {"x": 63, "y": 19, "w": 31, "h": 67},
        {"x": 6, "y": 23, "w": 34, "h": 61},
        {"x": 53, "y": 51, "w": 41, "h": 38},
        {"x": 6, "y": 52, "w": 42, "h": 37},
    )
    result.append(
        {"kind": "image", **fallback_slots[slide_index % len(fallback_slots)]}
    )
    return result


def _render_layout_elements(
    slide,
    elements: list[dict],
    slide_width: int,
    slide_height: int,
    default_font_name: str,
    default_font_size: float,
    default_color_rgb: tuple[int, int, int],
    image_bytes: bytes | None = None,
    has_background_image: bool = False,
    slide_bg_rgb: tuple[int, int, int] | None = None,
) -> bool:
    rendered = False
    image_used = False

    def _readable_color(c: tuple[int, int, int]) -> tuple[int, int, int]:
        """强制文本可读：照片背景上的深色文本始终使用白色；
        纯色幻灯片背景则应用对比度安全处理。"""
        # 叠加在卡片上的文本，其背景是最近绘制的先前不透明形状，
        # 而不是幻灯片画布。
        from pptx.enum.dml import MSO_FILL, MSO_COLOR_TYPE

        for shape in reversed(list(slide.shapes)):
            if (
                shape.left <= left
                and shape.top <= top
                and shape.left + shape.width >= left + width
                and shape.top + shape.height >= top + height
                and shape.fill.type == MSO_FILL.SOLID
                and shape.fill.fore_color.type == MSO_COLOR_TYPE.RGB
            ):
                return _contrast_safe(c, tuple(shape.fill.fore_color.rgb))
        if has_background_image:
            # 不可预测的照片 → 对深色/中间色强制使用白色
            lum = (0.2126 * c[0] + 0.7152 * c[1] + 0.0722 * c[2]) / 255
            return (255, 255, 255) if lum < 0.7 else c
        if slide_bg_rgb is not None:
            return _contrast_safe(c, slide_bg_rgb)
        return c

    def _contrast_safe(
        text_rgb: tuple[int, int, int], bg_rgb: tuple[int, int, int]
    ) -> tuple[int, int, int]:
        """若 text_rgb 与 bg_rgb 的对比度≥3:1 则返回 text_rgb，否则选择黑或白。"""

        def _lum(r: int, g: int, b: int) -> float:
            def _chan(c: int) -> float:
                s = c / 255.0
                return s / 12.92 if s <= 0.03928 else ((s + 0.055) / 1.055) ** 2.4

            return 0.2126 * _chan(r) + 0.7152 * _chan(g) + 0.0722 * _chan(b)

        def _cr(c1: tuple[int, int, int], c2: tuple[int, int, int]) -> float:
            l1, l2 = _lum(*c1), _lum(*c2)
            bright, dark = max(l1, l2), min(l1, l2)
            return (bright + 0.05) / (dark + 0.05)

        if _cr(text_rgb, bg_rgb) >= 3.0:
            return text_rgb
        white_cr = _cr((255, 255, 255), bg_rgb)
        black_cr = _cr((0, 0, 0), bg_rgb)
        return (255, 255, 255) if white_cr >= black_cr else (0, 0, 0)

    for element in elements:
        kind = str(element.get("type") or element.get("kind") or "text").strip().lower()
        left = _percent_to_emu(element.get("x"), slide_width, 8)
        top = _percent_to_emu(element.get("y"), slide_height, 18)
        width = _percent_to_emu(element.get("w"), slide_width, 84)
        height = _percent_to_emu(element.get("h"), slide_height, 20)
        min_width = max(int(slide_width * 0.08), 1)
        min_height = max(int(slide_height * 0.06), 1)
        width = max(width, min_width)
        height = max(height, min_height)
        animation = str(element.get("animation") or "").strip() or None

        if kind == "divider":
            try:
                divider_color = parse_hex_color(
                    element.get("color"), default=(200, 200, 200)
                )
                shape = slide.shapes.add_shape(
                    1,  # MSO_SHAPE.RECTANGLE
                    left,
                    top,
                    width,
                    max(int(slide_height * 0.003), 1),
                )
                shape.fill.solid()
                shape.fill.fore_color.rgb = PptRGB(*divider_color)
                shape.line.fill.background()
                rendered = True
            except Exception:
                logger.debug("could not render PPT divider", exc_info=True)
            continue

        if kind == "image":
            if isinstance(image_bytes, (bytes, bytearray)) and not image_used:
                try:
                    pic = _add_picture_keep_ratio(
                        slide,
                        image_bytes=bytes(image_bytes),
                        left=left,
                        top=top,
                        width=width,
                        height=height,
                        mode="contain",
                    )
                    if pic is not None:
                        if animation:
                            _apply_element_animation(slide, pic, animation)
                        rendered = True
                        image_used = True
                        continue
                except Exception:
                    logger.debug("could not render PPT image element", exc_info=True)
            # 无可用图片（或已使用）——跳过，而不是显示 "[图片]" 文本
            continue

        # --- 表格元素 ---
        if kind == "table":
            try:
                headers_raw = element.get("headers")
                rows_raw = element.get("rows")
                headers = (
                    [str(h) for h in headers_raw]
                    if isinstance(headers_raw, list)
                    else []
                )
                rows = []
                if isinstance(rows_raw, list):
                    for row in rows_raw:
                        if isinstance(row, list):
                            rows.append([str(c) for c in row])
                if not headers and not rows:
                    continue
                num_cols = max(len(headers), max((len(r) for r in rows), default=0))
                num_rows = (1 if headers else 0) + len(rows)
                if num_cols < 1 or num_rows < 1:
                    continue
                tbl_shape = slide.shapes.add_table(
                    num_rows, num_cols, left, top, width, height
                )
                table = tbl_shape.table
                tbl_font_size = _safe_float(element.get("font_size"), default_font_size)
                tbl_color = _readable_color(
                    parse_hex_color(element.get("color"), default=default_color_rgb)
                )
                header_color = parse_hex_color(
                    element.get("header_color"), default=(255, 255, 255)
                )
                header_bg = parse_hex_color(
                    element.get("header_bg_color"), default=(26, 35, 126)
                )
                row_offset = 0
                if headers:
                    for ci in range(num_cols):
                        cell = table.cell(0, ci)
                        cell.text = headers[ci] if ci < len(headers) else ""
                        for p in cell.text_frame.paragraphs:
                            p.font.name = str(
                                element.get("font_name") or default_font_name
                            )
                            p.font.size = PptPt(max(10, tbl_font_size))
                            p.font.color.rgb = PptRGB(*header_color)
                            p.font.bold = True
                        try:
                            cell_fill = cell.fill
                            cell_fill.solid()
                            cell_fill.fore_color.rgb = PptRGB(*header_bg)
                        except Exception:
                            logger.debug(
                                "could not style PPT table header", exc_info=True
                            )
                    row_offset = 1
                for ri, row in enumerate(rows):
                    for ci in range(num_cols):
                        cell = table.cell(ri + row_offset, ci)
                        cell.text = row[ci] if ci < len(row) else ""
                        for p in cell.text_frame.paragraphs:
                            p.font.name = str(
                                element.get("font_name") or default_font_name
                            )
                            p.font.size = PptPt(max(10, tbl_font_size - 1))
                            p.font.color.rgb = PptRGB(*tbl_color)
                if animation:
                    _apply_element_animation(slide, tbl_shape, animation)
                rendered = True
            except Exception:
                logger.debug("could not render PPT table element", exc_info=True)
            continue

        # --- 引用元素（左边框强调）---
        if kind == "quote":
            text = strip_markdown(str(element.get("text") or "")).strip()
            if not text:
                continue
            border_color = parse_hex_color(
                element.get("border_color"), default=(26, 35, 126)
            )
            accent_w = max(int(slide_width * 0.006), 1)
            try:
                accent = slide.shapes.add_shape(1, left, top, accent_w, height)
                accent.fill.solid()
                accent.fill.fore_color.rgb = PptRGB(*border_color)
                accent.line.fill.background()
            except Exception:
                logger.debug("could not render PPT quote accent", exc_info=True)
            box = slide.shapes.add_textbox(
                left + accent_w + int(slide_width * 0.01),
                top,
                width - accent_w - int(slide_width * 0.01),
                height,
            )
            tf = box.text_frame
            tf.clear()
            tf.word_wrap = True
            color_rgb = _readable_color(
                parse_hex_color(element.get("color"), default=(85, 85, 85))
            )
            font_name = str(element.get("font_name") or default_font_name)
            font_size = _safe_float(element.get("font_size"), default_font_size + 2)
            paragraph = tf.paragraphs[0]
            paragraph.text = text
            paragraph.font.name = font_name
            paragraph.font.size = PptPt(max(12, font_size))
            paragraph.font.color.rgb = PptRGB(*color_rgb)
            paragraph.font.italic = bool(element.get("italic", True))
            paragraph.font.bold = bool(element.get("bold", False))
            align = _resolve_align(element.get("align"))
            if align is not None:
                paragraph.alignment = align
            if animation:
                _apply_element_animation(slide, box, animation)
            rendered = True
            continue

        # --- 编号列表元素 ---
        if kind == "numbered_list":
            items_raw = element.get("items")
            items = (
                [strip_markdown(str(item)).strip() for item in items_raw]
                if isinstance(items_raw, list)
                else []
            )
            items = [item for item in items if item]
            if not items:
                text = strip_markdown(str(element.get("text") or "")).strip()
                if text:
                    items = [line.strip() for line in text.splitlines() if line.strip()]
            if not items:
                continue
            color_rgb = _readable_color(
                parse_hex_color(element.get("color"), default=default_color_rgb)
            )
            number_color = _readable_color(
                parse_hex_color(element.get("number_color"), default=(26, 35, 126))
            )
            font_name = str(element.get("font_name") or default_font_name)
            font_size = _safe_float(element.get("font_size"), default_font_size)
            box = slide.shapes.add_textbox(left, top, width, height)
            tf = box.text_frame
            tf.clear()
            tf.word_wrap = True
            for idx, item in enumerate(items):
                paragraph = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                run_num = paragraph.add_run()
                run_num.text = f"{idx + 1}. "
                run_num.font.name = font_name
                run_num.font.size = PptPt(max(12, font_size))
                run_num.font.color.rgb = PptRGB(*number_color)
                run_num.font.bold = True
                run_text = paragraph.add_run()
                run_text.text = item
                run_text.font.name = font_name
                run_text.font.size = PptPt(max(12, font_size))
                run_text.font.color.rgb = PptRGB(*color_rgb)
                line_spacing = _safe_float(element.get("line_spacing"), 0)
                if line_spacing > 0:
                    try:
                        paragraph.line_spacing = line_spacing
                    except Exception:
                        logger.debug(
                            "could not apply PPT numbered-list line spacing",
                            exc_info=True,
                        )
            if animation:
                _apply_element_animation(slide, box, animation)
            rendered = True
            continue

        # --- 进度条元素 ---
        if kind == "progress_bar":
            try:
                value = _safe_float(element.get("value"), 50)
                max_value = _safe_float(element.get("max_value"), 100)
                pct = max(0.0, min(1.0, value / max_value)) if max_value > 0 else 0.5
                track_color = parse_hex_color(
                    element.get("track_color"), default=(224, 224, 224)
                )
                bar_color = parse_hex_color(
                    element.get("bar_color"), default=(76, 175, 80)
                )
                bar_height = max(int(slide_height * 0.02), 1)
                track = slide.shapes.add_shape(1, left, top, width, bar_height)
                track.fill.solid()
                track.fill.fore_color.rgb = PptRGB(*track_color)
                track.line.fill.background()
                bar_w = max(1, int(width * pct))
                bar = slide.shapes.add_shape(1, left, top, bar_w, bar_height)
                bar.fill.solid()
                bar.fill.fore_color.rgb = PptRGB(*bar_color)
                bar.line.fill.background()
                label = str(element.get("label") or "").strip()
                if label:
                    lbl_box = slide.shapes.add_textbox(
                        left,
                        top + bar_height + int(slide_height * 0.005),
                        width,
                        int(slide_height * 0.05),
                    )
                    lbl_tf = lbl_box.text_frame
                    lbl_tf.clear()
                    lbl_p = lbl_tf.paragraphs[0]
                    lbl_p.text = label
                    lbl_p.font.name = str(element.get("font_name") or default_font_name)
                    lbl_p.font.size = PptPt(
                        max(10, _safe_float(element.get("font_size"), 12))
                    )
                    lbl_p.font.color.rgb = PptRGB(
                        *parse_hex_color(
                            element.get("color"), default=default_color_rgb
                        )
                    )
                    align = _resolve_align(element.get("align"))
                    if align is not None:
                        lbl_p.alignment = align
                rendered = True
            except Exception:
                logger.debug("could not render PPT progress bar", exc_info=True)
            continue

        # --- 形状元素（带可选文本的几何形状）---
        if kind == "shape":
            try:
                from pptx.enum.shapes import MSO_SHAPE

                shape_map = {
                    "rectangle": MSO_SHAPE.RECTANGLE,
                    "rounded_rectangle": MSO_SHAPE.ROUNDED_RECTANGLE,
                    "circle": MSO_SHAPE.OVAL,
                    "oval": MSO_SHAPE.OVAL,
                    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
                    "arrow_right": MSO_SHAPE.RIGHT_ARROW,
                    "arrow_left": MSO_SHAPE.LEFT_ARROW,
                    "arrow_up": MSO_SHAPE.UP_ARROW,
                    "arrow_down": MSO_SHAPE.DOWN_ARROW,
                    "double_arrow": MSO_SHAPE.LEFT_RIGHT_ARROW,
                    "star": MSO_SHAPE.STAR_5_POINT,
                    "diamond": MSO_SHAPE.DIAMOND,
                    "hexagon": MSO_SHAPE.HEXAGON,
                    "pentagon": MSO_SHAPE.PENTAGON,
                    "parallelogram": MSO_SHAPE.PARALLELOGRAM,
                    "trapezoid": MSO_SHAPE.TRAPEZOID,
                    "chevron": MSO_SHAPE.CHEVRON,
                    "donut": MSO_SHAPE.DONUT,
                    "cloud": MSO_SHAPE.CLOUD,
                    "heart": MSO_SHAPE.HEART,
                }
                shape_type_str = (
                    str(element.get("shape_type") or "rounded_rectangle")
                    .strip()
                    .lower()
                )
                mso_shape = shape_map.get(shape_type_str, MSO_SHAPE.ROUNDED_RECTANGLE)
                shape = slide.shapes.add_shape(mso_shape, left, top, width, height)
                fill_color = parse_hex_color(element.get("fill_color"), default=None)
                if fill_color is not None:
                    shape.fill.solid()
                    shape.fill.fore_color.rgb = PptRGB(*fill_color)
                else:
                    shape.fill.background()
                border_color = parse_hex_color(
                    element.get("border_color"), default=None
                )
                if border_color is not None:
                    shape.line.color.rgb = PptRGB(*border_color)
                    shape.line.width = PptPt(1.5)
                else:
                    shape.line.fill.background()
                shape_text = strip_markdown(str(element.get("text") or "")).strip()
                if shape_text and hasattr(shape, "text_frame"):
                    tf = shape.text_frame
                    tf.clear()
                    tf.word_wrap = True
                    p = tf.paragraphs[0]
                    p.text = shape_text
                    p.font.name = str(element.get("font_name") or default_font_name)
                    p.font.size = PptPt(
                        max(
                            10, _safe_float(element.get("font_size"), default_font_size)
                        )
                    )
                    ai_text_color = parse_hex_color(
                        element.get("color"), default=default_color_rgb
                    )
                    if fill_color is not None:
                        text_color = _contrast_safe(ai_text_color, fill_color)
                    else:
                        text_color = _readable_color(ai_text_color)
                    p.font.color.rgb = PptRGB(*text_color)
                    p.alignment = (
                        _resolve_align(element.get("align")) or PP_ALIGN.CENTER
                    )
                if animation:
                    _apply_element_animation(slide, shape, animation)
                rendered = True
            except Exception:
                logger.debug("could not render PPT shape element", exc_info=True)
            continue

        # --- icon_text 元素（图标字符 + 文本卡片）---
        if kind == "icon_text":
            icon = str(element.get("icon") or "★").strip()[:2]
            text = strip_markdown(str(element.get("text") or "")).strip()
            if not text and not icon:
                continue
            color_rgb = _readable_color(
                parse_hex_color(element.get("color"), default=default_color_rgb)
            )
            icon_color = parse_hex_color(
                element.get("icon_color"), default=(255, 152, 0)
            )
            font_name = str(element.get("font_name") or default_font_name)
            font_size = _safe_float(element.get("font_size"), default_font_size)
            icon_size = _safe_float(element.get("icon_size"), 36)
            icon_h = int(height * 0.45)
            text_h = height - icon_h
            icon_box = slide.shapes.add_textbox(left, top, width, icon_h)
            icon_tf = icon_box.text_frame
            icon_tf.clear()
            icon_p = icon_tf.paragraphs[0]
            icon_p.text = icon
            icon_p.font.size = PptPt(max(16, icon_size))
            icon_p.font.color.rgb = PptRGB(*icon_color)
            icon_p.alignment = _resolve_align(element.get("align")) or PP_ALIGN.CENTER
            if text:
                txt_box = slide.shapes.add_textbox(left, top + icon_h, width, text_h)
                txt_tf = txt_box.text_frame
                txt_tf.clear()
                txt_tf.word_wrap = True
                txt_p = txt_tf.paragraphs[0]
                txt_p.text = text
                txt_p.font.name = font_name
                txt_p.font.size = PptPt(max(10, font_size))
                txt_p.font.color.rgb = PptRGB(*color_rgb)
                txt_p.alignment = (
                    _resolve_align(element.get("align")) or PP_ALIGN.CENTER
                )
            if animation:
                _apply_element_animation(slide, icon_box, animation)
            rendered = True
            continue

        text = strip_markdown(str(element.get("text") or "")).strip()
        items_raw = element.get("items")
        items = (
            [strip_markdown(str(item)).strip() for item in items_raw]
            if isinstance(items_raw, list)
            else []
        )
        items = [item for item in items if item]
        if kind == "list" and not items and text:
            items = [line.strip() for line in text.splitlines() if line.strip()]
        if not text and kind != "list":
            continue

        box = slide.shapes.add_textbox(left, top, width, height)
        tf = box.text_frame
        tf.clear()
        tf.word_wrap = True
        align = _resolve_align(element.get("align"))
        ai_color = parse_hex_color(element.get("color"), default=default_color_rgb)
        # 对带 background_color 的 callout 强制对比度
        callout_bg = (
            parse_hex_color(element.get("background_color"), default=None)
            if kind == "callout"
            else None
        )
        if callout_bg is not None:
            color_rgb = _contrast_safe(ai_color, callout_bg)
        else:
            color_rgb = _readable_color(ai_color)
        font_name = str(element.get("font_name") or default_font_name)
        size_default = (
            default_font_size + 10
            if kind == "title"
            else (default_font_size + 4 if kind == "subtitle" else default_font_size)
        )
        font_size = _safe_float(element.get("font_size"), size_default)
        bold_default = kind in {"title", "subtitle", "quote", "callout"}
        bold = bool(element.get("bold", bold_default))
        italic = bool(element.get("italic", False))
        underline = bool(element.get("underline", False))
        line_spacing = _safe_float(element.get("line_spacing"), 0)

        # 为 callout 类型应用文本框填充/边框
        if kind == "callout":
            _apply_shape_fill(
                box, element.get("background_color"), element.get("border_color")
            )

        if kind == "list":
            lines = items if items else ([text] if text else [])
            for idx, line in enumerate(lines):
                paragraph = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
                paragraph.text = line
                paragraph.level = (
                    int(_safe_float(element.get("level"), 0))
                    if idx == 0
                    else int(_safe_float(element.get("level"), 1))
                )
                paragraph.font.name = font_name
                paragraph.font.size = PptPt(max(12, font_size - (0 if idx == 0 else 1)))
                paragraph.font.color.rgb = PptRGB(*color_rgb)
                paragraph.font.bold = bold
                paragraph.font.italic = italic
                paragraph.font.underline = underline
                if align is not None:
                    paragraph.alignment = align
                if line_spacing > 0:
                    try:
                        paragraph.line_spacing = line_spacing
                    except Exception:
                        logger.debug(
                            "could not apply PPT list line spacing", exc_info=True
                        )
        else:
            paragraph = tf.paragraphs[0]
            paragraph.text = text
            paragraph.font.name = font_name
            paragraph.font.size = PptPt(max(12, font_size))
            paragraph.font.color.rgb = PptRGB(*color_rgb)
            paragraph.font.bold = bold
            paragraph.font.italic = italic
            paragraph.font.underline = underline
            if align is not None:
                paragraph.alignment = align
            if line_spacing > 0:
                try:
                    paragraph.line_spacing = line_spacing
                except Exception:
                    logger.debug("could not apply PPT text line spacing", exc_info=True)

        if animation:
            _apply_element_animation(slide, box, animation)

        rendered = True

    return rendered


def _build_fallback_layout_elements(
    slide_title: str, bullets: list[str], variant: str, has_image: bool
) -> list[dict]:
    from .locale import infer_text_language, localized_text

    lines = [
        strip_markdown(str(item)).strip()
        for item in bullets
        if strip_markdown(str(item)).strip()
    ]
    if not lines:
        language = infer_text_language(slide_title, "zh-CN")
        lines = [
            localized_text(
                language,
                **{
                    "zh-CN": "请补充该页核心内容",
                    "zh-TW": "請補充本頁核心內容",
                    "en-US": "Add the key content for this slide",
                    "ja-JP": "このスライドの主要内容を追加してください",
                    "ko-KR": "이 슬라이드의 핵심 내용을 추가하세요",
                    "ru-RU": "Добавьте ключевое содержание этого слайда",
                    "fr-FR": "Ajoutez le contenu essentiel de cette diapositive",
                    "de-DE": "Fügen Sie die Kernaussage dieser Folie hinzu",
                    "es-ES": "Añada el contenido clave de esta diapositiva",
                },
            )
        ]

    if variant == "split":
        elements: list[dict] = [
            {
                "kind": "title",
                "text": slide_title,
                "x": 6,
                "y": 5,
                "w": 88,
                "h": 10,
                "font_size": 32,
            },
            {
                "kind": "list",
                "items": lines,
                "x": 6,
                "y": 20,
                "w": 48,
                "h": 66,
                "font_size": 19,
            },
            {
                "kind": "text",
                "text": lines[0],
                "x": 58,
                "y": 22,
                "w": 34,
                "h": 18,
                "font_size": 24,
                "bold": True,
            },
        ]
        if has_image:
            elements.append({"kind": "image", "x": 58, "y": 44, "w": 34, "h": 40})
        return elements

    if variant == "focus":
        elements = [
            {
                "kind": "title",
                "text": slide_title,
                "x": 8,
                "y": 5,
                "w": 84,
                "h": 10,
                "font_size": 34,
                "align": "center",
            },
            {
                "kind": "text",
                "text": lines[0],
                "x": 10,
                "y": 26,
                "w": 80,
                "h": 20,
                "font_size": 30,
                "align": "center",
                "bold": True,
            },
        ]
        remaining = lines[1:]
        if remaining:
            elements.append(
                {
                    "kind": "list",
                    "items": remaining,
                    "x": 12,
                    "y": 52,
                    "w": 76,
                    "h": 32,
                    "font_size": 18,
                }
            )
        if has_image:
            elements.append({"kind": "image", "x": 36, "y": 56, "w": 28, "h": 30})
        return elements

    elements = [
        {
            "kind": "title",
            "text": slide_title,
            "x": 6,
            "y": 5,
            "w": 88,
            "h": 10,
            "font_size": 30,
        },
        {
            "kind": "list",
            "items": lines,
            "x": 6,
            "y": 20,
            "w": 52 if has_image else 88,
            "h": 66,
            "font_size": 18,
        },
    ]
    if has_image:
        elements.append({"kind": "image", "x": 60, "y": 22, "w": 34, "h": 56})
    return elements


def _source_template_layout(prs, options: dict):
    layouts = list(prs.slide_layouts)
    if not layouts:
        raise ValueError("PPT 模板没有可用的幻灯片布局")
    requested = options.get("template_layout_index")
    if requested is not None:
        if (
            isinstance(requested, bool)
            or not isinstance(requested, int)
            or not 0 <= requested < len(layouts)
        ):
            raise ValueError(
                "template_layout_index 必须是模板中有效的从 0 开始的布局编号"
            )
        return layouts[requested]
    # 模板不必遵循 python-pptx 的默认版式顺序。
    for layout in layouts:
        if layout.name.strip().lower() in {"blank", "空白"}:
            return layout
    return min(layouts, key=lambda layout: len(layout.placeholders))


def _template_theme(layout):
    from pptx.opc.constants import RELATIONSHIP_TYPE

    try:
        part = layout.slide_master.part.part_related_by(RELATIONSHIP_TYPE.THEME)
        return parse_xml(part.blob)
    except (KeyError, ValueError):
        return None


def _template_font(theme, family: str) -> str | None:
    if theme is None:
        return None
    family_node = theme.find(
        f"{qn('a:themeElements')}/{qn('a:fontScheme')}/{qn('a:' + family)}"
    )
    if family_node is None:
        return None
    # 中文文档应保留显式配置的东亚字体。
    for tag in ("ea", "latin"):
        node = family_node.find(qn("a:" + tag))
        if node is not None and node.get("typeface"):
            return node.get("typeface")
    return None


def _template_background(layout, theme) -> tuple[int, int, int] | None:
    """读取简单的继承实心填充用于文本对比度，但不覆盖它们。"""
    for source in (layout, layout.slide_master):
        color = source._element.find(
            f"{qn('p:cSld')}/{qn('p:bg')}/{qn('p:bgPr')}/{qn('a:solidFill')}"
        )
        if color is None or not len(color):
            continue
        color = color[0]
        if color.tag == qn("a:schemeClr") and theme is not None:
            token = color.get("val", "")
            mapping = layout.slide_master._element.find(qn("p:clrMap"))
            if mapping is not None:
                token = mapping.get(token, token)
            if token not in {
                "dk1",
                "lt1",
                "dk2",
                "lt2",
                "accent1",
                "accent2",
                "accent3",
                "accent4",
                "accent5",
                "accent6",
                "hlink",
                "folHlink",
            }:
                return None
            resolved = theme.find(
                f"{qn('a:themeElements')}/{qn('a:clrScheme')}/{qn('a:' + token)}"
            )
            if resolved is None or not len(resolved):
                return None
            color = resolved[0]
        raw = color.get("lastClr") if color.tag == qn("a:sysClr") else color.get("val")
        if color.tag in {qn("a:srgbClr"), qn("a:sysClr")} and raw:
            return parse_hex_color(raw, default=None)
        return None
    return None


def render_ppt(
    title: str,
    content: list[dict],
    styles: dict,
    ppt_options: dict | None = None,
    *,
    template_bytes: bytes | None = None,
) -> bytes:
    if template_bytes is not None:
        from .render import _guard_ooxml

        _guard_ooxml(template_bytes)
    options = ppt_options if isinstance(ppt_options, dict) else {}
    prs = (
        Presentation(io.BytesIO(template_bytes))
        if template_bytes is not None
        else Presentation()
    )
    if template_bytes is not None:
        # 移除示例内容，而不是下面要使用的母版/版式/主题部件。
        for slide_id in list(prs.slides._sldIdLst):
            prs.part.drop_rel(slide_id.rId)
            prs.slides._sldIdLst.remove(slide_id)
    slide_layout_blank = _source_template_layout(prs, options)
    template_theme = (
        _template_theme(slide_layout_blank) if template_bytes is not None else None
    )
    configured_defaults = options.get("document_defaults")
    doc_defaults = (
        configured_defaults
        if isinstance(configured_defaults, dict)
        else get_default_section("document")
    )
    style_map = styles if isinstance(styles, dict) else {}

    def _style_value(section: str, key: str, default: Any) -> Any:
        node = style_map.get(section)
        if isinstance(node, dict) and node.get(key) not in (None, ""):
            return node.get(key)
        return default

    body_font_name = str(
        _style_value(
            "body",
            "font_name",
            _template_font(template_theme, "minorFont")
            or doc_defaults.get("document_default_font_name", "宋体"),
        )
    )
    body_font_size = float(_style_value("body", "font_size", 20))
    heading_font_name = str(
        _style_value(
            "heading",
            "font_name",
            _template_font(template_theme, "majorFont") or body_font_name,
        )
    )
    heading_font_size = float(
        _style_value("heading", "font_size", max(24, body_font_size + 6))
    )
    ratio = _resolve_ppt_ratio(ppt_options)
    default_color_rgb = parse_hex_color(
        _style_value(
            "body", "color", doc_defaults.get("document_default_text_color", "#000000")
        )
    )
    heading_color_rgb = parse_hex_color(
        _style_value(
            "heading",
            "color",
            _style_value(
                "title",
                "color",
                doc_defaults.get("document_default_text_color", "#000000"),
            ),
        )
    )
    explicit_theme = (
        options.get("theme") if isinstance(options.get("theme"), dict) else {}
    )
    explicit_background = explicit_theme.get("background_color") or options.get(
        "background_color"
    )
    background_rgb = (
        _resolve_background_color(options)
        if template_bytes is None or explicit_background
        else None
    )
    inherited_background = (
        _template_background(slide_layout_blank, template_theme)
        if template_bytes is not None
        else None
    )

    if template_bytes is None or options.get("aspect_ratio") or options.get("ratio"):
        prs.slide_width = Inches(10 if ratio == "4:3" else 13.333)
        prs.slide_height = Inches(7.5)

    def _add_content_slide():
        slide = prs.slides.add_slide(slide_layout_blank)
        # 自由定位内容不得包含可见的空标题/正文占位框；
        # 静态插图和继承的母版/版式内容保持不变。
        for placeholder in list(slide.placeholders):
            placeholder._element.getparent().remove(placeholder._element)
        return slide

    slide_width = int(cast(int, getattr(prs, "slide_width", 0) or 0))
    slide_height = int(cast(int, getattr(prs, "slide_height", 0) or 0))

    def _add_slide_image(slide, image_bytes: bytes, caption: str | None = None) -> None:
        left = int(slide_width * 0.55)
        top = int(slide_height * 0.2)
        width = int(slide_width * 0.4)
        height = int(slide_height * 0.6)
        pic = _add_picture_keep_ratio(
            slide,
            image_bytes=image_bytes,
            left=left,
            top=top,
            width=width,
            height=height,
            mode="contain",
        )
        if pic is None:
            return
        if caption:
            cap_top = int(slide_height * 0.82)
            cap_box = slide.shapes.add_textbox(
                left, cap_top, width, int(slide_height * 0.1)
            )
            cap_tf = cap_box.text_frame
            cap_tf.text = strip_markdown(caption)
            for p in cap_tf.paragraphs:
                p.font.name = doc_defaults.get("document_default_font_name", "宋体")
                p.font.size = PptPt(11)
                p.font.color.rgb = PptRGB(90, 90, 90)

    def _resolve_slide_background(meta: dict | None) -> tuple[int, int, int] | None:
        if isinstance(meta, dict) and meta.get("background_color"):
            return parse_hex_color(meta.get("background_color"), default=background_rgb)
        return background_rgb

    # 注意：不自动生成标题页——是否包含标题页由 AI 作为内容节点的一部分决定。

    for idx, node in enumerate(content, start=1):
        kind = node.get("kind", "slide")
        text = strip_markdown(str(node.get("text") or ""))
        raw_meta = node.get("meta")
        meta = raw_meta if isinstance(raw_meta, dict) else {}
        is_slide = kind == "slide"
        font_name = str(
            node.get("font_name") or (heading_font_name if is_slide else body_font_name)
        )
        font_size = float(
            node.get("font_size") or (heading_font_size if is_slide else body_font_size)
        )
        color_rgb = parse_hex_color(
            node.get("color"),
            default=heading_color_rgb if is_slide else default_color_rgb,
        )

        if kind == "slide":
            slide = None
            slide_title = meta.get("title") or text
            bullets = meta.get("bullets") or meta.get("content") or []
            bullets_list = (
                bullets
                if isinstance(bullets, list)
                else [bullets]
                if isinstance(bullets, str)
                else []
            )
            raw_image = meta.get("_image_bytes") or meta.get("_image_b64")
            image_bytes = _decode_image_value(raw_image)
            raw_background = meta.get("_background_image_bytes") or meta.get(
                "_background_image_b64"
            )
            background_image_bytes = _decode_image_value(raw_background)
            layout_elements = _extract_layout_elements(meta)
            # Office schema 允许显式的 meta.title 加上 node.text 正文。
            # 没有单独标题时，text 仍作为标题；无论哪种情况，
            # 显式的 layout/bullet 内容都具有权威性。
            if (
                not layout_elements
                and not bullets_list
                and meta.get("title")
                and text.strip()
                and text.strip() != strip_markdown(str(slide_title)).strip()
            ):
                bullets_list = [text]

            if layout_elements:
                image_mode = str(meta.get("image_mode") or "").strip().lower()
                layout_elements = _ensure_foreground_image_slot(
                    layout_elements,
                    slide_index=idx,
                    enabled=isinstance(image_bytes, (bytes, bytearray))
                    and image_mode in {"search", "generate"},
                )
                slide = _add_content_slide()
                _apply_slide_background(slide, _resolve_slide_background(meta))
                if background_image_bytes is not None:
                    _apply_slide_background_image(
                        slide, background_image_bytes, slide_width, slide_height
                    )
                slide_transition = str(meta.get("transition") or "fade").strip().lower()
                _apply_slide_transition(slide, slide_transition)
                rendered = _render_layout_elements(
                    slide,
                    layout_elements,
                    slide_width,
                    slide_height,
                    font_name,
                    max(14.0, font_size - 2),
                    color_rgb,
                    image_bytes,
                    has_background_image=background_image_bytes is not None,
                    slide_bg_rgb=_resolve_slide_background(meta)
                    or inherited_background,
                )
                if rendered:
                    # 检测 title-only 布局：若自定义 elements 仅有标题/分割线/图片，
                    # 或虽有正文类元素但内容为空，则把 bullets 作为 list 元素合入后再渲染一次，
                    # 避免只显示标题的空壳页。
                    _kind = lambda e: str(e.get("kind") or e.get("type") or "").lower()
                    _CONTENT_KINDS = {
                        "list",
                        "text",
                        "callout",
                        "table",
                        "shape",
                        "numbered_list",
                        "quote",
                        "icon_text",
                    }

                    def _el_has_content(el: dict) -> bool:
                        k = _kind(el)
                        if k in {"list", "numbered_list"}:
                            items = el.get("items")
                            return bool(
                                isinstance(items, list)
                                and any(str(i).strip() for i in items)
                            )
                        if k in {"text", "callout", "quote"}:
                            return bool(str(el.get("text") or "").strip())
                        if k == "table":
                            rows = el.get("rows")
                            return bool(isinstance(rows, list) and rows)
                        if k in {"shape", "icon_text"}:
                            return bool(str(el.get("text") or "").strip())
                        return False

                    _has_real_content = any(
                        _kind(e) in _CONTENT_KINDS and _el_has_content(e)
                        for e in layout_elements
                        if isinstance(e, dict)
                    )
                    if not _has_real_content and bullets_list:
                        fallback_content = {
                            "kind": "list",
                            "items": bullets_list,
                            "x": 6,
                            "y": 22,
                            "w": 88,
                            "h": 66,
                            "font_size": 18,
                        }
                        layout_elements.append(fallback_content)
                        _render_layout_elements(
                            slide,
                            [fallback_content],
                            slide_width,
                            slide_height,
                            font_name,
                            max(14.0, font_size - 2),
                            color_rgb,
                            image_bytes,
                            has_background_image=background_image_bytes is not None,
                            slide_bg_rgb=_resolve_slide_background(meta)
                            or inherited_background,
                        )

                    continue

            variant_raw = str(meta.get("layout_variant") or "").strip().lower()
            variants = ["classic", "split", "focus"]
            variant = (
                variant_raw
                if variant_raw in variants
                else variants[idx % len(variants)]
            )
            fallback_elements = _build_fallback_layout_elements(
                slide_title, bullets_list, variant, image_bytes is not None
            )
            # 未知/空的自定义元素回退到现有页面，
            # 而不是静默地额外添加一张空白幻灯片。
            if slide is None:
                slide = _add_content_slide()
            _apply_slide_background(slide, _resolve_slide_background(meta))
            if background_image_bytes is not None:
                _apply_slide_background_image(
                    slide, background_image_bytes, slide_width, slide_height
                )
            slide_transition = str(meta.get("transition") or "fade").strip().lower()
            _apply_slide_transition(slide, slide_transition)
            _render_layout_elements(
                slide,
                fallback_elements,
                slide_width,
                slide_height,
                font_name,
                max(14.0, font_size - 2),
                color_rgb,
                image_bytes,
                has_background_image=background_image_bytes is not None,
                slide_bg_rgb=_resolve_slide_background(meta) or inherited_background,
            )
            if isinstance(image_bytes, (bytes, bytearray)) and not any(
                str(el.get("kind") or el.get("type") or "").lower() == "image"
                for el in fallback_elements
            ):
                _add_slide_image(slide, bytes(image_bytes), meta.get("image_caption"))
            continue

        elif kind == "paragraph":
            slide = _add_content_slide()
            _apply_slide_background(slide, background_rgb)
            _apply_slide_transition(slide)
            tx_box = slide.shapes.add_textbox(
                Inches(0.5), Inches(1), Inches(9), Inches(5)
            )
            tx_box.text_frame.text = strip_markdown(text)
            for p in tx_box.text_frame.paragraphs:
                p.font.name = body_font_name
                p.font.size = PptPt(body_font_size)
                p.font.color.rgb = PptRGB(*default_color_rgb)

    # 确保演示文稿至少有一张幻灯片
    if not prs.slides:
        slide = _add_content_slide()
        _apply_slide_background(slide, background_rgb)
        _apply_slide_transition(slide)
        tx_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(3))
        tf = tx_box.text_frame
        tf.text = strip_markdown(title or "演示文稿")
        for p in tf.paragraphs:
            p.font.name = heading_font_name
            p.font.size = PptPt(heading_font_size)
            p.font.color.rgb = PptRGB(*heading_color_rgb)

    buf = io.BytesIO()
    from .pptx_text_styles import materialize_paragraph_styles

    materialize_paragraph_styles(prs)
    prs.save(buf)
    return buf.getvalue()
