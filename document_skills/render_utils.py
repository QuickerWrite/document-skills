"""Office 输出的共享渲染辅助函数。"""

from typing import overload


@overload
def parse_hex_color(
    color_value: str | None, default: tuple[int, int, int]
) -> tuple[int, int, int]: ...
@overload
def parse_hex_color(
    color_value: str | None, default: None
) -> tuple[int, int, int] | None: ...
@overload
def parse_hex_color(color_value: str | None) -> tuple[int, int, int]: ...
def parse_hex_color(
    color_value: str | None,
    default: tuple[int, int, int] | None = (0, 0, 0),
) -> tuple[int, int, int] | None:
    if not color_value:
        return default
    color_hex = str(color_value).strip().lstrip("#")
    if len(color_hex) != 6:
        return default
    try:
        return int(color_hex[0:2], 16), int(color_hex[2:4], 16), int(color_hex[4:6], 16)
    except ValueError:
        return default


def strip_markdown(text: str) -> str:
    cleaned = (text or "").strip()
    if not cleaned:
        return cleaned
    cleaned = cleaned.replace("**", "").replace("__", "").replace("`", "")
    while cleaned.startswith("#"):
        cleaned = cleaned[1:].lstrip()
    if cleaned.startswith("- "):
        cleaned = cleaned[2:].lstrip()
    if len(cleaned) > 2 and cleaned[0].isdigit() and cleaned[1] == ".":
        cleaned = cleaned[2:].lstrip()
    return cleaned
