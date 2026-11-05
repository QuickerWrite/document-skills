"""用于模型提示词和用户可见生成文本的共享语言策略。"""

from __future__ import annotations

import re
from typing import Any


_LANGUAGE_NAMES = {
    "zh-CN": ("简体中文", "Simplified Chinese"),
    "zh-TW": ("繁体中文", "Traditional Chinese"),
    "en-US": ("英语", "English"),
    "ja-JP": ("日语", "Japanese"),
    "ko-KR": ("韩语", "Korean"),
    "ru-RU": ("俄语", "Russian"),
    "fr-FR": ("法语", "French"),
    "de-DE": ("德语", "German"),
    "es-ES": ("西班牙语", "Spanish"),
}


def normalize_language_tag(value: Any, default: str = "zh-CN") -> str:
    raw = str(value or "").strip().replace("_", "-")
    if not raw:
        return default
    primary = raw.split(",", 1)[0].split(";", 1)[0].strip()
    lowered = primary.lower()
    if lowered in {"zh", "zh-cn", "zh-hans", "zh-sg"}:
        return "zh-CN"
    if lowered in {"zh-tw", "zh-hant", "zh-hk", "zh-mo"}:
        return "zh-TW"
    defaults = {
        "en": "en-US",
        "ja": "ja-JP",
        "ko": "ko-KR",
        "ru": "ru-RU",
        "fr": "fr-FR",
        "de": "de-DE",
        "es": "es-ES",
    }
    if lowered in defaults:
        return defaults[lowered]
    if re.fullmatch(r"[A-Za-z]{2,3}(?:-[A-Za-z0-9]{2,8}){0,2}", primary):
        parts = primary.split("-")
        return "-".join(
            [
                parts[0].lower(),
                *[
                    part.upper() if len(part) == 2 else part.title()
                    for part in parts[1:]
                ],
            ]
        )
    return default


def request_language(request: Any, default: str = "zh-CN") -> str:
    headers = getattr(request, "headers", {}) or {}
    explicit = headers.get("X-Client-Language") or headers.get("X-App-Language")
    if explicit:
        return normalize_language_tag(explicit, default)
    return normalize_language_tag(headers.get("Accept-Language"), default)


def language_name(language: Any) -> tuple[str, str]:
    tag = normalize_language_tag(language)
    return _LANGUAGE_NAMES.get(tag, (tag, tag))


def infer_text_language(value: Any, fallback: Any = "zh-CN") -> str:
    """在不借助模型的情况下推断主要文字体系。

    这里刻意保持保守。对于拉丁字母语言，调用方提供的客户端区域设置优先，
    因为短文本无法可靠区分英语、法语、德语和西班牙语。
    """
    text = str(value or "")
    fallback_tag = normalize_language_tag(fallback)
    counts = {
        "zh-CN": len(re.findall(r"[\u3400-\u9fff]", text)),
        "ja-JP": len(re.findall(r"[\u3040-\u30ff]", text)),
        "ko-KR": len(re.findall(r"[\uac00-\ud7af]", text)),
        "ru-RU": len(re.findall(r"[\u0400-\u04ff]", text)),
        "latin": len(re.findall(r"[A-Za-z]", text)),
    }
    if counts["ja-JP"]:
        return "ja-JP"
    if counts["ko-KR"]:
        return "ko-KR"
    if counts["ru-RU"] >= max(2, counts["zh-CN"], counts["latin"] // 2):
        return "ru-RU"
    if counts["zh-CN"] >= max(2, counts["latin"] // 3):
        return "zh-TW" if fallback_tag == "zh-TW" else "zh-CN"
    if counts["latin"]:
        return (
            fallback_tag
            if fallback_tag.startswith(("en", "fr", "de", "es"))
            else "en-US"
        )
    return fallback_tag


def localized_text(language: Any, **values: str) -> str:
    """按 BCP-47 语言选择一小段确定性的 UI/回退文本。"""
    tag = normalize_language_tag(language)
    value = values.get(tag) or values.get(tag.split("-", 1)[0])
    return (
        value
        or values.get("en-US")
        or values.get("en")
        or next(iter(values.values()), "")
    )
