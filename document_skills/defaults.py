"""Portable rendering defaults; override through document_defaults options."""


def get_default_section(section):
    return (
        {
            "document_default_font_name": "Noto Sans CJK SC",
            "document_heading_font_size": "16",
            "document_body_font_size": "12",
            "document_default_text_color": "#000000",
        }
        if section == "document"
        else {}
    )
