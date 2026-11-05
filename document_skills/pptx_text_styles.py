"""为兼容 Office 的轻量查看器实体化段落默认值。"""

from copy import deepcopy


def materialize_paragraph_styles(presentation):
    # 这不会推断颜色，也不会替换显式的 run 格式。Python-pptx 会将
    # paragraph.font 写入 defRPr，而某些浏览器查看器会忽略它。
    for slide in presentation.slides:
        for paragraph in slide._element.xpath(".//a:p"):
            defaults = paragraph.xpath("./a:pPr/a:defRPr")
            if not defaults:
                continue
            default = defaults[0]
            for run in paragraph.xpath("./a:r"):
                props = run.get_or_add_rPr()
                for name, value in default.attrib.items():
                    if name not in props.attrib:
                        props.set(name, value)
                for child in default:
                    if not any(existing.tag == child.tag for existing in props):
                        # 填充选项是互斥的；绝不能用段落的不同填充类型
                        # 覆盖显式的 run 填充。
                        fill_tags = {
                            "solidFill",
                            "noFill",
                            "gradFill",
                            "pattFill",
                            "blipFill",
                            "grpFill",
                        }
                        local = child.tag.rsplit("}", 1)[-1]
                        if local in fill_tags and any(
                            c.tag.rsplit("}", 1)[-1] in fill_tags for c in props
                        ):
                            continue
                        props.append(deepcopy(child))
