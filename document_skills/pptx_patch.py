"""Inspect and patch selected PPTX objects without regenerating other ZIP parts."""
from io import BytesIO
import math
import re
import zipfile
from lxml import etree
from pptx import Presentation

NS = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
      'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def inspect_pptx(data, slide_number=None):
    prs = Presentation(BytesIO(data))
    pages = []
    for number, slide in enumerate(prs.slides, 1):
        if slide_number and number != slide_number:
            continue
        elements = []
        for shape in slide.shapes:
            item = {'element_id': shape.shape_id, 'name': shape.name,
                    'x': shape.left / prs.slide_width * 100, 'y': shape.top / prs.slide_height * 100,
                    'w': shape.width / prs.slide_width * 100, 'h': shape.height / prs.slide_height * 100,
                    'text': shape.text if shape.has_text_frame else '',
                    'text_editable': shape.has_text_frame}
            if shape.has_text_frame:
                run = next((r for p in shape.text_frame.paragraphs for r in p.runs), None)
                if run:
                    item['font_size'] = run.font.size.pt if run.font.size else None
            elements.append(item)
        pages.append({'page': number, 'elements': elements})
    if slide_number and not pages:
        raise ValueError('invalid_slide_number')
    return {'total_pages': len(prs.slides), 'slides': pages, 'geometry_unit': 'percent', 'font_size_unit': 'pt'}


def patch_pptx(data, patches):
    from .render import _guard_ooxml
    _guard_ooxml(data)
    if not isinstance(patches, list) or not 1 <= len(patches) <= 50:
        raise ValueError('patches must contain 1 to 50 element changes')
    prs = Presentation(BytesIO(data))
    changed = {}
    with zipfile.ZipFile(BytesIO(data)) as source:
        if any(n.startswith('_xmlsignatures/') for n in source.namelist()):
            raise ValueError('Signed presentations cannot be patched')
        for patch in patches:
            number, element_id = patch.get('slide_number'), patch.get('element_id')
            if not isinstance(number, int) or not 1 <= number <= len(prs.slides):
                raise ValueError('invalid_slide_number')
            # Resolve through package relationships; slide order need not match filenames.
            name = str(prs.slides[number - 1].part.partname).lstrip('/')
            if name not in changed:
                changed[name] = etree.fromstring(source.read(name), etree.XMLParser(resolve_entities=False, no_network=True))
            root = changed[name]
            shapes = root.xpath('./p:cSld/p:spTree/*[p:nvSpPr/p:cNvPr/@id=$id or p:nvPicPr/p:cNvPr/@id=$id or p:nvGraphicFramePr/p:cNvPr/@id=$id]', namespaces=NS, id=str(element_id))
            if len(shapes) != 1:
                raise ValueError('element_id must identify an existing top-level element; read again')
            shape = shapes[0]
            changes = patch.get('changes')
            allowed = {'x', 'y', 'w', 'h', 'text', 'font_size', 'font_name', 'color', 'bold', 'italic'}
            if not isinstance(changes, dict) or not changes or set(changes) - allowed:
                raise ValueError('unsupported or empty element changes')
            if set(changes) & {'x', 'y', 'w', 'h'}:
                transforms = shape.xpath('./p:spPr/a:xfrm | ./p:xfrm', namespaces=NS)
                if not transforms:
                    raise ValueError('element has inherited geometry; use an explicit object')
                transform = transforms[0]
                off, ext = transform.find('a:off', NS), transform.find('a:ext', NS)
                if off is None or ext is None:
                    raise ValueError('element geometry unavailable')
                values = {'x': int(off.get('x')) / prs.slide_width * 100, 'y': int(off.get('y')) / prs.slide_height * 100,
                          'w': int(ext.get('cx')) / prs.slide_width * 100, 'h': int(ext.get('cy')) / prs.slide_height * 100}
                for key in values:
                    if key in changes:
                        value = float(changes[key])
                        if not math.isfinite(value) or not 0 <= value <= 100:
                            raise ValueError('geometry must be finite percentages between 0 and 100')
                        values[key] = value
                if values['w'] <= 0 or values['h'] <= 0 or values['x'] + values['w'] > 100.001 or values['y'] + values['h'] > 100.001:
                    raise ValueError('element must remain inside slide')
                for key, node, attr, size in [('x', off, 'x', prs.slide_width), ('y', off, 'y', prs.slide_height), ('w', ext, 'cx', prs.slide_width), ('h', ext, 'cy', prs.slide_height)]:
                    node.set(attr, str(round(values[key] / 100 * size)))
            if set(changes) - {'x', 'y', 'w', 'h'}:
                body = shape.find('p:txBody', NS)
                if body is None:
                    raise ValueError('element does not contain editable text')
                if 'text' in changes:
                    text = changes['text']
                    if not isinstance(text, str) or len(text) > 20000:
                        raise ValueError('text must contain at most 20000 characters')
                    texts = body.findall('.//a:t', NS)
                    if not texts:
                        raise ValueError('empty element has no text run; select another element')
                    texts[0].text = text
                    for node in texts[1:]:
                        node.text = ''
                for run in body.findall('.//a:r', NS):
                    props = run.find('a:rPr', NS)
                    if props is None:
                        props = etree.Element('{%s}rPr' % NS['a']); run.insert(0, props)
                    if 'font_size' in changes:
                        value = float(changes['font_size'])
                        if not math.isfinite(value) or not 6 <= value <= 144:
                            raise ValueError('font_size must be between 6 and 144 pt')
                        props.set('sz', str(round(value * 100)))
                    for key, attr in [('bold', 'b'), ('italic', 'i')]:
                        if key in changes:
                            if not isinstance(changes[key], bool):
                                raise ValueError(key + ' must be boolean')
                            props.set(attr, '1' if changes[key] else '0')
                    if 'color' in changes:
                        if not re.fullmatch(r'#[0-9a-fA-F]{6}', str(changes['color'])):
                            raise ValueError('color must be #RRGGBB')
                        for child in list(props):
                            if etree.QName(child).localname in {'solidFill', 'noFill', 'gradFill', 'pattFill', 'blipFill', 'grpFill'}:
                                props.remove(child)
                        fill = etree.Element('{%s}solidFill' % NS['a']); props.insert(0, fill)
                        etree.SubElement(fill, '{%s}srgbClr' % NS['a'], val=changes['color'][1:])
                    if 'font_name' in changes:
                        font = changes['font_name']
                        if not isinstance(font, str) or not font.strip() or len(font) > 128 or any(ord(c) < 32 for c in font):
                            raise ValueError('invalid font_name')
                        for tag in ['latin', 'ea', 'cs']:
                            child = props.find('a:' + tag, NS)
                            if child is None: child = etree.SubElement(props, '{%s}%s' % (NS['a'], tag))
                            child.set('typeface', font)
        output = BytesIO()
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as target:
            for item in source.infolist():
                value = etree.tostring(changed[item.filename], encoding='UTF-8', xml_declaration=True, standalone=True) if item.filename in changed else source.read(item)
                target.writestr(item, value)
    return output.getvalue()
