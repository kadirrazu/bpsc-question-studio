"""Import Word table rows and assemble question booklets without retyping them."""

import io
import re
from copy import deepcopy

import qrcode
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Emu, Inches, Pt, RGBColor
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docxcompose.composer import Composer
from PIL import Image
from PySide6.QtCore import QBuffer, QPointF, Qt
from PySide6.QtGui import QColor, QFont, QFontMetrics, QImage, QPainter, QTransform

from .project import footer_parts, label, post_name, set_number_text, template_path, text_appearance

SERIAL = re.compile(r"^([0-9০-৯]+)([.)।]?)$")
TEXT_TAGS = (("[[SET]]", "[[SET_NUMBER]]"), ("[[CODE]]", "[[SET_CODE_NAME]]"))
QR_TAGS = ("[[QR]]", "[[SET_QR]]")


def text_tag_values(number, code_name, language):
    values = (set_number_text(number, language), code_name)
    return {tag: value for aliases, value in zip(TEXT_TAGS, values) for tag in aliases}


def open_word(path):
    try:
        return Document(path)
    except Exception as error:
        raise ValueError(f"Cannot read this Word document. Open and save it as .docx first.\n{error}") from error


def question_rows(document, expected):
    if len(document.sections) != 1:
        raise ValueError("Use a question document with one Word section.")
    if len(document.tables) != 1:
        raise ValueError("Use one top-level table with two columns: serial and question.")
    table = document.tables[0]
    if len(table.columns) != 2:
        raise ValueError("The question table must have exactly two columns.")

    rows = {}
    for index, row in enumerate(table.rows, 1):
        cells = row.cells
        if len(cells) != 2 or cells[0]._tc is cells[1]._tc or len(row._tr.tc_lst) != 2:
            raise ValueError(f"Table row {index} must have two separate cells.")
        if row._tr.xpath(".//w:vMerge"):
            raise ValueError(f"Table row {index} uses a vertical merge. Give each question its own row.")
        serial_text = cells[0].text.strip()
        content = cells[1]._tc.xpath(".//w:drawing | .//w:pict | .//w:object | .//m:oMath | .//w:tbl")
        if not serial_text and not cells[1].text.strip() and not content:
            continue
        match = SERIAL.fullmatch(serial_text)
        if not match:
            raise ValueError(f"Table row {index}: enter a Bengali or English numeric serial.")
        serial = int(match[1])
        if serial in rows:
            raise ValueError(f"Question serial {serial} appears more than once.")
        if not cells[1].text.strip() and not content:
            raise ValueError(f"Question {serial} has an empty question cell.")
        rows[serial] = row

    missing = sorted(set(range(1, expected + 1)) - rows.keys())
    extra = sorted(rows.keys() - set(range(1, expected + 1)))
    if missing or extra:
        details = [f"Found {len(rows)} filled rows. The mapping expects serials 1 to {expected}."]
        for title, serials in (("Missing", missing), ("Out of range", extra)):
            if serials:
                example = ", ".join(map(str, serials[:12]))
                if len(serials) > 12:
                    example += f" … ({len(serials)} total)"
                details.append(f"{title}: {example}.")
        raise ValueError("\n".join(details))
    return table, rows


def template_paragraphs(document):
    return [Paragraph(element, document._body)
            for element in document.element.body.iter(qn("w:p"))]


def replace_text(paragraph, token, replacement=None, picture=None, width=None, *,
                 offset=None, template_text=False, rotation=0, style_settings=None, vertical_shift=0):
    """Replace a marker even when Word has split it across formatting runs."""
    runs = paragraph.runs
    text = "".join(run.text for run in runs)
    offset = text.find(token) if offset is None else offset
    if offset < 0:
        return False
    end = offset + len(token)
    touched = []
    position = 0
    for run in runs:
        stop = position + len(run.text)
        if position < end and stop > offset:
            touched.append((run, max(0, offset - position), min(stop, end) - position))
        position = stop
    first, start, _ = touched[0]
    last, _, finish = touched[-1]
    prefix, suffix = first.text[:start], last.text[finish:]
    if template_text and re.search(r"[ঀ-৿]", replacement):
        preserve_bengali_style(paragraph, first)
    if rotation and replacement:
        picture, width = rotated_text_picture(paragraph, first, replacement, rotation, style_settings)
    for run, _, _ in touched:
        if run._r.xpath("./w:drawing | ./w:pict | ./w:object"):
            raise ValueError("Put each placeholder in a plain text run next to embedded objects.")
        run.text = ""
    first.text = prefix
    if picture:
        placeholder = OxmlElement("w:r")
        if first._r.rPr is not None:
            placeholder.append(deepcopy(first._r.rPr))
        first._r.addnext(placeholder)
        image_run = Run(placeholder, paragraph)
        image = image_run.add_picture(io.BytesIO(picture), width=width)
        if rotation:
            image._inline.docPr.set("descr", replacement)
        # Writer applies a default horizontal gap when the inline distances
        # are omitted, moving a right-aligned image left of the original tag.
        for side in ("T", "B", "L", "R"):
            image._inline.set(f"dist{side}", "0")
        if vertical_shift:
            shift_inline_picture(image_run, image._inline, vertical_shift)
        if first is last and suffix:
            tail = OxmlElement("w:r")
            placeholder.addnext(tail)
            if first._r.rPr is not None:
                tail.append(deepcopy(first._r.rPr))
            Run(tail, paragraph).text = suffix
        elif first is not last:
            last.text = suffix
    else:
        first.text = prefix + replacement + (suffix if first is last else "")
        if first is not last:
            last.text = suffix
    return True


def shift_inline_picture(run, inline, shift):
    # Reserve the original line width and height, then position the visible QR
    # relative to the character after that space. Word spacing cannot be negative.
    png = io.BytesIO()
    Image.new("RGBA", (1, 1)).save(png, format="PNG")
    png.seek(0)
    width, height = inline.extent.cx, inline.extent.cy
    spacer = run.add_picture(png, width=width, height=height)._inline
    for side in ("T", "B", "L", "R"):
        spacer.set(f"dist{side}", "0")
    drawing = inline.getparent()
    spacer.getparent().addnext(drawing)

    anchor = OxmlElement("wp:anchor")
    for key, value in {
        "distT": "0", "distB": "0", "distL": "0", "distR": "0", "simplePos": "0",
        "relativeHeight": "1", "behindDoc": "0", "locked": "0", "layoutInCell": "1", "allowOverlap": "1",
    }.items():
        anchor.set(key, value)
    origin = OxmlElement("wp:simplePos")
    origin.set("x", "0")
    origin.set("y", "0")
    anchor.append(origin)
    for axis, relative_to, amount in (("H", "character", -width), ("V", "line", shift)):
        position = OxmlElement("wp:position" + axis)
        position.set("relativeFrom", relative_to)
        offset = OxmlElement("wp:posOffset")
        offset.text = str(amount)
        position.append(offset)
        anchor.append(position)
    for child in list(inline):
        if child.tag == qn("wp:docPr"):
            anchor.append(OxmlElement("wp:wrapNone"))
        anchor.append(child)
    drawing.replace(inline, anchor)


def rotated_text_picture(paragraph, run, text, angle, settings):
    # Vertical drawings keep the rotation local to the tag, including when a
    # QR or other text shares the same cell.
    properties = list(run._r.xpath("./w:rPr"))
    styles = (run.style if run._r.xpath("./w:rPr/w:rStyle") else None, paragraph.style)
    for style in styles:
        while style is not None:
            properties.extend(style.element.xpath("./w:rPr"))
            style = style.base_style
    properties.extend(paragraph.part.styles.element.xpath("./w:docDefaults/w:rPrDefault/w:rPr"))

    def value(*keys, attribute="w:val"):
        for props in properties:
            for key in keys:
                node = props.find(qn(key))
                if node is not None:
                    found = node.get(qn(attribute), "1" if key in ("w:b", "w:i", "w:bCs", "w:iCs") else None)
                    if found is not None:
                        return found
        return None

    bengali = re.search(r"[ঀ-৿]", text)
    family = value("w:rFonts", attribute="w:cs" if bengali else "w:ascii")
    family = family or value("w:rFonts", attribute="w:hAnsi") or settings["font"]
    size = value("w:szCs", "w:sz") if bengali else value("w:sz")
    point_size = int(size) / 2 if size else settings["font_size"]
    scale = 6
    font = QFont(family)
    font.setPixelSize(round(point_size * scale))
    font.setBold(value(*(("w:bCs", "w:b") if bengali else ("w:b",))) in ("1", "true", "on"))
    font.setItalic(value(*(("w:iCs", "w:i") if bengali else ("w:i",))) in ("1", "true", "on"))
    bounds = QFontMetrics(font).tightBoundingRect(text).adjusted(-2, -2, 2, 2)
    image = QImage(bounds.size(), QImage.Format.Format_ARGB32_Premultiplied)
    image.fill(Qt.GlobalColor.transparent)
    color = QColor("#" + str(value("w:color")))
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
    painter.setFont(font)
    painter.setPen(color if color.isValid() else QColor("black"))
    painter.drawText(QPointF(-bounds.x(), -bounds.y()), text)
    painter.end()
    rotated = image.transformed(QTransform().rotate(angle))
    buffer = QBuffer()
    if not rotated.save(buffer, "PNG"):
        raise ValueError("Unable to draw a rotated template tag.")
    return bytes(buffer.data()), Pt(rotated.width() / scale)


def preserve_bengali_style(paragraph, run):
    """Use the marker’s typography when the replacement switches text script."""
    direct = run._r.get_or_add_rPr()
    style_layers = [direct]
    for style in (run.style if direct.find(qn("w:rStyle")) is not None else None, paragraph.style):
        while style is not None:
            props = style.element.find(qn("w:rPr"))
            if props is not None:
                style_layers.append(props)
            style = style.base_style
    style_layers.extend(paragraph.part.styles.element.xpath("./w:docDefaults/w:rPrDefault/w:rPr"))
    for latin, complex_script in (("w:sz", "w:szCs"), ("w:b", "w:bCs"), ("w:i", "w:iCs")):
        for layer in style_layers:
            original = layer.find(qn(complex_script))
            if original is None:
                original = layer.find(qn(latin))
            if original is not None:
                value = deepcopy(original)
                value.tag = qn(complex_script)
                old = direct.find(value.tag)
                if old is not None:
                    direct.remove(old)
                if complex_script == "w:szCs":
                    direct.insert_element_before(value, "w:highlight", "w:u", "w:vertAlign", "w:rtl", "w:cs", "w:lang")
                elif complex_script == "w:bCs":
                    direct._insert_bCs(value)
                else:
                    direct._insert_iCs(value)
                break
    for layer in style_layers:
        font = layer.find(qn("w:rFonts"))
        if font is None:
            continue
        family = font.get(qn("w:cs"))
        theme = font.get(qn("w:cstheme"))
        if not (family or theme):
            family = font.get(qn("w:ascii")) or font.get(qn("w:hAnsi"))
            theme = font.get(qn("w:asciiTheme")) or font.get(qn("w:hAnsiTheme"))
        if family or theme:
            target = direct.get_or_add_rFonts()
            if family:
                target.attrib.pop(qn("w:cstheme"), None)
                target.set(qn("w:cs"), family)
            else:
                target.set(qn("w:cstheme"), theme)
            break


def align_tag_paragraph(paragraph, side, *, standalone):
    if side != "Template":
        paragraph.alignment = getattr(WD_ALIGN_PARAGRAPH, side.upper())
        if standalone:
            paragraph.paragraph_format.left_indent = Pt(0)
            paragraph.paragraph_format.right_indent = Pt(0)
            paragraph.paragraph_format.first_line_indent = Pt(0)
        return standalone
    return False


def align_tag_cell(cell, side):
    if side != "Template":
        cell.get_or_add_tcPr().get_or_add_vAlign().val = getattr(WD_CELL_VERTICAL_ALIGNMENT, side.upper())


def align_qr(paragraph, settings, *, standalone):
    side = settings["qr_horizontal"]
    strip_padding = align_tag_paragraph(paragraph, side, standalone=standalone)
    if side == "Template" and standalone:
        style = paragraph.style
        formats = [paragraph.paragraph_format]
        while style is not None:
            formats.append(style.paragraph_format)
            style = style.base_style
        # Keep explicit tab stops in template mode. A leading default tab would
        # offset a centered or right-aligned image, so remove it on those lines.
        tab_positioned = "\t" in paragraph.text and any(len(fmt.tab_stops) for fmt in formats)
        for fmt in formats:
            if fmt.alignment is not None:
                paragraph.alignment = fmt.alignment
                strip_padding = not tab_positioned and fmt.alignment in (
                    WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT)
                break
    if standalone and settings["qr_inset"]:
        fmt = paragraph.paragraph_format
        if paragraph.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            fmt.right_indent = (fmt.right_indent or 0) + Inches(settings["qr_inset"])
        else:
            fmt.left_indent = (fmt.left_indent or 0) + Inches(settings["qr_inset"])
    side = settings["qr_vertical"]
    cells = paragraph._p.xpath("ancestor::w:tc[1]")
    if cells:
        align_tag_cell(cells[0], side)
    bottom = side == "Bottom" or (side == "Template" and cells and
                                  cells[0].xpath("./w:tcPr/w:vAlign[@w:val='bottom']"))
    gap = settings["qr_gap"] if standalone else 0
    if gap > 0:
        fmt = paragraph.paragraph_format
        if bottom:
            fmt.space_after = (fmt.space_after or 0) + Inches(gap)
        else:
            fmt.space_before = (fmt.space_before or 0) + Inches(gap)
    shift = Inches(-gap if bottom else gap) if gap < 0 else 0
    return strip_padding, shift


def validate_template(document):
    if len(document.sections) != 1:
        raise ValueError("A template must have one Word section. Put its content in the document body.")


def fill_template(document, values, number, record, settings, exam):
    qr = None
    tags = (*QR_TAGS, *values)
    pattern = re.compile("|".join(re.escape(tag) for tag in tags))
    rotated_cells = {}
    for paragraph in template_paragraphs(document):
        # Replace from right to left, so inserted text is never treated as another tag.
        original = "".join(run.text for run in paragraph.runs)
        standalone_qr = original.strip() in QR_TAGS
        has_qr = any(tag in original for tag in QR_TAGS) and settings["qr_fields"]
        strip_padding, vertical_shift = (align_qr(paragraph, settings, standalone=standalone_qr)
                                         if has_qr else (False, 0))
        rotated_alignment = "Template"
        for match in reversed(list(pattern.finditer(original))):
            tag = match.group()
            if tag in QR_TAGS:
                if qr is None:
                    qr = qr_png(number, record, settings, exam)
                replace_text(paragraph, tag, replacement="", picture=qr,
                             width=Inches(settings["qr_size"]), offset=match.start(), vertical_shift=vertical_shift)
            else:
                kind = "set" if tag in TEXT_TAGS[0] else "code"
                rotation = settings[f"{kind}_rotation"]
                if rotation:
                    side = settings[f"{kind}_horizontal"]
                    if side != "Template":
                        if rotated_alignment not in ("Template", side):
                            raise ValueError("Rotated tags in the same paragraph need the same horizontal alignment.")
                        rotated_alignment = side
                    strip_padding |= align_tag_paragraph(paragraph, side, standalone=original.strip() == tag)
                    cell_side = settings[f"{kind}_vertical"]
                    cells = paragraph._p.xpath("ancestor::w:tc[1]") if cell_side != "Template" else []
                    if cells:
                        if rotated_cells.setdefault(cells[0], cell_side) != cell_side:
                            raise ValueError("Rotated tags in the same cell need the same cell alignment.")
                replace_text(paragraph, tag, values[tag], offset=match.start(), template_text=True,
                             rotation=rotation, style_settings=settings)
        if strip_padding:
            # Leading tabs and trailing spaces consume the line's right-aligned width.
            for run in paragraph.runs:
                if run.text and not run.text.strip() and not run._r.xpath("./w:drawing"):
                    run.text = ""
    for cell, side in rotated_cells.items():
        align_tag_cell(cell, side)


def copy_default_run_format(document):
    # Word defaults belong to the source file. Materialize them before a template
    # becomes part of a different document.
    formats = document.styles.element.xpath("./w:docDefaults/w:rPrDefault/w:rPr")
    if not formats:
        return
    for paragraph in template_paragraphs(document):
        for run in paragraph.runs:
            props = run._r.get_or_add_rPr()
            for prop in formats[0]:
                if prop.tag in (qn("w:rFonts"), qn("w:sz"), qn("w:szCs")):
                    # A style may already provide a different font or size.
                    candidates = [paragraph.style, run.style]
                    if props.find(prop.tag) is None and not any(
                        style_has_format(style, prop.tag) for style in candidates
                    ):
                        props.append(deepcopy(prop))


def style_has_format(style, tag):
    while style is not None:
        properties = style.element.find(qn("w:rPr"))
        if properties is not None and properties.find(tag) is not None:
            return True
        style = style.base_style
    return False


def apply_page(document, settings):
    section = document.sections[0]
    width, height = (8.5, 14.0) if settings["paper"] == "Legal" else (8.2677, 11.6929)
    if settings["left"] + settings["right"] >= width - 1:
        raise ValueError("Left and right margins leave too little room for the question table.")
    if settings["top"] + settings["bottom"] >= height - 1:
        raise ValueError("Top and bottom margins leave too little room for the question table.")
    section.page_width, section.page_height = Inches(width), Inches(height)
    for side in ("top", "bottom", "left", "right"):
        setattr(section, side + "_margin", Inches(settings[side]))
    section.different_first_page_header_footer = False
    document.settings.odd_and_even_pages_header_footer = False
    section.footer_distance = Inches(min(0.2, settings["bottom"] / 3))
    return section


def fit_tables(document, available):
    """Scale the existing table grid only when the paper and margins require it."""
    for table in document.tables:
        widths = [column.width or 0 for column in table.columns]
        if not widths or sum(widths) <= available:
            continue
        factor = available / sum(widths)
        for column, width in zip(table.columns, widths):
            if width:
                column.width = Emu(round(width * factor))
        for row in table.rows:
            for cell in row.cells:
                if cell.width:
                    cell.width = Emu(round(cell.width * factor))
        preferred = table._tbl.tblPr.find(qn("w:tblW"))
        if preferred is not None and preferred.get(qn("w:type")) == "dxa":
            preferred.set(qn("w:w"), str(round(available / 635)))


def ensure_serial_width(table, count):
    # Fixed cell margins reduce the room available for consecutive digits.
    # Keep three-digit serials on one line even on narrower A4 pages.
    minimum = Inches(0.55 if count >= 100 else 0.45)
    current = table.columns[0].width
    second = table.columns[1].width
    if not current or not second or current >= minimum or second < Inches(1):
        return
    delta = minimum - current
    table.columns[0].width = minimum
    table.columns[1].width = Emu(second - delta)
    for row in table.rows:
        serial, question = row.cells
        serial.width = minimum
        question.width = Emu((question.width or second) - delta)


def apply_font(run, settings):
    run.font.name = settings["font"]
    run.font.size = Pt(settings["font_size"])
    properties = run._r.get_or_add_rPr()
    properties.rFonts.set(qn("w:cs"), settings["font"])
    size = OxmlElement("w:szCs")
    size.set(qn("w:val"), str(round(settings["font_size"] * 2)))
    properties.append(size)


def block_gap(paragraph, inches, keep=False):
    props = paragraph.paragraph_format
    props.space_before = Pt(0)
    props.space_after = Pt(0)
    props.line_spacing = Pt(inches * 72 if inches else 0.1)
    props.keep_with_next = keep


def footer_run(paragraph, text, settings, kind="footer"):
    run = paragraph.add_run(text)
    family, size, color, opacity = text_appearance(settings, kind)
    apply_font(run, {"font": family, "font_size": size})
    # Word run colors have no alpha channel. Blend against white paper.
    rgb = RGBColor.from_string(color.lstrip("#"))
    run.font.color.rgb = RGBColor(*(round(255 + (channel - 255) * opacity) for channel in rgb))
    if settings["footer_language"] == "Bangla":
        language = OxmlElement("w:lang")
        language.set(qn("w:val"), "bn-BD")
        language.set(qn("w:bidi"), "bn-BD")
        run._r.rPr.append(language)
    return run


def footer_field(paragraph, name, settings):
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), name)
    field.set(qn("w:dirty"), "true")
    field.append(footer_run(paragraph, "১" if settings["footer_language"] == "Bangla" else "1", settings)._r)
    paragraph._p.append(field)


def add_print_footer(section, record, number, settings, blank):
    footer = section.footer
    for element in list(footer._element):
        footer._element.remove(element)
    paragraph = footer.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(0)
    if blank:
        # Keep the footer line height when PDF uses an explicit Bengali numeral overlay.
        footer_run(paragraph, " ", settings)
        return
    width = section.page_width - section.left_margin - section.right_margin
    paragraph.paragraph_format.tab_stops.add_tab_stop(width, WD_TAB_ALIGNMENT.RIGHT)
    if settings["print_code_enabled"] and settings["print_code_opacity"] > 0:
        footer_run(paragraph, record["code"], settings, "print_code")
    footer_run(paragraph, "\t", settings)
    prefix, separator = footer_parts(number, settings)
    footer_run(paragraph, prefix, settings)
    footer_field(paragraph, "PAGE", settings)
    footer_run(paragraph, separator, settings)
    footer_field(paragraph, "NUMPAGES", settings)


def qr_content(number, record, settings, exam):
    lines = {
        "set": f"Set {set_number_text(number, settings['tag_language'])}", "post": f"Post: {post_name(exam)}",
        "post_code": f"Post Code: {exam['post_code']}", "unit": f"Unit: {exam['unit']}",
        "date": f"Exam Date: {exam['date']}", "timestamp": record["timestamp"],
        "booklet": f"Booklet Code: {record['code']}", "text": settings["credit"].strip(),
    }
    selected = [line for field, line in lines.items() if field in settings["qr_fields"] and line]
    return "\n".join(selected)


def qr_png(number, record, settings, exam):
    content = qr_content(number, record, settings, exam)
    if not content:
        return b""
    try:
        image = qrcode.make(content, box_size=8, border=4)
    except qrcode.exceptions.DataOverflowError as error:
        raise ValueError("The QR content is too long. Shorten the custom text or select fewer fields.") from error
    output = io.BytesIO()
    image.save(output, format="PNG")
    return output.getvalue()


def make_booklet(source, exam, number, settings, target, *, pdf_target=None, keep_docx=True):
    document = open_word(source)
    table, rows = question_rows(document, exam["count"])
    for row in list(table.rows):
        table._tbl.remove(row._tr)
    for position, question in enumerate(exam["mapping"][number], 1):
        row = rows[question]
        paragraph = next(p for p in row.cells[0].paragraphs if p.text.strip())
        raw = paragraph.text.strip()
        match = SERIAL.fullmatch(raw)
        if match is None:
            raise ValueError(f"Could not renumber the serial of question {question}.")
        digits = label(position) if any(c in "০১২৩৪৫৬৭৮৯" for c in match[1]) else str(position)
        suffix = settings["serial_suffix"]
        replacement = digits + (match[2] if suffix == "Original" else suffix)
        if raw != replacement and not replace_text(paragraph, raw, replacement):
            raise ValueError(f"Could not renumber the serial of question {question}.")
        serial_cell = row.cells[0]._tc
        cell_props = serial_cell.get_or_add_tcPr()
        if cell_props.find(qn("w:noWrap")) is None:
            cell_props.insert_element_before(OxmlElement("w:noWrap"), "w:tcMar", "w:textDirection", "w:vAlign")
        table._tbl.append(row._tr)

    # The two-column question table is the source of truth for the booklet body.
    body = document.element.body
    for element in list(body):
        if element is not table._tbl and element.tag != qn("w:sectPr"):
            body.remove(element)
    section = apply_page(document, settings)
    available = section.page_width - section.left_margin - section.right_margin
    fit_tables(document, available)
    ensure_serial_width(table, exam["count"])
    for existing in (section.header, section.first_page_header, section.even_page_header):
        if not existing.is_linked_to_previous:
            existing.is_linked_to_previous = True
    record = exam["records"][number]
    values = text_tag_values(number, exam["code_name"], settings["tag_language"])
    composer = Composer(document, preserve_styles=True)
    header_path = template_path(settings, "header")
    if header_path:
        header = open_word(header_path)
        validate_template(header)
        fit_tables(header, available)
        fill_template(header, values, number, record, settings, exam)
        copy_default_run_format(header)
        while len(header.element.body) > 1:
            last = header.element.body[-2]
            if last.tag != qn("w:p") or any(child.tag != qn("w:pPr") for child in last):
                break
            header.element.body.remove(last)
        gap = document.add_paragraph()
        block_gap(gap, settings["header_gap"], keep=True)
        table._tbl.addprevious(gap._p)
        composer.insert(0, header, remove_property_fields=False)
    footer_path = template_path(settings, "footer")
    if footer_path:
        footer = open_word(footer_path)
        validate_template(footer)
        fit_tables(footer, available)
        fill_template(footer, values, number, record, settings, exam)
        copy_default_run_format(footer)
        for paragraph in table.rows[-1]._tr.xpath(".//w:p"):
            Paragraph(paragraph, document._body).paragraph_format.keep_with_next = True
        gap = document.add_paragraph()
        block_gap(gap, settings["footer_gap"], keep=True)
        composer.append(footer, remove_property_fields=False)
    if keep_docx or pdf_target is None:
        add_print_footer(section, record, number, settings, blank=False)
        document.save(target)
    if pdf_target:
        add_print_footer(section, record, number, settings, blank=True)
        document.save(pdf_target)


def declared_fonts(path):
    document = open_word(path)
    names = set()
    for paragraph in template_paragraphs(document):
        for run in paragraph.runs:
            if not run.text.strip() or run._r.rPr is None:
                continue
            fonts = run._r.rPr.find(qn("w:rFonts"))
            if fonts is not None:
                key = "w:cs" if re.search(r"[\u0980-\u09ff]", run.text) else "w:ascii"
                preferred = fonts.get(qn(key)) or fonts.get(qn("w:hAnsi"))
                if preferred:
                    names.add(preferred)
    return names
