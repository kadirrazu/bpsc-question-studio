"""Build printable room files and update only the workbook's status cells."""

from collections import Counter
from contextlib import ExitStack
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from io import BytesIO
from pathlib import Path
from posixpath import normpath
import re
from tempfile import NamedTemporaryFile
from time import sleep
from zipfile import ZipFile

import fitz
from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from PySide6.QtCore import QLockFile, QPointF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QFontDatabase, QFontMetricsF, QPen

from .pdf_canvas import margin_painter
from .file_io import write_bytes

COVER_FIELDS = {
    "c_code": "Center Code", "c_name": "Center Name", "building": "Building",
    "floor": "Floor", "room_no": "Room Number", "no_of_cand": "Number of Candidates",
    "question_count": "Number of Questions",
}
ROOM_CODE_POSITIONS = ("Top Left", "Top Center", "Top Right", "Bottom Left", "Bottom Center", "Bottom Right")
ROOM_CODE_SCOPES = ("Every Question Page", "First Page of Each Question")
PAPER_SIZES = {"Legal": (612, 1008), "A4": (595.28, 841.89), "Letter": (612, 792)}
DEFAULT_COVER_TITLE = "Room Bundle Details"
DEFAULT_END_NOTE = "=== END OF THE ROOM FILE ==="
REQUIRED = {"c_code", "room_no", "no_of_cand", "question_count", "file_name", "instruction", "status"}
MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in "123456789¹²³"),
            *(f"LPT{i}" for i in "123456789¹²³")}


@dataclass(frozen=True)
class BundleOptions:
    paper: str = "Legal"
    cover: bool = False
    cover_blanks: int = 1
    cover_fields: tuple = tuple(COVER_FIELDS)
    end: bool = False
    end_blanks: int = 1
    end_note: str = DEFAULT_END_NOTE
    duplex: bool = True
    end_fields: tuple = ()
    cover_title: str = DEFAULT_COVER_TITLE
    cover_credit: bool = True
    cover_timestamp: bool = True
    end_timestamp: bool = True
    end_credit: bool = True
    footer_size: float = 9.0
    footer_opacity: float = 0.8
    cover_uppercase: bool = False
    end_uppercase: bool = False
    print_room_code: bool = False
    room_code_position: str = "Top Center"
    room_code_scope: str = "Every Question Page"
    room_code_format: str = "{code}"
    room_code_size: float = 8.0
    room_code_color: str = "#000000"
    room_code_opacity: float = 0.9


def component(value, label):
    text = str(value or "")
    if (not text or not text.strip() or re.search(r'[<>:"/\\|?*\x00-\x1f]', text)
            or text.endswith((".", " ")) or text.split(".", 1)[0].upper() in RESERVED):
        raise ValueError(f"{label} contains an invalid Windows file or folder name.")
    return text


def whole_count(value, label, minimum):
    if isinstance(value, bool) or not re.fullmatch(r"[0-9]+(?:\.0+)?", str(value or 0).strip()):
        raise ValueError(f"{label} must be a whole number.")
    number = int(str(value or 0).strip().split(".")[0])
    if value is None or number < minimum:
        raise ValueError(f"{label} must be at least {minimum}.")
    return number


def cell_text(cell):
    value = cell.value
    if value is None:
        return ""
    if isinstance(value, int) and re.fullmatch(r"0+", cell.number_format):
        return str(value).zfill(len(cell.number_format))
    return str(value)


def read_rooms(data, cover_fields, require_room_code=False):
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
    try:
        sheet = next((workbook[name] for name in workbook.sheetnames
                      if name.replace(" ", "").casefold() == "sheet1"), None)
        if sheet is None:
            raise ValueError("The workbook needs a Sheet1 or Sheet 1 worksheet.")
        rows = sheet.iter_rows()
        first = next(rows, ())
        names = [str(cell.value or "").strip().casefold() for cell in first]
        required = REQUIRED | ({"room_code"} if require_room_code else set())
        missing = required - set(names)
        if missing:
            raise ValueError("Missing worksheet columns: " + ", ".join(sorted(missing)))
        if any(names.count(name) > 1 for name in required | set(cover_fields)):
            raise ValueError("A required worksheet column appears more than once.")
        columns = {name: index for index, name in enumerate(names) if name}
        room_rows, errors, destinations, codes = [], [], set(), set()
        for index, row in enumerate(rows, 2):
            if cell_text(row[columns["instruction"]]).strip().casefold() != "do":
                continue
            try:
                def content(key):
                    return cell_text(row[columns[key]]) if key in columns else ""

                center = component(content("c_code"), "Center code")
                filename = component(content("file_name"), "File name")
                if not filename.lower().endswith(".pdf"):
                    filename += ".pdf"
                target = Path(center) / filename
                normalized = str(target).casefold()
                if normalized in destinations:
                    raise ValueError(f"Duplicate output path: {center}/{filename}")
                candidate_value = row[columns["no_of_cand"]].value
                candidates = whole_count(candidate_value, "no_of_cand", 0) if candidate_value not in (None, "") else None
                count = whole_count(row[columns["question_count"]].value, "question_count", 1)
                room_code = content("room_code").strip().upper() if require_room_code else ""
                if require_room_code:
                    if not re.fullmatch("[A-Z0-9]{8}", room_code):
                        raise ValueError("room_code needs exactly 8 letters or digits.")
                    if room_code in codes:
                        raise ValueError(f"Duplicate room_code: {room_code}.")
                    codes.add(room_code)
                destinations.add(normalized)
                room_rows.append((index, target, content("room_no"), candidates, count,
                                  {key: content(key) for key in cover_fields}, room_code))
            except ValueError as error:
                errors.append(f"Row {index}: {error}")
        if errors:
            raise ValueError("\n".join(errors) + "\nIf a formula result is blank, recalculate and save the workbook in Excel first.")
        return sheet.title, columns["status"] + 1, room_rows
    finally:
        workbook.close()


def worksheet_member(archive, sheet_name):
    sheets = etree.fromstring(archive.read("xl/workbook.xml")).findall(f"./{MAIN}sheets/{MAIN}sheet")
    identifier = next(sheet.get(REL + "id") for sheet in sheets if sheet.get("name") == sheet_name)
    relations = etree.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    target = next(relation.get("Target") for relation in relations if relation.get("Id") == identifier)
    return target.lstrip("/") if target.startswith("/") else normpath("xl/" + target)


def mark_done(data, sheet_name, column, row_number):
    result = BytesIO()
    with ZipFile(BytesIO(data)) as source, ZipFile(result, "w") as target:
        name = worksheet_member(source, sheet_name)
        sheet = etree.fromstring(source.read(name))
        row = sheet.find(f"./{MAIN}sheetData/{MAIN}row[@r='{row_number}']")
        address = f"{get_column_letter(column)}{row_number}"
        cell = row.find(f"{MAIN}c[@r='{address}']")
        if cell is None:
            cell = etree.Element(MAIN + "c", r=address)
            for existing in row.findall(MAIN + "c"):
                letters = re.match(r"[A-Z]+", existing.get("r", "")).group()
                if column_index_from_string(letters) > column:
                    existing.addprevious(cell)
                    break
            else:
                row.append(cell)
        for child in list(cell):
            if child.tag in {MAIN + "v", MAIN + "f", MAIN + "is"}:
                cell.remove(child)
        cell.set("t", "inlineStr")
        etree.SubElement(etree.SubElement(cell, MAIN + "is"), MAIN + "t").text = "Done"
        for member in source.infolist():
            content = etree.tostring(sheet, encoding="utf-8") if member.filename == name else source.read(member)
            target.writestr(member, content)
    return result.getvalue()


def draw_note_footer(painter, width, height, timestamp, credit, size, opacity):
    font = QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont)
    font.setPointSizeF(size)
    font.setBold(False)
    bold = QFont(font)
    bold.setBold(True)
    normal_metrics = QFontMetricsF(font, painter.device())
    bold_metrics = QFontMetricsF(bold, painter.device())
    line_height = max(normal_metrics.height(), bold_metrics.height()) + 4
    ascent = max(normal_metrics.ascent(), bold_metrics.ascent())
    lines = int(bool(timestamp)) + int(credit)
    rule_y = height - 26 - 8 - lines * line_height
    painter.setPen(QPen(QColor("black"), 0.6))
    painter.drawLine(QPointF(42, rule_y), QPointF(width - 42, rule_y))
    painter.save()
    painter.setOpacity(opacity)
    painter.setPen(QColor("black"))
    painter.setFont(font)
    baseline = rule_y + 8 + ascent
    if timestamp:
        painter.drawText(QPointF((width - normal_metrics.horizontalAdvance(timestamp)) / 2, baseline), timestamp)
        baseline += line_height
    if credit:
        prefix, team = "Software Developed by ", "IT Section, BPSC"
        prefix_width = normal_metrics.horizontalAdvance(prefix)
        x = (width - prefix_width - bold_metrics.horizontalAdvance(team)) / 2
        painter.drawText(QPointF(x, baseline), prefix)
        painter.setFont(bold)
        painter.drawText(QPointF(x + prefix_width, baseline), team)
    painter.restore()


def note_page(width, height, title, table, family, *, timestamp="", credit=False,
              footer_size=9.0, footer_opacity=0.8, uppercase=False):
    buffer, writer, painter = margin_painter(width, height)
    try:
        ink = QColor("black")
        heading_shade = QColor("#F2F2F2")
        label_shade = QColor("#F5F5F5")
        if uppercase:
            title = title.upper()
            if table:
                table = [(label.upper(), value.upper()) for label, value in table]
        footer_space = max(90, 52 + 3.2 * footer_size) if timestamp or credit else 48
        painter.setPen(ink)
        heading = QFont(family, 22)
        heading.setBold(True)
        painter.setFont(heading)
        word_wrap = Qt.TextFlag.TextWordWrap
        centered = Qt.AlignmentFlag.AlignCenter | word_wrap
        left, span = 42, width - 84
        heading_box = QRectF(left + 16, 0, span - 32, height)
        title_height = max(65, painter.boundingRect(heading_box, centered, title).height() + 24)
        space = height - 48 - footer_space
        if title_height > space:
            raise ValueError("The page heading or end note does not fit on one page.")
        title_y = 48 if table else 48 + (space - title_height) / 2
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(heading_shade)
        painter.drawRoundedRect(QRectF(left, title_y, span, title_height), 9, 9)
        painter.setPen(ink)
        painter.drawText(QRectF(left + 16, title_y, span - 32, title_height), centered, title)
        if table:
            font = QFont(family, 12)
            label_font = QFont(font)
            label_font.setBold(True)
            words = Qt.AlignmentFlag.AlignVCenter | word_wrap
            top, label_width = 48 + title_height + 19, 182
            for label, value in table:
                value_width = span - label_width
                painter.setFont(font)
                value_rect = QRectF(0, 0, value_width - 28, height)
                row_height = max(42, painter.boundingRect(value_rect, words, value).height() + 22)
                if top + row_height > height - footer_space:
                    raise ValueError("The selected fields and heading do not fit on one page.")
                painter.fillRect(QRectF(left, top, label_width, row_height), label_shade)
                painter.setBrush(Qt.BrushStyle.NoBrush)
                painter.setPen(QPen(ink, 0.6))
                painter.drawRect(QRectF(left, top, span, row_height))
                painter.drawLine(QPointF(left + label_width, top), QPointF(left + label_width, top + row_height))
                painter.setPen(ink)
                painter.setFont(label_font)
                painter.drawText(QRectF(left + 14, top, label_width - 28, row_height), words, label)
                painter.setFont(font)
                painter.drawText(QRectF(left + label_width + 14, top, value_width - 28, row_height), words, value)
                top += row_height
        if timestamp or credit:
            draw_note_footer(painter, width, height, timestamp, credit, footer_size, footer_opacity)
    finally:
        painter.end()
    del painter, writer
    buffer.close()
    return bytes(buffer.data())


def add_blanks(pdf, count, size):
    for _ in range(count):
        pdf.new_page(width=size[0], height=size[1])


def append_booklet(pdf, booklet, size):
    if all(abs(page.rect.width - size[0]) < 1 and abs(page.rect.height - size[1]) < 1 for page in booklet):
        pdf.insert_pdf(booklet, final=0)
        return
    for index in range(len(booklet)):
        page = pdf.new_page(width=size[0], height=size[1])
        if booklet[index].get_contents():
            page.show_pdf_page(page.rect, booklet, index, keep_proportion=True)


def room_code_layer(size, options, code):
    width, height = size
    buffer, writer, painter = margin_painter(width, height, 144)
    width, height = width * 2, height * 2
    try:
        font = QFont("Courier New")
        font.setStyleHint(QFont.StyleHint.Monospace)
        font.setPointSizeF(options.room_code_size)
        painter.setFont(font)
        color = QColor(options.room_code_color)
        color.setAlphaF(options.room_code_opacity)
        painter.setPen(color)
        vertical, horizontal = options.room_code_position.split()
        alignment = {"Left": Qt.AlignmentFlag.AlignLeft, "Center": Qt.AlignmentFlag.AlignHCenter,
                     "Right": Qt.AlignmentFlag.AlignRight}[horizontal]
        text = options.room_code_format.replace("{code}", code)
        metrics = painter.fontMetrics()
        if metrics.horizontalAdvance(text) > width - 96:
            raise ValueError("The room code text is too wide for the selected paper.")
        band = max(36, metrics.height() + 4)
        painter.drawText(QRectF(48, 16 if vertical == "Top" else height - band - 16, width - 96, band),
                         alignment | Qt.AlignmentFlag.AlignVCenter, text)
    finally:
        painter.end()
    del painter, writer
    buffer.close()
    return bytes(buffer.data())


def reserve_run_folder(output_base):
    dhaka = timezone(timedelta(hours=6))
    for attempt in range(5):
        started = datetime.now(dhaka)
        output = output_base / started.strftime("%Y-%m-%d-%H%M%S_CENTER_WISE_GENERATED_FILES")
        try:
            output.mkdir(parents=True)
        except FileExistsError:
            if attempt < 4:
                sleep(1 - started.microsecond / 1_000_000)
        else:
            return output, started
    raise ValueError("Could not reserve a new time-stamped output folder. Try again.")


def build_bundles(workbook_path, set_paths, output_base, options, progress, family="Arial"):
    workbook_path = workbook_path.resolve()
    if len(set_paths) == 0 or len(set(path.resolve() for path in set_paths)) != len(set_paths):
        raise ValueError("Select a different PDF for each set, starting with Set 1.")
    if options.end and not options.end_note.strip():
        raise ValueError("Enter the end note or turn it off.")
    if options.cover and not options.cover_title.strip():
        raise ValueError("Enter a cover title or turn the cover page off.")
    if options.cover and not options.cover_fields:
        raise ValueError("Select at least one field for the cover page.")
    if options.print_room_code and (
        options.room_code_position not in ROOM_CODE_POSITIONS or options.room_code_scope not in ROOM_CODE_SCOPES
        or "{code}" not in options.room_code_format
        or any(char in options.room_code_format for char in "\r\n")
    ):
        raise ValueError("Choose a room code placement and enter a single-line format containing {code}.")
    if options.print_room_code and (
        not 6 <= options.room_code_size <= 24 or not 0 <= options.room_code_opacity <= 1
        or not re.fullmatch(r"#[0-9A-Fa-f]{6}", options.room_code_color)
    ):
        raise ValueError("Use a room code font size from 6 to 24 pt, a #RRGGBB color and an opacity from 0 to 1.")
    size = PAPER_SIZES[options.paper]
    lock = QLockFile(str(workbook_path) + ".room-bundles.lock")
    if not lock.tryLock(0):
        raise ValueError("Another Room Print Packs job is using this workbook.")
    try:
        data = workbook_path.read_bytes()
        cover_fields = options.cover_fields if options.cover else ()
        end_fields = options.end_fields if options.end else ()
        fields = tuple(dict.fromkeys((*cover_fields, *end_fields)))
        sheet, status_column, rows = read_rooms(data, fields, options.print_room_code)
        if not rows:
            progress("No rows marked DO were found.")
            return None, 0, 0, 0
        source_files = {path.resolve() for path in set_paths}
        progress(f"{sheet} · {len(rows)} selected rooms · {len(set_paths)} source sets · {options.paper}")
        if options.cover and (1 + options.cover_blanks) % 2:
            progress("Cover and blank pages total an odd number. The first question paper will start on the back in duplex printing.")
        done = failed = copies = 0
        with ExitStack() as stack:
            booklets = [stack.enter_context(fitz.open(path)) for path in set_paths]
            for index, pdf in enumerate(booklets, 1):
                if not pdf.is_pdf or pdf.needs_pass or not len(pdf):
                    raise ValueError(f"Set {index} needs a readable, nonempty PDF without a password.")
            output, started = reserve_run_folder(output_base.resolve())
            progress(f"Output folder · {output}")
            timestamp = started.strftime("Generated on: %Y-%m-%d %H:%M:%S (UTC+06:00)")
            footer_style = {"footer_size": options.footer_size, "footer_opacity": options.footer_opacity}
            cover_footer = dict(footer_style, timestamp=timestamp if options.cover_timestamp else "",
                                credit=options.cover_credit, uppercase=options.cover_uppercase)
            end_footer = dict(footer_style, timestamp=timestamp if options.end_timestamp else "",
                              credit=options.end_credit, uppercase=options.end_uppercase)
            end = None
            if options.end and not end_fields:
                end = stack.enter_context(fitz.open(stream=note_page(*size, options.end_note, None, family, **end_footer),
                                                   filetype="pdf"))
            for index, room_path, room, candidates, count, details, room_code in rows:
                target = output / room_path
                resolved_target = target.resolve()
                if not resolved_target.is_relative_to(output) or resolved_target in source_files:
                    raise ValueError(f"The room output would overwrite a source PDF or leave the output folder: {target}")
                distribution = Counter(position % len(booklets) + 1 for position in range(count))
                order = ", ".join(f"Set {number} × {copies_for_set}" for number, copies_for_set in distribution.items())
                room_label = f"Room {room}" if room else f"Row {index}"
                candidate_count = candidates if candidates is not None else "—"
                warning = " · Fewer questions than candidates" if candidates is not None and count < candidates else ""
                prepared = None
                try:
                    if workbook_path.read_bytes() != data:
                        raise RuntimeError("The Excel workbook changed during generation. Close Excel and run again.")
                    target.parent.mkdir(parents=True, exist_ok=True)
                    with NamedTemporaryFile(dir=target.parent, suffix=".pdf", delete=False) as handle:
                        prepared = Path(handle.name)
                    with fitz.open() as pdf, ExitStack() as room_resources:
                        layer = None
                        if options.print_room_code and options.room_code_opacity > 0:
                            layer = room_resources.enter_context(fitz.open(
                                stream=room_code_layer(size, options, room_code), filetype="pdf"))
                        if options.cover:
                            table = [(COVER_FIELDS[key], details[key]) for key in options.cover_fields]
                            cover_page = note_page(*size, options.cover_title, table, family, **cover_footer)
                            with fitz.open(stream=cover_page, filetype="pdf") as cover:
                                pdf.insert_pdf(cover)
                            add_blanks(pdf, options.cover_blanks, size)
                        for position in range(count):
                            booklet = booklets[position % len(booklets)]
                            first_page = len(pdf)
                            append_booklet(pdf, booklet, size)
                            if layer is not None:
                                stop = len(pdf) if options.room_code_scope == ROOM_CODE_SCOPES[0] else first_page + 1
                                for page_number in range(first_page, stop):
                                    page = pdf[page_number]
                                    page.show_pdf_page(page.rect, layer, 0)
                            if options.duplex and len(booklet) % 2:
                                add_blanks(pdf, 1, size)
                        if options.end:
                            add_blanks(pdf, options.end_blanks, size)
                            if end is not None:
                                pdf.insert_pdf(end)
                            else:
                                table = [(COVER_FIELDS[key], details[key]) for key in end_fields]
                                with fitz.open(stream=note_page(*size, options.end_note, table, family, **end_footer),
                                               filetype="pdf") as last_page:
                                    pdf.insert_pdf(last_page)
                        pages = len(pdf)
                        pdf.save(prepared, garbage=3, deflate=True)
                    updated = mark_done(data, sheet, status_column, index)
                    if workbook_path.read_bytes() != data:
                        raise RuntimeError("The Excel workbook changed during generation. Close Excel and run again.")
                    prepared.replace(target)
                    try:
                        write_bytes(workbook_path, updated)
                    except OSError as error:
                        raise RuntimeError(f"PDF saved at {target}, but Excel could not be updated. Close it and run again.") from error
                    data = updated
                    done += 1
                    copies += count
                    unit = "question" if count == 1 else "questions"
                    progress(f"{room_label} · {count}/{count} {unit} assembled · {candidate_count} candidates"
                             f" · {order} · {pages} pages · {target.name}{warning}")
                except (OSError, ValueError, fitz.FileDataError) as error:
                    failed += 1
                    progress(f"{room_label} · Failed (row {index}) · {error}")
                finally:
                    if prepared:
                        prepared.unlink(missing_ok=True)
        return output, done, failed, copies
    finally:
        lock.unlock()
