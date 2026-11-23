"""Print-friendly Excel and PDF copies of a verified question mapping."""

from math import ceil

from PySide6.QtCore import QMarginsF, QRectF, Qt
from PySide6.QtGui import QColor, QFont, QPageLayout, QPageSize, QPainter, QPdfWriter
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from .project import mapping_rows, post_name

PAPER = {"Legal": QPageSize.PageSizeId.Legal, "A4": QPageSize.PageSizeId.A4,
         "Letter": QPageSize.PageSizeId.Letter}
INK = "#172D4A"
ACCENT = "#2563A3"
GRID = "#D5E0EF"
PALE = "#F1F6FC"
CHECK = "Verified: each set column contains every number from 1 to {count} exactly once."
CREDIT = "Software Developed by "
CREDIT_TEAM = "IT Section, BPSC"
TABLE_TOP = 35
TITLE = "Question Mapping"


def labels(exam):
    return [
        (f"Post Code: {exam.get('post_code') or '—'}", f"Unit: {exam['unit']}"),
        (f"Post Name: {post_name(exam)}", f"Exam Date: {exam['date']}"),
    ]


def panels(exam, settings):
    if settings["mapping_paper"] == "Legal":
        return [(1, exam["sets"] + 1)]
    width = 7 if settings["mapping_orientation"] == "Portrait" else 11
    return [(start, min(start + width, exam["sets"] + 1))
            for start in range(1, exam["sets"] + 1, width)]


def verified_rows(exam):
    rows = mapping_rows(exam)
    expected = set(range(1, exam["count"] + 1))
    if (len(rows) != len(expected)
            or any(len(row) != exam["sets"] + 1 for row in rows)
            or any(any(type(value) is not int for value in column) or set(column) != expected
                   for column in zip(*rows))):
        raise ValueError("Every mapping column must contain each question number exactly once.")
    return rows


def export_mapping_xlsx(exam, target, settings):
    workbook = Workbook()
    workbook.remove(workbook.active)
    rows = verified_rows(exam)
    thin = Side(style="thin", color=GRID[1:])
    for start, end in panels(exam, settings):
        sheet = workbook.create_sheet("Question Mapping" if start == 1 else f"Sets {start}–{end-1}")
        sheet.sheet_view.showGridLines = False
        columns = end - start + 1
        # The title needs a middle cell even when the mapping has only two columns.
        display_columns = max(3, columns)
        left_end = max(1, min(display_columns - 2, display_columns // 3))
        right_start = display_columns - left_end + 1

        def header(row, first, last, text, align, *, title=False):
            if first < last:
                sheet.merge_cells(start_row=row, start_column=first, end_row=row, end_column=last)
            cell = sheet.cell(row, first, text)
            cell.font = Font(name=settings["font"], color=INK[1:], size=12 if title else 10, bold=title)
            cell.alignment = Alignment(horizontal=align, vertical="center", shrink_to_fit=True)

        metadata = labels(exam)
        header(1, 1, left_end, metadata[0][0], "left")
        header(1, left_end + 1, right_start - 1, TITLE, "center", title=True)
        header(1, right_start, display_columns, metadata[0][1], "right")
        header(2, 1, right_start - 1, metadata[1][0], "left")
        header(2, right_start, display_columns, metadata[1][1], "right")
        sheet.row_dimensions[1].height = 18
        sheet.row_dimensions[2].height = 15

        lines = [["Set 0", *(f"Set {i}" for i in range(start, end))]]
        lines.extend([row[0], *row[start:end]] for row in rows)
        row_height = 16 if settings["mapping_paper"] == "Legal" else 19
        for index, line in enumerate(lines, 3):
            sheet.row_dimensions[index].height = 18 if index == 3 else row_height
            for col, value in enumerate(line, 1):
                if columns == 2 and col == 2:
                    sheet.merge_cells(start_row=index, start_column=2, end_row=index, end_column=3)
                cell = sheet.cell(index, col, value)
                cell.font = Font(name=settings["font"], size=8 if columns > 12 else 10,
                                 bold=index == 3 or col == 1,
                                 color="FFFFFF" if index == 3 else INK[1:])
                cell.fill = PatternFill("solid", fgColor=(ACCENT if index == 3 else
                                        PALE if index % 2 or col == 1 else "#FFFFFF")[1:])
                cell.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                cell.alignment = Alignment(horizontal="center", vertical="center", shrink_to_fit=True)
        column_width = max(4 if settings["mapping_paper"] == "Legal" else 12,
                           min(38, 96 / display_columns))
        for column in range(1, display_columns + 1):
            sheet.column_dimensions[get_column_letter(column)].width = column_width
        note = sheet.max_row + 1
        sheet.merge_cells(start_row=note, start_column=1, end_row=note, end_column=display_columns)
        cell = sheet.cell(note, 1, CHECK.format(count=len(rows)))
        cell.font = Font(name=settings["font"], size=8, color="166047")
        cell.alignment = Alignment(vertical="center", shrink_to_fit=True, indent=1)
        sheet.row_dimensions[note].height = 15

        sheet.freeze_panes = "B4"
        sheet.print_title_rows = "1:3"
        sheet.print_options.horizontalCentered = True
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 2 if settings["mapping_paper"] == "Legal" and 50 < len(rows) <= 100 else 0
        sheet.page_setup.orientation = settings["mapping_orientation"].lower()
        sheet.page_setup.paperSize = {"Legal": sheet.PAPERSIZE_LEGAL, "A4": sheet.PAPERSIZE_A4,
                                      "Letter": sheet.PAPERSIZE_LETTER}[settings["mapping_paper"]]
        sheet.page_margins.left = sheet.page_margins.right = 0.5
        sheet.page_margins.top = sheet.page_margins.bottom = 0.5
        sheet.page_margins.footer = 0.24
        sheet.oddFooter.left.text = CREDIT + "&B" + CREDIT_TEAM + "&B"
        sheet.oddFooter.left.size = 8
        sheet.oddFooter.right.text = "Page &P of &N"
        sheet.oddFooter.right.size = 8
    workbook.save(target)


def export_mapping_pdf(exam, target, settings):
    rows = verified_rows(exam)
    writer = QPdfWriter(str(target))
    writer.setTitle("Question Mapping")
    writer.setResolution(144)
    writer.setPageSize(QPageSize(PAPER[settings["mapping_paper"]]))
    direction = QPageLayout.Orientation
    writer.setPageOrientation(direction.Portrait if settings["mapping_orientation"] == "Portrait" else direction.Landscape)
    left, top, right, bottom = [float(settings.get(f"mapping_margin_{side}", 0.5)) * 72
                               for side in ("left", "top", "right", "bottom")]
    bounds = writer.pageLayout().fullRect(QPageLayout.Unit.Point)
    if bounds.width() - left - right < 160 or bounds.height() - top - bottom < 180:
        raise ValueError("Mapping PDF margins leave too little room for the report.")
    if not writer.setPageMargins(QMarginsF(left, top, right, bottom), QPageLayout.Unit.Point):
        raise ValueError("The mapping PDF margins do not fit this page size.")
    painter = QPainter(writer)
    if not painter.isActive():
        raise ValueError("Could not create the mapping PDF.")
    try:
        painter.scale(2, 2)
        width, height = writer.width() / 2, writer.height() / 2
        row_height = 18
        available = height - TABLE_TOP - 39
        if settings["mapping_paper"] == "Legal":
            row_height = max(8, min(16, available / (1 + ceil(min(len(rows), 100) / 2))))
        capacity = max(1, int(available / row_height + 1e-7) - 1)
        if settings["mapping_paper"] == "Legal" and ceil(len(rows) / 2) <= capacity < len(rows) <= 100:
            capacity = ceil(len(rows) / 2)
        pages = [(range(start, end), rows[offset:offset + capacity], offset + capacity >= len(rows))
                 for start, end in panels(exam, settings) for offset in range(0, len(rows), capacity)]
        for index, (columns, page_rows, last_in_panel) in enumerate(pages, 1):
            if index > 1:
                writer.newPage()
            paint_mapping_page(painter, exam, settings, width, height, columns, page_rows,
                               index, len(pages), len(rows) if last_in_panel else None, row_height)
    finally:
        painter.end()


def paint_mapping_page(painter, exam, settings, width, height, columns, rows, page, pages, verified_count, row_height):
    def font(size, bold=False):
        value = QFont(settings["font"])
        value.setPixelSize(size)
        value.setBold(bold)
        painter.setFont(value)

    def header(box, alignment, text, size=10, bold=False):
        font(size, bold)
        while size > 7 and painter.fontMetrics().horizontalAdvance(text) > box.width():
            size -= 1
            font(size, bold)
        painter.drawText(box, alignment | Qt.AlignmentFlag.AlignVCenter,
                         painter.fontMetrics().elidedText(text, Qt.TextElideMode.ElideRight, int(box.width())))

    painter.setPen(QColor(INK))
    top_line, second_line = labels(exam)
    header(QRectF(0, 0, width * .35, 16), Qt.AlignmentFlag.AlignLeft, top_line[0])
    header(QRectF(width * .36, 0, width * .28, 16), Qt.AlignmentFlag.AlignHCenter, TITLE, 12, True)
    header(QRectF(width * .65, 0, width * .35, 16), Qt.AlignmentFlag.AlignRight, top_line[1])
    header(QRectF(0, 16, width * .64, 15), Qt.AlignmentFlag.AlignLeft, second_line[0])
    header(QRectF(width * .65, 16, width * .35, 15), Qt.AlignmentFlag.AlignRight, second_line[1])
    headers = ["Set 0", *(f"Set {i}" for i in columns)]
    x_width = width / len(headers)
    for index, line in enumerate([headers, *([row[0], *(row[i] for i in columns)] for row in rows)]):
        for column, value in enumerate(line):
            cell = QRectF(column * x_width, TABLE_TOP + row_height * index, x_width, row_height)
            painter.fillRect(cell, QColor(ACCENT if not index else PALE if index % 2 or not column else "#FFFFFF"))
            painter.setPen(QColor(GRID))
            painter.drawRect(cell)
            painter.setPen(QColor(INK if index else "#FFFFFF"))
            size = min(10, max(6, int(row_height * .7)))
            bold = not index or not column
            font(size, bold)
            text = str(value)
            while size > 4 and painter.fontMetrics().horizontalAdvance(text) > x_width - 2:
                size -= 1
                font(size, bold)
            painter.drawText(cell, Qt.AlignmentFlag.AlignCenter, text)
    if verified_count:
        font(8)
        painter.setPen(QColor("#166047"))
        painter.drawText(QRectF(0, TABLE_TOP + row_height * (len(rows) + 1) + 3, width, 15),
                         Qt.AlignmentFlag.AlignVCenter, CHECK.format(count=verified_count))
    font(8)
    painter.setPen(QColor("#4C6079"))
    y = round(height - 4)
    painter.drawText(0, y, CREDIT)
    offset = painter.fontMetrics().horizontalAdvance(CREDIT)
    font(8, True)
    painter.drawText(offset, y, CREDIT_TEAM)
    font(8)
    paging = f"Page {page} of {pages}"
    painter.drawText(round(width - painter.fontMetrics().horizontalAdvance(paging)), y, paging)
