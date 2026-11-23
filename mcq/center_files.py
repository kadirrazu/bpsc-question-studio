"""Derive Room Print Packs columns without reserializing workbook source cells."""

from collections import Counter
from dataclasses import dataclass
from io import BytesIO
import re
import secrets
import string
from zipfile import ZipFile

from lxml import etree
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter, range_boundaries

from .file_io import write_bytes
from .room_bundles import MAIN, cell_text, component, whole_count, worksheet_member

SOURCE_FIELDS = ("c_code", "c_name", "building", "floor", "room_no", "no_of_cand")
OUTPUT_FIELDS = ("room_serial", "question_count", "file_name", "instruction", "status")
ROOM_CODE = "room_code"


@dataclass(frozen=True)
class CenterRow:
    excel_row: int
    values: dict
    candidates: int
    serial: int
    code: str


def read_center_file(data):
    workbook = load_workbook(BytesIO(data), read_only=True, data_only=True)
    try:
        sheet = next((workbook[name] for name in workbook.sheetnames
                      if name.replace(" ", "").casefold() == "sheet1"), None)
        if sheet is None:
            raise ValueError("The workbook needs a Sheet1 or Sheet 1 worksheet.")
        sheet.reset_dimensions()
        rows = sheet.iter_rows()
        headers = [str(cell.value or "").strip().casefold() for cell in next(rows, ())]
        invalid = [field for field in SOURCE_FIELDS if headers.count(field) != 1]
        if invalid:
            raise ValueError("These headers must appear once in row 1: " + ", ".join(invalid))
        collisions = set(headers) & {*OUTPUT_FIELDS, ROOM_CODE}
        if collisions:
            raise ValueError("The workbook already has derived columns: " + ", ".join(sorted(collisions)))
        positions = {field: headers.index(field) for field in SOURCE_FIELDS}
        serials = Counter()
        codes = set()
        source_rows = []
        for excel_row, row in enumerate(rows, 2):
            cells = {field: row[index] for field, index in positions.items() if index < len(row)}
            values = {field: cell_text(cells[field]) if field in cells else "" for field in SOURCE_FIELDS}
            if not any(values.values()):
                continue
            try:
                center = component(values["c_code"], "Center code").casefold()
                if not values["room_no"].strip():
                    raise ValueError("Room number is required.")
                count_cell = cells.get("no_of_cand")
                candidates = whole_count(count_cell.value if count_cell else None, "no_of_cand", 1)
            except ValueError as error:
                raise ValueError(f"Row {excel_row}: {error}") from error
            serials[center] += 1
            while True:
                code = "".join(secrets.choice(string.ascii_uppercase + string.digits) for _ in range(8))
                if code not in codes:
                    codes.add(code)
                    break
            source_rows.append(CenterRow(excel_row, values, candidates, serials[center], code))
        if not source_rows:
            raise ValueError("The workbook contains no room rows.")
        return sheet.title, source_rows
    finally:
        workbook.close()


def derive_center_rows(rows, multiple, name_fields, separator, auto_do, include_code):
    if multiple not in (1, 4):
        raise ValueError("Choose an exact count or the next multiple of 4.")
    if not name_fields:
        raise ValueError("Select at least one file-name component.")
    if any(field not in (*SOURCE_FIELDS, "room_serial", "question_count", ROOM_CODE) for field in name_fields):
        raise ValueError("Choose valid file-name components.")
    if ROOM_CODE in name_fields and not include_code:
        raise ValueError("Enable room codes before using them in a file name.")
    if re.search(r'[<>:"/\\|?*\x00-\x1f]', separator):
        raise ValueError("Choose a separator without Windows file-name punctuation.")
    largest_serial = Counter(row.values["c_code"].casefold() for row in rows)
    destinations = set()
    results = {}
    for row in rows:
        center = row.values["c_code"].casefold()
        width = len(str(largest_serial[center]))
        count = ((row.candidates + multiple - 1) // multiple) * multiple
        values = {**row.values, "room_serial": str(row.serial).zfill(width),
                  "question_count": str(count), ROOM_CODE: row.code}
        parts = [re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "-", values[field]).strip(" .")
                 for field in name_fields]
        try:
            filename = component(separator.join(part for part in parts if part), "File name")
        except ValueError as error:
            raise ValueError(f"Row {row.excel_row}: {error}") from error
        destination = (center, (filename if filename.lower().endswith(".pdf") else filename + ".pdf").casefold())
        if destination in destinations:
            raise ValueError(f"Row {row.excel_row}: duplicate file name in this center. Add Room Serial to the file name.")
        destinations.add(destination)
        results[row.excel_row] = (row.serial, count, filename, "DO" if auto_do else "", "",
                                  *((row.code,) if include_code else ()))
    return results


def save_center_file(source, expected, destination, sheet_name, derived, include_code):
    if source.resolve() == destination.resolve() or (destination.exists() and source.samefile(destination)):
        raise ValueError("Choose a different output file. The uploaded workbook stays unchanged.")
    if source.read_bytes() != expected:
        raise ValueError("The source workbook changed. Reload it to refresh the preview.")
    new_data = BytesIO()
    with ZipFile(BytesIO(expected)) as original, ZipFile(new_data, "w") as updated:
        member = worksheet_member(original, sheet_name)
        sheet = etree.fromstring(original.read(member))
        sheet_data = sheet.find(MAIN + "sheetData")
        positions = {int(row.get("r")): row for row in sheet_data.findall(MAIN + "row")}
        cells = [cell.get("r") for cell in sheet_data.iter(MAIN + "c") if cell.get("r")]
        dimension = sheet.find(MAIN + "dimension")
        previous_width = range_boundaries(dimension.get("ref"))[2] if dimension is not None else 0
        last_cell_column = max((column_index_from_string(re.match("[A-Z]+", cell).group())
                                for cell in cells), default=0)
        last_column = max(previous_width or 0, last_cell_column)
        fields = (*OUTPUT_FIELDS, *((ROOM_CODE,) if include_code else ()))
        if last_column + len(fields) > 16384:
            raise ValueError("The worksheet has no space for the derived columns.")
        for index, values in {1: fields, **derived}.items():
            row = positions[index]
            for column, value in enumerate(values, last_column + 1):
                if value == "":
                    continue
                cell = etree.SubElement(row, MAIN + "c", r=f"{get_column_letter(column)}{index}")
                if isinstance(value, int):
                    cell.set("t", "n")
                    etree.SubElement(cell, MAIN + "v").text = str(value)
                else:
                    cell.set("t", "inlineStr")
                    etree.SubElement(etree.SubElement(cell, MAIN + "is"), MAIN + "t").text = value
            row.attrib.pop("spans", None)
        if dimension is not None:
            dimension.set("ref", f"A1:{get_column_letter(last_column + len(fields))}{max(positions)}")
        for entry in original.infolist():
            body = etree.tostring(sheet, encoding="utf-8") if entry.filename == member else original.read(entry)
            updated.writestr(entry, body)
    write_bytes(destination, new_data.getvalue())
