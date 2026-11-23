from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from unittest.mock import patch
from zipfile import ZipFile

import fitz
from openpyxl import Workbook, load_workbook
from PySide6.QtWidgets import QApplication

from mcq.center_files import SOURCE_FIELDS, derive_center_rows, read_center_file, save_center_file
from mcq.room_bundles import BundleOptions, ROOM_CODE_SCOPES, build_bundles


class CenterFilesTest(TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.source = self.root / "center.xlsx"
        book = Workbook()
        sheet = book.active
        sheet.title = "Sheet 1"
        sheet.append((*SOURCE_FIELDS, "Manual Note"))
        sheet.append((101, "Center A", "Main", 1, 1001, 45, "Manual data"))
        sheet.append((102, "Center B", "Annex", 2, 1002, 2, "=1+1"))
        sheet.append((101, "Center A", "Main", 1, 2005, 3, "Keep this"))
        sheet["A2"].number_format = "0000"
        sheet["A4"].number_format = "0000"
        book.create_sheet("Notes")["A1"] = "Preserve this sheet"
        book.save(self.source)

    def derive(self, multiple=4, auto_do=True, include_code=True, separator="-"):
        original = self.source.read_bytes()
        name, rows = read_center_file(original)
        result = derive_center_rows(rows, multiple, ("room_serial", "c_code", "room_no"),
                                    separator, auto_do, include_code)
        output = self.root / "derived.xlsx"
        save_center_file(self.source, original, output, name, result, include_code)
        return original, result, output

    def test_derivation_preserves_original_and_existing_excel_cells(self):
        original, result, output = self.derive()
        self.assertEqual(result[2][:5], (1, 48, "1-0101-1001", "DO", ""))
        self.assertEqual(result[3][:3], (1, 4, "1-102-1002"))
        self.assertEqual(result[4][:3], (2, 4, "2-0101-2005"))
        codes = {values[-1] for values in result.values()}
        self.assertEqual(len(codes), 3)
        for code in codes:
            self.assertRegex(code, "^[A-Z0-9]{8}$")
        self.assertEqual(self.source.read_bytes(), original)
        with ZipFile(self.source) as before, ZipFile(output) as after:
            for path in before.namelist():
                if path != "xl/worksheets/sheet1.xml":
                    self.assertEqual(before.read(path), after.read(path), path)
        book = load_workbook(output)
        self.assertEqual(book.active["G2"].value, "Manual data")
        self.assertEqual(book.active["G3"].value, "=1+1")
        self.assertEqual(book.active["A2"].number_format, "0000")
        self.assertEqual(book.active["H2"].value, 1)
        self.assertEqual(book["Notes"]["A1"].value, "Preserve this sheet")
        _, exact, output = self.derive(1, False, False, "_")
        self.assertEqual(exact[2], (1, 45, "1_0101_1001", "", ""))
        self.assertNotIn("room_code", [cell.value for cell in load_workbook(output).active[1]])

    def test_room_pack_prints_code_only_on_selected_question_pages(self):
        workbook = load_workbook(self.source)
        workbook.active["F2"] = 1
        workbook.save(self.source)
        _, rows, derived = self.derive(1)
        question = self.root / "set.pdf"
        with fitz.open() as pdf:
            for page in range(2):
                pdf.new_page(width=612, height=1008).insert_text((72, 100), "Sample question")
            pdf.save(question)
        code = rows[2][-1]
        for scope, count in zip(ROOM_CODE_SCOPES, (2, 1)):
            output, done, failed, _ = build_bundles(
                derived, [question], self.root, BundleOptions(cover=True, end=True, print_room_code=True,
                    room_code_scope=scope, room_code_position="Bottom Right", room_code_format="Track {code}",
                    room_code_size=8.5, room_code_color="#204060", room_code_opacity=0.65),
                lambda _: None)
            self.assertEqual((done, failed), (3, 0))
            with fitz.open(output / "0101" / "1-0101-1001.pdf") as pdf:
                self.assertEqual(sum(f"Track {code}" in page.get_text() for page in pdf), count)
                self.assertNotIn(code, pdf[0].get_text())
                self.assertNotIn(code, pdf[-1].get_text())
                stamp = pdf[2].search_for(f"Track {code}")[0]
                self.assertGreater(stamp.y0, 970)
                self.assertGreater(stamp.x0, 400)
                styling = [span for span in pdf[2].get_texttrace() if abs(span["size"] - 8.5) < 0.01]
                self.assertTrue(styling)
                self.assertAlmostEqual(styling[0]["opacity"], 0.65, delta=0.01)
                for actual, channel in zip(styling[0]["color"], (32, 64, 96)):
                    self.assertAlmostEqual(actual, channel / 255, delta=0.01)
        self.assertEqual(load_workbook(derived).active["L2"].value, "Done")

    def test_cached_candidate_formula_drives_count_without_changing_the_formula(self):
        with_formula = BytesIO()
        member = "xl/worksheets/sheet1.xml"
        with ZipFile(self.source) as before, ZipFile(with_formula, "w") as after:
            for entry in before.infolist():
                body = before.read(entry)
                if entry.filename == member:
                    body = body.replace(b'<c r="F2" t="n"><v>45</v></c>',
                                        b'<c r="F2"><f>40+5</f><v>45</v></c>')
                after.writestr(entry, body)
        self.source.write_bytes(with_formula.getvalue())
        _, rows, output = self.derive()
        self.assertEqual(rows[2][1], 48)
        with output.open("rb") as stream:
            self.assertEqual(load_workbook(stream).active["F2"].value, "=40+5")
        with output.open("rb") as stream:
            self.assertEqual(load_workbook(stream, data_only=True).active["F2"].value, 45)

    def test_wizard_preview_changes_without_modifying_source(self):
        from mcq.center_wizard import CenterFileWizard

        wizard = CenterFileWizard()
        with patch("mcq.center_wizard.QFileDialog.getOpenFileName", return_value=(str(self.source), "")):
            wizard.open_file()
        original = self.source.read_bytes()
        self.assertEqual(wizard.preview.item(0, 5).text(), "45")
        wizard.count.setCurrentIndex(1)
        wizard.include_code.setChecked(True)
        wizard.auto_do.setChecked(True)
        self.assertEqual(wizard.preview.item(0, 5).text(), "48")
        self.assertEqual(wizard.preview.item(0, 7).text(), "DO")
        self.assertRegex(wizard.preview.item(0, 9).text(), "^[A-Z0-9]{8}$")
        with self.assertRaisesRegex(ValueError, "different output"):
            save_center_file(self.source, original, self.source, wizard.sheet_name, wizard.derived, True)
        self.assertEqual(self.source.read_bytes(), original)
