from datetime import datetime, timedelta, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

import fitz
from openpyxl import Workbook, load_workbook
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QEventLoop, QTimer

from mcq.room_bundles import DEFAULT_END_NOTE, BundleOptions, build_bundles


class RoomBundleTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.application = QApplication.instance() or QApplication([])

    def setUp(self):
        temp = TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.workbook = self.root / "center.xlsx"
        workbook = Workbook()
        workbook.active.title = "Sheet 1"
        workbook.active.append(("c_code", "c_name", "room_no", "no_of_cand", "question_count",
                                "file_name", "instruction", "Status"))
        workbook.active.append(("007", "Test Center", 201, 11, 12, "Room 201", "Do"))
        workbook.active.append(("007", "Test Center", 202, 10, 12, "Room 202", "skip", "Pending"))
        other = workbook.create_sheet("Other")
        other["A1"] = "=SUM(1,2)"
        workbook.save(self.workbook)
        self.sources = []
        for index, size, page_count in ((1, (612, 1008), 2), (2, (612, 1008), 2),
                                         (3, (595.28, 841.89), 2), (4, (612, 1008), 2)):
            path = self.root / f"SET-{index}.pdf"
            with fitz.open() as document:
                for number in range(1, page_count + 1):
                    page = document.new_page(width=size[0], height=size[1])
                    page.insert_text((40, 65), f"SET {index} PAGE {number}")
                document.save(path)
            self.sources.append(path)

    def test_order_count_cover_end_and_regenerate(self):
        messages = []
        options = BundleOptions(cover=True, cover_fields=("c_code", "c_name", "room_no", "question_count", "building"),
                                end=True, end_blanks=2, end_note="Room closed.",
                                end_fields=("c_code", "room_no", "building"))
        output, done, failed, copies = build_bundles(
            self.workbook, self.sources, self.root, options, messages.append)
        self.assertEqual((done, failed, copies), (1, 0, 12))
        self.assertEqual(output.parent, self.root)
        self.assertRegex(output.name, r"^\d{4}-\d{2}-\d{2}-\d{6}_CENTER_WISE_GENERATED_FILES$")
        target = output / "007" / "Room 201.pdf"
        with fitz.open(target) as pdf:
            self.assertEqual(len(pdf), 29)
            self.assertTrue(all((page.rect.width, page.rect.height) == (612, 1008) for page in pdf))
            cover = pdf[0].get_text()
            for item in ("Room Bundle Details", "Center Name", "Test Center", "Number of Questions", "007", "201",
                         "12", "Building", "UTC+06:00", "Software Developed by IT Section, BPSC"):
                self.assertIn(item, cover)
            self.assertEqual(pdf[1].get_text(), "")
            actual = [page.get_text().strip() for page in pdf][2:26]
            self.assertEqual(actual, [f"SET {index} PAGE {page}" for index in range(1, 5) for page in (1, 2)] * 3)
            self.assertEqual(pdf[-3].get_text(), "")
            self.assertEqual(pdf[-2].get_text(), "")
            end_text = pdf[-1].get_text()
            self.assertEqual(end_text.split()[:9], "Room closed. Center Code 007 Room Number 201 Building".split())
            for page in (pdf[0], pdf[-1]):
                text = page.get_text()
                self.assertRegex(text, r"Generated on: \d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2} \(UTC\+06:00\)")
                self.assertIn("Software Developed by IT Section, BPSC", text)
                footer_blocks = [block for block in page.get_text("dict")["blocks"]
                                 if block["bbox"][1] > page.rect.height - 85]
                lines = [line for block in footer_blocks for line in block["lines"]]
                self.assertEqual(len(lines), 2)
                for line in lines:
                    bbox = line["bbox"]
                    self.assertAlmostEqual((bbox[0] + bbox[2]) / 2, page.rect.width / 2, delta=3)
                self.assertTrue(all(abs(span["size"] - 9) < 0.1 and abs(span["opacity"] - 0.8) < 0.02
                                    for span in page.get_texttrace() if span["bbox"][1] > page.rect.height - 85))
                credit = [span for span in lines[1]["spans"] if span["text"]]
                self.assertFalse(credit[0]["flags"] & fitz.TEXT_FONT_BOLD)
                self.assertTrue(credit[-1]["flags"] & fitz.TEXT_FONT_BOLD)
                drawings = page.get_drawings()
                fills = [d for d in drawings if d.get("fill")]
                self.assertTrue(any(d["rect"].width > 500 and 0.93 < d["fill"][0] < 0.96 for d in fills))
                self.assertTrue(any(abs(d["rect"].width - 182) < 1 and 0.95 < d["fill"][0] < 0.98 for d in fills))
                self.assertTrue(all(d["fill"][0] > 0.93 for d in fills))
                for drawing in drawings:
                    for key in ("color", "fill"):
                        rgb = drawing.get(key)
                        if rgb:
                            self.assertAlmostEqual(max(rgb), min(rgb), delta=0.002)
                self.assertTrue(any(d["rect"].width > 500 and d["rect"].y0 > page.rect.height - 90
                                    for d in drawings))
        book = load_workbook(self.workbook)
        self.assertEqual(book["Sheet 1"]["H2"].value, "Done")
        self.assertEqual(book["Sheet 1"]["H3"].value, "Pending")
        self.assertEqual(book["Other"]["A1"].value, "=SUM(1,2)")
        self.assertFalse((output / "007" / "Room 202.pdf").exists())
        self.assertTrue(any("Room 201 · 12/12 questions assembled · 11 candidates · Set 1 × 3" in line
                            for line in messages))
        result = build_bundles(self.workbook, self.sources, self.root, BundleOptions(paper="A4"), messages.append)
        self.assertEqual(result[1:], (1, 0, 12))
        self.assertNotEqual(output, result[0])
        with fitz.open(target) as pdf:
            self.assertEqual(len(pdf), 29)
        with fitz.open(result[0] / "007" / "Room 201.pdf") as pdf:
            self.assertEqual(len(pdf), 24)
            self.assertAlmostEqual(pdf[0].rect.width, 595.28, delta=0.5)
        custom = BundleOptions(cover=True, cover_title="Custom Cover", cover_credit=False, end=True,
                               cover_fields=("no_of_cand",), end_timestamp=False, end_credit=True,
                               footer_size=11.0, footer_opacity=0.55)
        book = load_workbook(self.workbook)
        book.active["D2"] = None
        book.save(self.workbook)
        custom_output, *_ = build_bundles(self.workbook, self.sources, self.root, custom, messages.append)
        with fitz.open(custom_output / "007" / "Room 201.pdf") as pdf:
            self.assertEqual(len(pdf), 28)
            self.assertIn("Custom Cover", pdf[0].get_text())
            self.assertIn("Number of Candidates", pdf[0].get_text())
            self.assertNotIn("11", pdf[0].get_text(clip=fitz.Rect(223, 132, 568, 177)))
            self.assertIn("UTC+06:00", pdf[0].get_text())
            self.assertNotIn("Software Developed", pdf[0].get_text())
            self.assertFalse(pdf[-2].get_text())
            self.assertNotIn("UTC+06:00", pdf[-1].get_text())
            self.assertIn(DEFAULT_END_NOTE, pdf[-1].get_text())
            self.assertTrue(any(d.get("fill") and 0.93 < d["fill"][0] < 0.96
                                for d in pdf[-1].get_drawings()))
            self.assertIn("Software Developed by IT Section, BPSC", pdf[-1].get_text())
            self.assertTrue(any(abs(span["size"] - 11) < 0.1 and abs(span["opacity"] - 0.55) < 0.02
                                for span in pdf[-1].get_texttrace()))
        for cover_on, end_on in ((True, False), (False, True)):
            with self.subTest(cover=cover_on, end=end_on):
                options = BundleOptions(cover=cover_on, end=end_on)
                partial_output, *_ = build_bundles(self.workbook, self.sources, self.root, options, messages.append)
                with fitz.open(partial_output / "007" / "Room 201.pdf") as pdf:
                    self.assertEqual(len(pdf), 26)
                    self.assertEqual("Room Bundle Details" in pdf[0].get_text(), cover_on)
                    self.assertEqual(DEFAULT_END_NOTE in pdf[-1].get_text(), end_on)

    def test_cover_and_end_uppercase_are_independent(self):
        for cover_uppercase, end_uppercase in ((True, False), (False, True)):
            with self.subTest(cover=cover_uppercase, end=end_uppercase):
                options = BundleOptions(cover=True, cover_fields=("c_name",), end=True,
                                        end_note="Room closed.", end_fields=("c_name",),
                                        cover_uppercase=cover_uppercase, end_uppercase=end_uppercase)
                output, done, failed, _ = build_bundles(
                    self.workbook, self.sources, self.root, options, lambda _: None)
                self.assertEqual((done, failed), (1, 0))
                with fitz.open(output / "007" / "Room 201.pdf") as pdf:
                    for page, upper, title in ((pdf[0], cover_uppercase, "Room Bundle Details"),
                                               (pdf[-1], end_uppercase, "Room closed.")):
                        text = page.get_text()
                        for value in (title, "Center Name", "Test Center"):
                            self.assertIn(value.upper() if upper else value, text)
                        self.assertIn("Generated on:", text)

    def test_bad_rows_leave_workbook_and_output_untouched(self):
        workbook = load_workbook(self.workbook)
        sheet = workbook.active
        sheet["I1"] = "Another Column"
        sheet["F2"] = "../illegal"
        workbook.save(self.workbook)
        original = self.workbook.read_bytes()
        with self.assertRaisesRegex(ValueError, "Row 2"):
            build_bundles(self.workbook, self.sources, self.root, BundleOptions(), lambda _: None)
        self.assertEqual(self.workbook.read_bytes(), original)
        self.assertFalse(list(self.root.glob("*_CENTER_WISE_GENERATED_FILES")))

    def test_same_second_regeneration_reserves_a_fresh_folder(self):
        dhaka = timezone(timedelta(hours=6))
        first = datetime(2026, 11, 21, 15, 30, 45, 500000, tzinfo=dhaka)
        next_second = first.replace(second=46, microsecond=0)
        options = BundleOptions(cover=True, end=True)
        with patch("mcq.room_bundles.datetime") as clock, patch("mcq.room_bundles.sleep") as pause:
            clock.now.side_effect = [first, first, next_second]
            a, *_ = build_bundles(self.workbook, self.sources, self.root, options, lambda _: None)
            original = (a / "007" / "Room 201.pdf").read_bytes()
            b, *_ = build_bundles(self.workbook, self.sources, self.root, options, lambda _: None)
        pause.assert_called_once_with(0.5)
        self.assertEqual(clock.now.call_args.args, (dhaka,))
        self.assertEqual(a.name, "2026-11-21-153045_CENTER_WISE_GENERATED_FILES")
        self.assertEqual(b.name, "2026-11-21-153046_CENTER_WISE_GENERATED_FILES")
        self.assertEqual((a / "007" / "Room 201.pdf").read_bytes(), original)
        with fitz.open(b / "007" / "Room 201.pdf") as pdf:
            for page in (pdf[0], pdf[-1]):
                self.assertIn("Generated on: 2026-11-21 15:30:46 (UTC+06:00)", page.get_text())

    def test_odd_booklet_padding_and_background_progress(self):
        from mcq.room_dialog import RoomBundlesDialog

        workbook = load_workbook(self.workbook)
        workbook.active["E2"] = 30
        workbook.save(self.workbook)
        odd = self.sources[0]
        with fitz.open() as pdf:
            page = pdf.new_page(width=612, height=1008)
            page.insert_text((40, 65), "ODD PAGE")
            pdf.save(odd)
        dialog = RoomBundlesDialog(self.root)
        dialog.workbook.setText(str(self.workbook))
        dialog.output_base.setText(str(self.root))
        dialog.set_count.setValue(1)
        dialog.set_fields[0][1].setText(str(odd))
        dialog.start()
        loop = QEventLoop()
        dialog.task.finished.connect(loop.quit)
        QTimer.singleShot(15000, loop.quit)
        loop.exec()
        self.assertIsNone(dialog.task)
        console = dialog.console.toPlainText()
        self.assertIn("COMPLETED · 1 rooms saved  |  0 failed  |  30 questions", console)
        self.assertEqual(console.count("questions assembled"), 1)
        self.assertIn("Room 201 · 30/30 questions assembled", console)
        self.assertRegex(console.splitlines()[-1], r"Total time · \d+m \d{2}\.\d{2}s")
        self.assertEqual(dialog.status.property("state"), "success")
        self.assertEqual(dialog.generate.text(), "Regenerate Room Files")
        self.assertTrue(dialog.open_folder.isEnabled())
        self.assertIn(str(dialog.output), console)
        with fitz.open(dialog.output / "007" / "Room 201.pdf") as pdf:
            self.assertEqual(len(pdf), 60)
            self.assertIn("ODD PAGE", pdf[0].get_text())
            self.assertTrue(all(not pdf[number].get_text() for number in range(1, 60, 2)))

    def test_room_defaults_are_independent_of_set_generation(self):
        from mcq.room_dialog import RoomBundlesDialog

        config = self.root / "Config"
        config.mkdir()
        shared = config / "settings.json"
        shared.write_text('{"paper":"A4","output_root":"old-exams","font":"Nikosh"}')
        before = shared.read_bytes()
        dialog = RoomBundlesDialog(self.root)
        self.assertEqual(dialog.paper.currentText(), "Legal")
        self.assertFalse(dialog.cover.isChecked())
        self.assertFalse(dialog.end.isChecked())
        self.assertFalse(dialog.cover_uppercase.isChecked())
        self.assertFalse(dialog.end_uppercase.isChecked())
        self.assertEqual(dialog.end_note.toPlainText(), DEFAULT_END_NOTE)
        self.assertEqual(dialog.footer_size.value(), 9)
        self.assertEqual(dialog.footer_opacity.value(), 0.8)
        self.assertFalse(dialog.set_fields[0][1].text())
        dialog.workbook.setText(str(self.workbook))
        dialog.paper.setCurrentText("Letter")
        dialog.cover.setChecked(True)
        dialog.cover_blanks.setValue(3)
        dialog.end.setChecked(True)
        dialog.end_note.setPlainText("Last page")
        self.assertEqual(dialog.end_blanks.value(), 1)
        dialog.cover_title.setText("My Cover")
        dialog.cover_credit.setChecked(False)
        dialog.cover_uppercase.setChecked(True)
        dialog.end_uppercase.setChecked(True)
        dialog.end_credit.setChecked(False)
        dialog.end_timestamp.setChecked(False)
        dialog.footer_size.setValue(11.5)
        dialog.footer_opacity.setValue(0.65)
        dialog.end_fields["room_no"].setChecked(True)
        dialog.set_fields[0][1].setText(str(self.sources[0]))
        dialog.reject()
        reopened = RoomBundlesDialog(self.root)
        self.assertEqual(reopened.options(), dialog.options())
        self.assertEqual(reopened.workbook.text(), str(self.workbook))
        self.assertEqual(reopened.set_fields[0][1].text(), str(self.sources[0]))
        self.assertEqual(shared.read_bytes(), before)

    def test_open_exam_prefills_generated_pdfs_in_order(self):
        from shutil import copy2
        from mcq.ui import Window

        window = Window(self.root)
        window.folder = self.root / "exam"
        folder = window.folder / "GENERATED_SETS"
        folder.mkdir(parents=True)
        copy2(self.sources[0], folder / "first.pdf")
        copy2(self.sources[2], folder / "third.PDF")
        window.exam = {"sets": 3, "generated": {
            "1": ["first.docx", "first.pdf"], "2": ["missing.pdf"], "3": ["third.PDF"],
        }}
        with patch("mcq.room_dialog.RoomBundlesDialog") as dialog:
            window.room_bundles()
            self.assertEqual(dialog.call_args.kwargs["set_files"], [
                str(folder / "first.pdf"), "", str(folder / "third.PDF")])
            window.exam = None
            window.room_bundles()
            self.assertIsNone(dialog.call_args.kwargs["set_files"])

    def test_numeric_and_choice_fields_ignore_wheel(self):
        from PySide6.QtCore import QPoint, QPointF, Qt
        from PySide6.QtGui import QWheelEvent
        from mcq.controls import Choice, CountInput, DateInput, DecimalInput

        choice = Choice()
        choice.addItems(["Legal", "A4"])
        choice.setCurrentIndex(1)
        count = CountInput()
        count.setValue(4)
        decimal = DecimalInput()
        decimal.setValue(0.75)
        date = DateInput()
        for control, read_value in ((choice, choice.currentIndex), (count, count.value),
                                    (decimal, decimal.value), (date, date.date)):
            before = read_value()
            control.show()
            control.setFocus()
            for delta in (120, -120):
                event = QWheelEvent(QPointF(5, 5), QPointF(5, 5), QPoint(), QPoint(0, delta),
                                    Qt.MouseButton.NoButton, Qt.KeyboardModifier.NoModifier,
                                    Qt.ScrollPhase.NoScrollPhase, False)
                QApplication.sendEvent(control, event)
                self.assertEqual(read_value(), before)
            control.close()
