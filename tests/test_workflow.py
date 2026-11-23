"""Run with python -m unittest discover -s tests -v."""

import os
import json
import subprocess
from hashlib import sha256
import unittest
from copy import deepcopy
from datetime import date
from unittest.mock import patch
from pathlib import Path
from tempfile import TemporaryDirectory

import fitz
from docx import Document
from openpyxl import load_workbook
from PySide6.QtWidgets import QApplication

from mcq.access import initialize_access, matches
from mcq.documents import make_booklet, qr_content, replace_text
from mcq.project import DEFAULTS, create_exam, delete_exam, label, load_units, make_mapping, mapping_rows, read_exam, save_exam, save_json, validate_mapping
from mcq.mapping import CHECK, verified_rows
from mcq.workflow import (
    confirm_zero, digest, export_booklets, generate_mapping, generated_complete, generated_numbers, import_source,
    mapping_files, review_signature, review_status, save_mapping_exports, save_template,
)

APP = Path(__file__).resolve().parents[1]
RENDERER = os.environ.get("MCQ_TEST_RENDERER", "")


class WorkflowTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = QApplication.instance() or QApplication([])

    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.output = self.root / "EXAMS"
        self.folder, self.exam = create_exam(
            self.output, "Unit 10", "260010", "Assistant Registrar", "ইছামতি",
            3, 4, "2026-11-19",
        )
        self.exam = generate_mapping(self.output, self.folder, self.exam, DEFAULTS)

    def import_demo(self):
        self.exam = import_source(self.folder, self.exam, APP / "examples" / "DEMO_ZERO_SET_3_QUESTIONS.docx")

    def settings(self):
        header = save_template(self.root, "header", APP / "examples" / "HEADER_TEMPLATE.docx")
        return {**DEFAULTS, "header": header, "renderer": RENDERER, "font": "Nikosh"}

    def test_login_prefills_an_editable_email(self):
        from mcq.ui import LoginDialog

        credentials = initialize_access(self.root)
        dialog = LoginDialog(credentials)
        self.assertEqual(dialog.email.text(), "kadir@bpsc.gov.bd")
        self.assertFalse(dialog.email.isReadOnly())
        credentials.write_text("email=operator@example.org\npassword=test-password\n", encoding="utf-8")
        dialog.email.setText("operator@example.org")
        dialog.password.setText("test-password")
        dialog.sign_in()
        self.assertEqual(dialog.result(), dialog.DialogCode.Accepted)

    def test_writer_timeout_is_reported_and_retry_uses_fresh_profile(self):
        from mcq.pdf import render_batch

        commands = []
        source = self.root / "question.docx"
        out = self.root / "pdf"
        out.mkdir()

        def launch(command, **options):
            from unittest.mock import Mock
            commands.append(command)
            self.assertEqual(options["stdin"], subprocess.DEVNULL)
            proc = Mock()
            if len(commands) == 1:
                proc.poll.return_value = None
                proc.wait.side_effect = [subprocess.TimeoutExpired(command, 80), -9]
            else:
                render_folder = Path(command[command.index("--outdir") + 1])
                (render_folder / "question.pdf").write_bytes(b"example output")
                proc.wait.return_value = 0
            return proc

        with patch("mcq.pdf.subprocess.Popen", side_effect=launch), patch("mcq.pdf.subprocess.run"):
            with self.assertRaisesRegex(ValueError, "LibreOffice timed out.*soffice.com"):
                render_batch(Path("soffice.com"), [source], out)
            self.assertEqual(render_batch(Path("soffice.com"), [source], out), [out / "question.pdf"])
        self.assertNotEqual(commands[0][1], commands[1][1])
        self.assertEqual(list(self.root.glob("writer-*")), [])

    def test_folder_mapping_and_reopen(self):
        self.assertEqual(self.folder.relative_to(self.output).as_posix(),
                         f"NC_{date.today().year}/UNIT_10/2026-11-19-ASSISTANT_REGISTRAR")
        self.assertEqual(len({tuple(order) for order in self.exam["mapping"]}), 5)
        self.assertEqual(read_exam(self.folder)["mapping"], self.exam["mapping"])
        xlsx, pdf = mapping_files(self.folder, self.exam)
        sheet = load_workbook(xlsx).active
        self.assertEqual(tuple(sheet.values)[2], ("Set 0", "Set 1", "Set 2", "Set 3", "Set 4"))
        self.assertEqual(sheet.page_setup.orientation, "portrait")
        self.assertEqual(sheet.print_title_rows, "$1:$3")
        rows = list(sheet.values)[3:-1]
        self.assertEqual(sheet["B1"].alignment.horizontal, "center")
        self.assertIn("IT Section, BPSC", sheet.oddFooter.left.text)
        self.assertIn("&B", sheet.oddFooter.left.text)
        self.assertEqual(sheet.oddFooter.right.text, "Page &P of &N")
        self.assertEqual(sheet.cell(sheet.max_row, 1).value, CHECK.format(count=3))
        with fitz.open(pdf) as printable:
            self.assertEqual((printable[0].rect.width, printable[0].rect.height), (612, 1008))
            for value in ("Question Mapping", "2026-11-19", "Unit 10", "Assistant Registrar",
                          "IT Section, BPSC", "Page 1 of 1", CHECK.format(count=3)):
                self.assertIn(value, printable[0].get_text())
        configured = DEFAULTS | {"mapping_paper": "A4", "mapping_orientation": "Landscape"}
        xlsx, pdf = save_mapping_exports(self.folder, self.exam, configured)
        self.assertEqual(load_workbook(xlsx).active.page_setup.orientation, "landscape")
        with fitz.open(pdf) as printable:
            self.assertGreater(printable[0].rect.width, printable[0].rect.height)
        self.assertEqual(rows, [tuple(row) for row in mapping_rows(self.exam)])
        first = {record["code"] for record in self.exam["records"]}
        updated = generate_mapping(self.output, self.folder, self.exam, DEFAULTS)
        self.assertTrue(first.isdisjoint(record["code"] for record in updated["records"]))
        self.assertEqual(len(list((self.folder / "MAPPING").glob("*.json"))), 2)
        with self.assertRaises(ValueError):
            make_mapping(2, 4)

    def test_invalid_source_and_word_export(self):
        self.import_demo()
        source = Document(self.folder / "SOURCE" / "ZERO_SET.docx")
        original = source.tables[0]
        broken = self.root / "broken.docx"
        replace_text(original.rows[2].cells[0].paragraphs[0], "৩", "১০০")
        source.save(broken)
        with self.assertRaisesRegex(ValueError, r"Out of range: 100"):
            import_source(self.folder, self.exam, broken)
        template = self.settings()
        footer = Document()
        footer.add_paragraph("END OF QUESTION PAPER")
        footer_path = self.root / "footer.docx"
        footer.save(footer_path)
        template["footer"] = str(footer_path)
        source_path = self.folder / "SOURCE" / "ZERO_SET.docx"
        files = [self.root / f"set-{number}.docx" for number in (0, 1)]
        for number, target in enumerate(files):
            make_booklet(source_path, self.exam, number, template, target)
        original_docx = Document(self.folder / "SOURCE" / "ZERO_SET.docx")
        self.assertEqual(template["header"], str(self.root / "Config" / "Template" / "HEADER.docx"))
        for number, file in enumerate(files):
            booklet = Document(file)
            self.assertAlmostEqual(booklet.sections[0].page_height.inches, 14.0)
            self.assertEqual(booklet.tables[0].rows[0].cells[0].paragraphs[0].text, f"সেট - {label(number)}")
            self.assertEqual(booklet.tables[0].rows[0].cells[2].text, "ইছামতি")
            self.assertTrue(booklet.tables[0]._tbl.xpath(".//w:drawing"))
            self.assertIn("END OF QUESTION PAPER", [paragraph.text for paragraph in booklet.paragraphs])
            for index, row in enumerate(booklet.tables[1].rows):
                expected = self.exam["mapping"][number][index] - 1
                original_cell = original_docx.tables[0].rows[expected].cells[1]
                self.assertEqual(row.cells[1].text, original_cell.text)
                for actual, original in zip(row.cells[1].paragraphs[0].runs, original_cell.paragraphs[0].runs):
                    self.assertEqual(actual._r.rPr.xml, original._r.rPr.xml)
                self.assertEqual(row.cells[0].text, label(index + 1) + ".")
            record = self.exam["records"][number]
            self.assertRegex(record["code"], r"^[A-Z0-9]{8}$")
            self.assertIn(record["code"], booklet.sections[0].footer.paragraphs[0].text)
        with self.assertRaisesRegex(ValueError, "confirm"):
            export_booklets(self.folder, self.exam, template, APP, [1], ["docx"])
        broken_exam = deepcopy(self.exam)
        broken_exam["mapping"][1][0] = broken_exam["mapping"][1][1]
        with self.assertRaisesRegex(ValueError, "repeated or missing"):
            validate_mapping(broken_exam)

    @unittest.skipUnless(RENDERER, "Set MCQ_TEST_RENDERER to test with LibreOffice.")
    def test_render_multi_page_100_questions(self):
        self.folder, self.exam = create_exam(
            self.output, "Unit 10", "260100", "100 Question Test", "ইছামতি", 100, 6, "2026-11-19",
        )
        self.exam = generate_mapping(self.output, self.folder, self.exam, DEFAULTS)
        document = Document(APP / "examples" / "DEMO_ZERO_SET_3_QUESTIONS.docx")
        table = document.tables[0]
        sample = deepcopy(table.rows[0]._tr)
        for row in list(table.rows):
            table._tbl.remove(row._tr)
        for i in range(1, 101):
            table._tbl.append(deepcopy(sample))
            row = table.rows[-1]
            replace_text(row.cells[0].paragraphs[0], "১", label(i))
        fixture = self.root / "100.docx"
        document.save(fixture)
        self.exam = import_source(self.folder, self.exam, fixture)
        settings = self.settings() | {
            "paper": "A4", "sides": ["top", "bottom", "left", "right"],
            "runner_text": "BPSC RUNNER", "runner_interval": 0.8,
        }
        footer = Document()
        footer.add_paragraph("END OF QUESTION PAPER")
        footer_path = self.root / "footer.docx"
        footer.save(footer_path)
        settings["footer"] = str(footer_path)
        updates = []
        zero = export_booklets(self.folder, self.exam, settings, APP, [0], ["docx", "pdf"],
                               progress=lambda *update: updates.append(update))
        self.assertEqual(updates[-1][:2], (5, 5))
        self.assertEqual(sorted(current for current, _, _ in updates), [current for current, _, _ in updates])
        self.assertTrue(any("converting" in text for _, _, text in updates))
        self.assertEqual(review_status(self.folder, self.exam, settings), (False, zero[1]))
        with self.assertRaisesRegex(ValueError, "confirm"):
            export_booklets(self.folder, self.exam, settings, APP, [2], ["docx", "pdf"])
        confirm_zero(self.folder, self.exam, settings)
        self.assertEqual(review_status(self.folder, read_exam(self.folder), settings), (True, zero[1]))
        outputs = export_booklets(self.folder, self.exam, settings, APP, [1, 2, 3, 4, 5, 6], ["docx", "pdf"])
        self.assertTrue(generated_complete(self.folder, self.exam, settings))
        paths = outputs[2:4]
        before = self.exam["records"][2]["timestamp"]
        self.assertRegex(qr_content(2, self.exam["records"][2], settings, self.exam), r"[+]06:00")
        export_booklets(self.folder, self.exam, settings, APP, [2], ["docx"])
        self.assertEqual(read_exam(self.folder)["records"][2]["timestamp"], before)
        with self.assertRaisesRegex(ValueError, "confirm"):
            export_booklets(self.folder, self.exam, settings | {"header_gap": 0.2}, APP, [2], ["docx"])
        self.assertTrue(all(path.parent.name == "GENERATED_SETS" for path in paths))
        with fitz.open(paths[1]) as pdf:
            self.assertGreaterEqual(len(pdf), 3)
            self.assertAlmostEqual(pdf[0].rect.height, 841.89, delta=1)
            all_text = [page.get_text() for page in pdf]
            ident = self.exam["records"][2]["code"]
            self.assertTrue(all(ident in text for text in all_text), all_text[-1][-500:])
            self.assertTrue(all(f"Set 2 / Page {index} of {len(pdf)}" in text
                                for index, text in enumerate(all_text, 1)), all_text[-1][-500:])
            code_rect = pdf[0].search_for(ident)[0]
            paging_rect = pdf[0].search_for(f"Set 2 / Page 1 of {len(pdf)}")[0]
            self.assertLess(code_rect.x0, 50)
            self.assertGreater(paging_rect.x1, pdf[0].rect.width - 45)
            self.assertTrue(all("BPSC RUNNER" in text for text in all_text))
            self.assertTrue(pdf[0].get_image_info())
            self.assertTrue(all(not page.get_image_info() for page in list(pdf)[1:]))
            self.assertTrue(all("END OF QUESTION PAPER" not in text for text in all_text[:-1]))
            self.assertIn("END OF QUESTION PAPER", all_text[-1])
            self.assertIn("১০০", all_text[-1])
            preserve = os.environ.get("MCQ_TEST_COPY_PDF")
            if preserve:
                Path(preserve).write_bytes(paths[1].read_bytes())
        original_review = zero[1].read_bytes()
        zero[1].write_bytes(original_review + b"\n% changed after review")
        with self.assertRaisesRegex(ValueError, "confirm"):
            export_booklets(self.folder, self.exam, settings, APP, [1], ["docx"])

        bangla = settings | {"footer_language": "Bangla", "serial_suffix": ""}
        zero = export_booklets(self.folder, self.exam, bangla, APP, [0], ["docx", "pdf"])
        with fitz.open(zero[1]) as pdf:
            for index, page in enumerate(pdf, 1):
                self.assertIn(f"সেট - ০, পৃষ্ঠা {label(index)}/{label(len(pdf))}", page.get_text())
        word = Document(zero[0])
        self.assertEqual(word.tables[1].rows[0].cells[0].text, "১")

    def test_no_shared_serial_positions_when_possible(self):
        for count, sets in [(5, 4), (100, 4), (500, 25)]:
            orders = make_mapping(count, sets)
            for values in zip(*orders):
                self.assertEqual(len(set(values)), sets + 1)
        self.assertEqual(len(make_mapping(3, 4)), 5)

    def test_optional_tags_and_qr_fields(self):
        self.import_demo()
        untagged = Document()
        untagged.add_paragraph("Static Header")
        template = self.root / "plain.docx"
        untagged.save(template)
        settings = DEFAULTS | {
            "header": save_template(self.root, "header", template),
            "qr_fields": ["post", "unit", "text"],
            "credit": "Custom QR Text",
        }
        source = self.folder / "SOURCE" / "ZERO_SET.docx"
        target = self.root / "set.docx"
        make_booklet(source, self.exam, 0, settings, target)
        self.assertEqual(Document(target).paragraphs[0].text, "Static Header")
        self.assertEqual(qr_content(0, self.exam["records"][0], settings, self.exam),
                         "Post: Assistant Registrar\nUnit: Unit 10\nCustom QR Text")
        short = self.root / "short.docx"
        short_header = Document()
        paragraph = short_header.add_paragraph("[[QR]] · [[SET]] · [[CODE]] · [[SET_NUMBER]] · [[SET_QR]]")
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        short_header.save(short)
        settings["header"] = str(short)
        settings["code_name"] = "unused"
        exam = {**self.exam, "code_name": "ইছামতি"}
        make_booklet(source, exam, 0, settings, target)
        filled = Document(target).paragraphs[0]
        self.assertEqual(filled.alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(len(filled._p.xpath(".//w:drawing")), 2)
        self.assertNotIn("[[", filled.text)
        self.assertIn("০ · ইছামতি · ০", filled.text)
        cadre, _ = create_exam(self.output, "Unit 10", "260010", "Another Post", "ইছামতি",
                               3, 4, "2025-01-01", "Cadre")
        self.assertEqual(cadre.relative_to(self.output).parts[0], f"CADRE_{date.today().year}")

    def test_optional_fields_units_and_access(self):
        units = self.root / "Config" / "units.json"
        self.assertEqual(load_units(units)[0], "Unit 01")
        self.assertEqual(len(load_units(units)), 23)
        save_json(units, ["Custom Unit"])
        self.assertEqual(load_units(units), ["Custom Unit"])
        with self.assertRaisesRegex(ValueError, "six-digit"):
            create_exam(self.output, "Unit 01", "123", "Invalid", "Name", 3, 4)
        with self.assertRaisesRegex(ValueError, "whole number"):
            create_exam(self.output, "Unit 01", "", "Invalid", "Name", 3, 4, grade="-1")
        folder, exam = create_exam(self.output, "Unit 01", "", "New Post", "Name", 3, 4, grade="9")
        self.assertEqual(exam["post_code"], "")
        self.assertEqual(exam["post_grade"], 9)
        credentials = initialize_access(self.root)
        credentials.write_text("email=operator@example.invalid\npassword=custom=secret\n", "utf-8")
        self.assertFalse(matches(credentials, "wrong", "operator@example.invalid"))
        self.assertTrue(matches(credentials, "custom=secret", "Operator@example.invalid"))
        self.assertTrue(matches(initialize_access(self.root), "custom=secret"))
        delete_exam(self.output, folder)
        self.assertFalse(folder.exists())
        self.assertTrue(self.folder.exists())

    def test_exported_columns_and_repeating_pdf_metadata(self):
        folder, exam = create_exam(self.output, "Unit 10", "260055", "Test Grade", "Name",
                                   100, 4, grade="10")
        exam = generate_mapping(self.output, folder, exam, DEFAULTS)
        xlsx, pdf = mapping_files(folder, exam)
        with fitz.open(pdf) as printable:
            self.assertEqual(len(printable), 2)
            for index, page in enumerate(printable, 1):
                text = page.get_text()
                for phrase in ("Question Mapping", "Test Grade (Grade: 10)", "Set 0", "Set 1",
                               "Set 4", "IT Section, BPSC", f"Page {index} of {len(printable)}"):
                    self.assertIn(phrase, text)
            self.assertIn(CHECK.format(count=100), printable[-1].get_text())
        rows = verified_rows(exam)
        for column in zip(*rows):
            self.assertEqual(set(column), set(range(1, 101)))
        rows[0][3] = rows[1][3]
        with patch("mcq.mapping.mapping_rows", return_value=rows):
            with self.assertRaisesRegex(ValueError, "exactly once"):
                verified_rows(exam)
        sheet = load_workbook(xlsx).active
        self.assertEqual(sheet["A2"].value, "Post Name: Test Grade (Grade: 10)")

    def test_template_tag_appearance_and_qr_alignment(self):
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.oxml.ns import qn
        from docx.shared import Inches, Pt

        self.import_demo()
        template = Document()
        cells = template.add_table(rows=1, cols=3).rows[0].cells
        set_tag = cells[0].paragraphs[0].add_run("[[SE")
        set_tag.font.name = "Nikosh"
        set_tag.font.size = Pt(16)
        set_tag.bold = True
        cells[0].paragraphs[0].add_run("T]]")
        code_tag = cells[1].paragraphs[0].add_run("[[CODE]]")
        code_tag.font.name = "Nikosh"
        code_tag.font.size = Pt(14)
        code_tag.italic = True
        cells[2].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
        cells[2].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.BOTTOM
        cells[2].paragraphs[0].add_run("\t[[QR]]  ")
        header = self.root / "formatted.docx"
        template.save(header)
        source = self.folder / "SOURCE" / "ZERO_SET.docx"
        target = self.root / "booklet.docx"
        settings = DEFAULTS | {"header": str(header), "qr_size": 0.65}
        make_booklet(source, self.exam, 0, settings, target)
        result = Document(target)
        cells = result.tables[0].rows[0].cells
        for cell, size, style in ((cells[0], 16, "bCs"), (cells[1], 14, "iCs")):
            run = next(run for run in cell.paragraphs[0].runs if run.text)
            self.assertEqual(run.font.size.pt, size)
            self.assertEqual(run._r.rPr.find(qn("w:szCs")).get(qn("w:val")), str(size * 2))
            self.assertEqual(run._r.rPr.rFonts.get(qn("w:cs")), "Nikosh")
            self.assertIsNotNone(run._r.rPr.find(qn("w:" + style)))
        self.assertEqual(result.inline_shapes[0].width, Inches(0.65))
        self.assertEqual(cells[2].paragraphs[0].text, "")
        self.assertEqual(cells[2].paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.RIGHT)
        self.assertEqual(cells[2].vertical_alignment, WD_CELL_VERTICAL_ALIGNMENT.BOTTOM)

        make_booklet(source, self.exam, 0,
                     settings | {"qr_horizontal": "Center", "qr_vertical": "Top"}, target)
        qr_cell = Document(target).tables[0].rows[0].cells[2]
        self.assertEqual(qr_cell.paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)
        self.assertEqual(qr_cell.vertical_alignment, WD_CELL_VERTICAL_ALIGNMENT.TOP)

    @unittest.skipUnless(RENDERER, "Set MCQ_TEST_RENDERER to test with LibreOffice.")
    def test_shipped_templates_qr_position_and_fallback(self):
        from shutil import copytree
        from docx.shared import Inches
        from mcq.documents import template_paragraphs
        from mcq.pdf import render_words

        self.import_demo()
        base = self.root / "templates"
        defaults = base / "Config" / "Template"
        copytree(APP / "Config" / "Template", defaults)
        settings = DEFAULTS | {"template_root": str(base), "renderer": RENDERER, "font": "Nikosh"}
        source = self.folder / "SOURCE" / "ZERO_SET.docx"
        signature = review_signature(self.folder, self.exam, settings)
        marker = Document(defaults / "DEFAULT_HEADER.docx")
        for paragraph in template_paragraphs(marker):
            replace_text(paragraph, "[[QR]]", "QR-ANCHOR")
        custom = self.root / "custom.docx"
        marker.save(custom)
        versions = (
            ("marker", {"header": str(custom), "use_default_header": False}),
            ("qr", {}),
            ("adjusted", {"qr_size": 0.72, "qr_inset": 0.15, "qr_gap": 0.05,
                          "qr_horizontal": "Right", "qr_vertical": "Top"}),
            ("excluded", {"use_default_header": False, "use_default_footer": False}),
        )
        sources = []
        for name, overrides in versions:
            target = self.root / f"{name}.docx"
            make_booklet(source, self.exam, 0, settings | overrides, target)
            sources.append(target)
        self.assertNotEqual(signature, review_signature(self.folder, self.exam, settings | versions[-1][1]))
        self.assertEqual(Document(sources[2]).inline_shapes[0].width, Inches(0.72))
        paths = render_words(sources, self.root / "pdf", settings, self.root)
        with fitz.open(paths[0]) as pdf:
            marker_rect = pdf[0].search_for("QR-ANCHOR")[0]
        with fitz.open(paths[1]) as pdf:
            qr = fitz.Rect(pdf[0].get_image_info()[0]["bbox"])
            self.assertAlmostEqual(qr.x1, marker_rect.x1, delta=1.5)
            self.assertAlmostEqual(qr.y0, marker_rect.y0, delta=3)
            self.assertTrue(pdf[-1].search_for("---------- 0 ---------"))
        with fitz.open(paths[2]) as pdf:
            tuned = fitz.Rect(pdf[0].get_image_info()[0]["bbox"])
            self.assertAlmostEqual(tuned.width, 0.72 * 72, delta=0.5)
            self.assertAlmostEqual(qr.x1 - tuned.x1, 0.15 * 72, delta=0.5)
            self.assertAlmostEqual(tuned.y0 - qr.y0, 0.05 * 72, delta=0.5)
        with fitz.open(paths[3]) as pdf:
            self.assertFalse(pdf[0].get_image_info())
            self.assertFalse(pdf[-1].search_for("---------- 0 ---------"))
            self.assertTrue(pdf[-1].search_for("Set 0 / Page 1 of 1"))

    @unittest.skipUnless(RENDERER, "Set MCQ_TEST_RENDERER to test with LibreOffice.")
    def test_runner_footer_appearance_and_settings(self):
        from mcq.ui import SettingsDialog, Window

        self.import_demo()
        app_root = self.root / "app"
        settings = self.settings() | {"output_root": str(self.output)}
        save_json(app_root / "Config" / "settings.json", settings)
        window = Window(app_root)
        self.addCleanup(window.close)
        window.open_exam(window.saved.item(0))
        self.assertIn("Post Code: 260010", window.context.text())
        self.assertIn("Unit 10", window.context.text())
        dialog = SettingsDialog(window)
        self.addCleanup(dialog.close)
        tabs = {control.text(): control for control in dialog.navigation_buttons.buttons()}
        for name in ("Templates", "Question Pages", "QR Content", "Runner", "Footer Text"):
            self.assertFalse(tabs[name].icon().pixmap(16, 16).isNull())
            tabs[name].click()
            self.assertEqual(dialog.tabs.currentIndex(), dialog.navigation_buttons.id(tabs[name]))
        self.assertEqual(dialog.fields["runner_size"].value(), 10)
        self.assertEqual(dialog.fields["runner_opacity"].value(), 0.8)
        self.assertEqual(dialog.fields["footer_language"].currentData(), "English")
        self.assertEqual(dialog.navigation.rowCount(), 2)
        self.assertEqual(len(tabs), dialog.tabs.count())
        self.assertEqual([window.exam_type.itemText(i) for i in range(window.exam_type.count())],
                         ["NON CADRE", "CADRE"])
        self.assertEqual(dialog.fields["qr_size"].value(), 0.75)
        self.assertTrue(dialog.fields["print_code_enabled"].isChecked())
        self.assertEqual(dialog.fields["print_code_size"].value(), 8)
        self.assertEqual(dialog.fields["print_code_color"].text(), "#000000")
        self.assertEqual(dialog.fields["print_code_font"].text(), "Courier New")
        self.assertEqual(dialog.fields["print_code_opacity"].value(), 0.9)
        for key in ("font_size", "runner_size", "footer_size", "print_code_size"):
            control = dialog.fields[key]
            original = control.value()
            control.stepUp()
            self.assertEqual(control.value(), original + 0.5)
            control.stepDown()
            self.assertEqual(control.value(), original)
        for key in ("qr_inset", "qr_gap"):
            dialog.fields[key].setValue(-0.15)
            self.assertEqual(dialog.fields[key].value(), -0.15)
        self.exam["code_name"] = "Test Name"
        save_exam(self.folder, self.exam)
        dialog.fields["runner_font"].setText("Nikosh")
        dialog.fields["runner_size"].setValue(11)
        dialog.fields["runner_text"].setText("runner [[CODE]]")
        dialog.sides["top"].setChecked(True)
        dialog.fields["runner_color"].setText("#204060")
        dialog.fields["runner_uppercase"].setChecked(True)
        dialog.fields["footer_font"].setText("Nikosh")
        dialog.fields["footer_size"].setValue(12)
        dialog.fields["footer_color"].setText("#803020")
        dialog.fields["footer_opacity"].setValue(0.60)
        dialog.fields["footer_uppercase"].setChecked(True)
        dialog.fields["print_code_font"].setText("Nikosh")
        dialog.fields["print_code_size"].setValue(7.5)
        dialog.fields["print_code_color"].setText("#123456")
        dialog.fields["print_code_opacity"].setValue(0.75)
        dialog.fields["use_default_footer"].setChecked(False)
        dialog.save()
        self.assertEqual(dialog.result(), dialog.DialogCode.Accepted)
        self.assertFalse(window.settings["use_default_footer"])
        docx, pdf = export_booklets(self.folder, self.exam, window.settings, app_root, [0], ["docx", "pdf"])
        footer = Document(docx).sections[0].footer.paragraphs[0]
        self.assertIn("SET 0 / PAGE ", footer.text)
        self.assertEqual(footer.runs[-1].font.name, "Nikosh")
        self.assertEqual(footer.runs[-1].font.size.pt, 12)
        self.assertEqual(str(footer.runs[-1].font.color.rgb), "B38379")
        identifier = self.exam["records"][0]["code"]
        self.assertEqual(footer.runs[0].text.strip(), identifier)
        self.assertEqual(footer.runs[0].font.size.pt, 7.5)
        with fitz.open(pdf) as booklet:
            page = booklet[0]
            self.assertIn(identifier, page.get_text())
            self.assertIn(f"RUNNER {self.exam['code_name'].upper()}", page.get_text())
            self.assertIn("SET 0 / PAGE 1 OF 1", page.get_text())
            styling = [(span["size"], span["opacity"], span["color"], span["font"])
                       for span in page.get_texttrace()]
            for size, opacity, color in ((11, 0.8, (32, 64, 96)), (12, 0.6, (128, 48, 32)),
                                         (7.5, 0.75, (18, 52, 86))):
                self.assertTrue(any(abs(found_size - size) < 0.02 and abs(alpha - opacity) < 0.02
                                    and all(abs(found - expected / 255) < 0.02
                                            for found, expected in zip(rgb, color))
                                    and "nikosh" in font.lower()
                                    for found_size, alpha, rgb, font in styling), styling)
        disabled = {**window.settings, "print_code_enabled": False}
        docx, pdf = export_booklets(self.folder, self.exam, disabled, app_root, [0], ["docx", "pdf"])
        self.assertNotIn(identifier, Document(docx).sections[0].footer.paragraphs[0].text)
        with fitz.open(pdf) as booklet:
            self.assertNotIn(identifier, booklet[0].get_text())
            self.assertIn("SET 0 / PAGE 1 OF 1", booklet[0].get_text())

    def test_create_exam_in_existing_year_and_unit(self):
        original = (self.folder / "exam.json").read_bytes()
        unit = self.folder.parent
        year = unit.parent
        sentinel = unit / "existing.txt"
        sentinel.write_text("preserve")
        name = "Retained Folder"
        expected = unit / "2026-11-21-RETAINED_FOLDER"
        expected.mkdir()
        leftover = expected / "existing.txt"
        leftover.write_text("preserve")
        folder, exam = create_exam(self.output, "Unit 10", "", name, "Code", 3, 4, "2026-11-21")
        self.assertEqual(folder, expected)
        self.assertEqual(exam["exam_type"], "NON CADRE")
        self.assertEqual(year.name, f"NC_{date.today().year}")
        self.assertEqual(leftover.read_text(), "preserve")
        another, _ = create_exam(self.output, "Unit 10", "012345", name, "Code", 3, 4, "2026-11-21")
        self.assertEqual(another, unit / "2026-11-21-RETAINED_FOLDER-012345")
        with self.assertRaisesRegex(ValueError, "Saved Exams"):
            create_exam(self.output, "Unit 10", "012345", name, "New Code", 3, 4, "2026-11-21")
        cadre, exam = create_exam(self.output, "Unit 10", "", name, "Code", 3, 4, "2026-11-21", "CADRE")
        self.assertEqual(cadre.relative_to(self.output).parts[0], f"CADRE_{date.today().year}")
        self.assertEqual(exam["exam_type"], "CADRE")
        self.assertEqual(sentinel.read_text(), "preserve")
        self.assertEqual((self.folder / "exam.json").read_bytes(), original)

    def test_font_choice_updates_fields_and_survives_reopen(self):
        from PySide6.QtCore import QTimer
        from PySide6.QtGui import QFontDatabase
        from PySide6.QtWidgets import QComboBox, QDialogButtonBox, QInputDialog, QPushButton
        from mcq.ui import SettingsDialog, Window

        app_root = self.root / "app"
        app_root.mkdir()
        window = Window(app_root)
        dialog = SettingsDialog(window)
        families = QFontDatabase.families()
        name = next((item for item in ("Arial Black", "Liberation Sans") if item in families), None)
        name = name or next(item for item in families if item != "Arial")
        dialog.show()

        def select_family():
            picker = QApplication.activeModalWidget()
            if isinstance(picker, QInputDialog):
                combo = picker.findChild(QComboBox)
                combo.setCurrentIndex(combo.findText(name))
                picker.findChild(QDialogButtonBox).button(QDialogButtonBox.StandardButton.Ok).click()

        for section_name, field in (("Question Pages", "font"), ("Runner", "runner_font"),
                                    ("Footer Text", "footer_font")):
            section = next(control for control in dialog.navigation_buttons.buttons() if control.text() == section_name)
            section.click()
            chooser = next(button for button in dialog.tabs.currentWidget().findChildren(QPushButton)
                           if button.text() == "Choose Font")
            QTimer.singleShot(0, select_family)
            chooser.click()
            self.assertEqual(dialog.fields[field].text(), name)
        self.assertIn(name, dialog.fields["footer_font"].placeholderText())
        dialog.fields["footer_font"].setText("Arial")
        dialog.save()
        reopened = SettingsDialog(Window(app_root))
        self.assertEqual([reopened.fields[key].text() for key in ("font", "runner_font", "footer_font")],
                         [name, name, "Arial"])
        reopened.fields["footer_font"].clear()
        reopened.save()
        self.import_demo()
        target = self.root / "font.docx"
        make_booklet(self.folder / "SOURCE" / "ZERO_SET.docx", self.exam, 0,
                     reopened.owner.settings, target)
        self.assertEqual(Document(target).sections[0].footer.paragraphs[0].runs[-1].font.name, name)

    @unittest.skipUnless(RENDERER, "Set MCQ_TEST_RENDERER to test with LibreOffice.")
    def test_runner_and_global_tag_languages_are_independent(self):
        from mcq.pdf import finish_booklet_pdf, render_words, runner_overlay

        self.import_demo()
        header = Document()
        header.add_paragraph("Header [[SET]] / [[SET_NUMBER]] / [[CODE]]").runs[0].font.name = "Nikosh"
        template = self.root / "header.docx"
        header.save(template)
        settings = DEFAULTS | {"header": str(template), "renderer": RENDERER,
                               "runner_font": "Nikosh", "sides": ["top"], "runner_interval": 7}
        for index, (runner, global_language, text, expected, digits) in enumerate((
            ("English", "Bangla", "Set [[SET]]", "Set 1", "১"),
            ("Bangla", "English", "সেট [[SET_NUMBER]]", "সেট ১", "1"),
        )):
            selected = settings | {"runner_tag_language": runner, "tag_language": global_language, "runner_text": text}
            target = self.root / f"paper-{index}.docx"
            make_booklet(self.folder / "SOURCE" / "ZERO_SET.docx", self.exam, 1, selected, target)
            self.assertEqual(Document(target).paragraphs[0].text,
                             f"Header {digits} / {digits} / {self.exam['code_name']}")
            self.assertIn(f"Set {digits}", qr_content(1, self.exam["records"][1], selected, self.exam))
            rendered = render_words([target], self.root / "pdf", selected, self.root)[0]
            pdf = self.root / f"booklet-{index}.pdf"
            finish_booklet_pdf(rendered, pdf, self.exam, 1, selected)
            with fitz.open(pdf) as result:
                page = result[0]
                clip = fitz.Rect(32, 4, 90, 30)
                if runner == "English":
                    self.assertIn(expected, page.get_text(clip=clip))
                else:
                    self.assertIn("১", page.get_text(clip=clip))
                # Bengali glyphs may be shaped correctly even when a font's
                # PDF character map omits a compound glyph during extraction.
                overlay = runner_overlay(page.rect.width, page.rect.height, selected, expected, None)
                with fitz.open(stream=overlay, filetype="pdf") as reference:
                    matrix = fitz.Matrix(2, 2)
                    self.assertEqual(page.get_pixmap(matrix=matrix, clip=clip).samples,
                                     reference[0].get_pixmap(matrix=matrix, clip=clip).samples)

    @unittest.skipUnless(RENDERER, "Set MCQ_TEST_RENDERER to test with LibreOffice.")
    def test_rotated_tag_cell_alignment_controls(self):
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_ROW_HEIGHT_RULE
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        from docx.shared import Inches
        from mcq.pdf import render_words

        self.import_demo()
        header = Document()
        table = header.add_table(rows=1, cols=2)
        table.autofit = False
        row = table.rows[0]
        row.height = Inches(1.5)
        row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST
        for cell, text in zip(row.cells, ("[[SET]]", "[[CODE]]")):
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
            cell.paragraphs[0].add_run(text).font.name = "Nikosh"
        template = self.root / "header.docx"
        header.save(template)
        settings = DEFAULTS | {"header": str(template), "renderer": RENDERER,
                               "set_rotation": 90, "code_rotation": 270}
        source = self.folder / "SOURCE" / "ZERO_SET.docx"
        default, custom = self.root / "default.docx", self.root / "custom.docx"
        make_booklet(source, self.exam, 0, settings, default)
        make_booklet(source, self.exam, 0, settings | {
            "set_horizontal": "Right", "set_vertical": "Bottom",
            "code_horizontal": "Center", "code_vertical": "Center",
        }, custom)
        for path, horizontal, vertical in (
            (default, (WD_ALIGN_PARAGRAPH.LEFT,) * 2, (WD_CELL_VERTICAL_ALIGNMENT.TOP,) * 2),
            (custom, (WD_ALIGN_PARAGRAPH.RIGHT, WD_ALIGN_PARAGRAPH.CENTER),
             (WD_CELL_VERTICAL_ALIGNMENT.BOTTOM, WD_CELL_VERTICAL_ALIGNMENT.CENTER)),
        ):
            cells = Document(path).tables[0].rows[0].cells
            self.assertEqual(tuple(cell.paragraphs[0].alignment for cell in cells), horizontal)
            self.assertEqual(tuple(cell.vertical_alignment for cell in cells), vertical)
        pdfs = render_words([default, custom], self.root / "pdf", settings, self.root)
        with fitz.open(pdfs[0]) as first, fitz.open(pdfs[1]) as second:
            before = [fitz.Rect(entry["bbox"]) for entry in first[0].get_image_info()]
            after = [fitz.Rect(entry["bbox"]) for entry in second[0].get_image_info()]
            self.assertEqual(len(before), 2)
            self.assertEqual(len(after), 2)
            for original, adjusted in zip(before, after):
                self.assertGreater(adjusted.x0 - original.x0, 50)
                self.assertGreater(adjusted.y0 - original.y0, 20)

    @unittest.skipUnless(RENDERER, "Set MCQ_TEST_RENDERER to test with LibreOffice.")
    def test_signed_qr_offsets_and_independent_tag_rotation(self):
        from io import BytesIO
        from PIL import Image
        from mcq.pdf import render_words

        self.import_demo()
        source = self.folder / "SOURCE" / "ZERO_SET.docx"
        settings = DEFAULTS | {"template_root": str(APP), "renderer": RENDERER,
                               "qr_horizontal": "Right", "qr_vertical": "Top"}
        variants = (
            ("normal", {}),
            ("negative", {"qr_inset": -0.15, "qr_gap": -0.05}),
            ("clockwise", {"set_rotation": 90, "code_rotation": 270}),
            ("reverse", {"set_rotation": 270, "code_rotation": 90}),
        )
        paths = []
        for name, changes in variants:
            target = self.root / f"{name}.docx"
            make_booklet(source, self.exam, 0, settings | changes, target)
            paths.append(target)
        self.assertNotEqual(review_signature(self.folder, self.exam, settings),
                            review_signature(self.folder, self.exam, settings | variants[-1][1]))
        self.assertFalse(Document(paths[0]).element.xpath(".//wp:anchor"))
        self.assertTrue(Document(paths[1]).element.xpath(".//wp:anchor"))

        def drawings(path):
            document = Document(path)
            return {drawing.docPr.get("descr"): Image.open(BytesIO(
                        document.part.related_parts[drawing.xpath(".//a:blip")[0].get(
                            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")].blob))
                    for drawing in (shape._inline for shape in document.inline_shapes)
                    if drawing.docPr.get("descr")}

        clockwise, reverse = drawings(paths[2]), drawings(paths[3])
        self.assertEqual(set(clockwise), {"০", self.exam["code_name"]})
        for title, image in clockwise.items():
            self.assertEqual(image.transpose(Image.Transpose.ROTATE_180).tobytes(), reverse[title].tobytes())
        self.assertGreater(clockwise[self.exam["code_name"]].height, clockwise[self.exam["code_name"]].width)
        pdfs = render_words(paths, self.root / "pdf", settings, self.root)
        with fitz.open(pdfs[0]) as normal, fitz.open(pdfs[1]) as moved, fitz.open(pdfs[2]) as rotated:
            qr = fitz.Rect(normal[0].get_image_info()[0]["bbox"])
            shifted = fitz.Rect(next(image["bbox"] for image in moved[0].get_image_info() if image["width"] > 1))
            self.assertAlmostEqual(shifted.x1 - qr.x1, 0.15 * 72, delta=0.5)
            self.assertAlmostEqual(shifted.y0 - qr.y0, -0.05 * 72, delta=0.5)
            images = [fitz.Rect(image["bbox"]) for image in rotated[0].get_image_info()]
            self.assertEqual(len(images), 3)
            self.assertTrue(any(rect.height > 2 * rect.width for rect in images))

    def test_mapping_pdf_margin_and_compact_header(self):
        from mcq.mapping import export_mapping_pdf

        folder, exam = create_exam(self.output, "Unit 10", "260055", "Registrar", "Name", 100, 8)
        exam = generate_mapping(self.output, folder, exam, DEFAULTS)
        _, pdf = mapping_files(folder, exam)

        def header_locations(path):
            with fitz.open(path) as printable:
                self.assertEqual(len(printable), 2)
                page = printable[0]
                code = page.search_for("Post Code: 260055")[0]
                title = page.search_for("Question Mapping")[0]
                unit = page.search_for("Unit: Unit 10")[0]
                self.assertTrue(page.search_for("Set 8"))
                self.assertLess(abs(code.y0 - title.y0), 4)
                self.assertLess(abs(unit.y0 - title.y0), 4)
                self.assertAlmostEqual((title.x0 + title.x1) / 2, page.rect.width / 2, delta=1)
                return code.x0, code.y0
        left, top = header_locations(pdf)
        self.assertAlmostEqual(left, 36, delta=1)
        custom = self.root / "custom.pdf"
        export_mapping_pdf(exam, custom, DEFAULTS | {"mapping_margin_left": 0.75, "mapping_margin_top": 0.75})
        with fitz.open(custom) as printable:
            self.assertEqual(len(printable), 2)
            code = printable[0].search_for("Post Code: 260055")[0]
            self.assertAlmostEqual(code.x0 - left, 18, delta=1)
            self.assertAlmostEqual(code.y0 - top, 18, delta=1)

    def test_legacy_windows_renderer_slashes_preserve_review(self):
        self.import_demo()
        settings = DEFAULTS | {"renderer": "C:/Program Files/LibreOffice/program/soffice.com"}
        output = self.folder / "GENERATED_SETS"
        output.mkdir()
        pdf = output / "SET0.pdf"
        with fitz.open() as doc:
            doc.new_page()
            doc.save(pdf)

        # Reproduce the published v0.3.1 signature before a Windows restart.
        later_settings = {"runner_font", "runner_size", "runner_color", "runner_opacity", "runner_uppercase",
                          "footer_font", "footer_size", "footer_color", "footer_opacity", "footer_uppercase",
                          "qr_inset", "qr_gap", "use_default_header", "use_default_footer",
                          "set_rotation", "code_rotation", "set_horizontal", "set_vertical",
                          "code_horizontal", "code_vertical", "tag_language", "runner_tag_language"}
        layout = {key: value for key, value in settings.items() if not key.startswith("mapping_")
                  and key not in {"output_root", "template_root", "renderer", "qr_horizontal", "qr_vertical"}
                  and key not in later_settings}
        payload = {"layout": layout, "renderer": settings["renderer"], "mapping": self.exam["mapping"],
                   "codes": [record["code"] for record in self.exam["records"]], "code_name": self.exam["code_name"],
                   "source": digest(self.folder / "SOURCE" / "ZERO_SET.docx")}
        signature = sha256(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode()).hexdigest()
        self.exam["review"] = {"file": pdf.name, "signature": signature, "pdf_digest": digest(pdf), "confirmed_at": "2026-11-20"}
        save_exam(self.folder, self.exam)
        reopened = read_exam(self.folder)
        restarted = settings | {"renderer": settings["renderer"].replace("/", chr(92))}
        self.assertEqual(review_status(self.folder, reopened, restarted), (True, pdf))
        self.assertEqual(review_signature(self.folder, reopened, settings),
                         review_signature(self.folder, reopened, restarted | {"font_size": 9.0}))
        self.assertEqual(review_status(self.folder, reopened, restarted | {"qr_horizontal": "Right"}), (False, None))
        self.assertEqual(review_status(self.folder, reopened, restarted | {"renderer": "D:/Other/soffice.com"}), (False, None))

    @unittest.skipUnless(RENDERER, "Set MCQ_TEST_RENDERER to test with LibreOffice.")
    def test_restart_keeps_steps_and_downgrades_only_missing_outputs(self):
        from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
        from docx.enum.text import WD_TAB_ALIGNMENT
        from docx.shared import Inches
        from mcq.ui import Window

        self.import_demo()
        self.exam["post_grade"] = 10
        save_exam(self.folder, self.exam)
        template = Document()
        table = template.add_table(rows=1, cols=1)
        table.autofit = False
        table.columns[0].width = Inches(7.5)
        cell = table.cell(0, 0)
        cell.width = Inches(7.5)
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.TOP
        paragraph = cell.paragraphs[0]
        paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(7.2), WD_TAB_ALIGNMENT.RIGHT)
        paragraph.add_run("\t[[QR]]  ")
        header = self.root / "right-top.docx"
        template.save(header)
        settings = self.settings() | {"header": str(header), "output_root": str(self.output)}
        app_root = self.root / "app"
        save_json(app_root / "Config" / "settings.json", settings)
        zero = export_booklets(self.folder, self.exam, settings, app_root, [0], ["docx", "pdf"])
        with fitz.open(zero[1]) as booklet:
            first = booklet[0]
            rect = fitz.Rect(first.get_image_info()[0]["bbox"])
            self.assertGreater(rect.x1, first.rect.width - 55)
            self.assertLess(rect.y0, 55)
        confirm_zero(self.folder, self.exam, settings)
        outputs = export_booklets(self.folder, self.exam, settings, app_root, [1, 2, 3, 4], ["pdf"])
        export_booklets(self.folder, self.exam, settings, app_root, [1], ["docx"])
        self.assertEqual(len(self.exam["generated"]["1"]), 2)
        window = Window(app_root)
        self.addCleanup(window.close)
        entry = window.saved.item(0)
        self.assertIn("Assistant Registrar\nPost Grade: 10\n", entry.text())
        window.open_exam(entry)
        self.assertEqual(window.tabs.currentIndex(), 4)
        self.assertEqual(window.generated_state.text(), "4 of 4 sets generated.")
        self.assertIn("(Grade: 10)", window.context.text())

        outputs[0].unlink()
        window.open_exam(entry)
        self.assertEqual(window.tabs.currentIndex(), 4)
        self.assertIn("3 of 4 sets generated", window.generated_state.text())
        self.assertEqual(generated_numbers(self.folder, read_exam(self.folder)), [2, 3, 4])
        zero[0].unlink()
        window.open_exam(entry)
        self.assertEqual(window.tabs.currentIndex(), 3)
        self.assertFalse(window.tabs.isTabEnabled(4))
        export_booklets(self.folder, self.exam, settings, app_root, [0], ["docx", "pdf"])
        confirm_zero(self.folder, self.exam, settings)
        window.open_exam(entry)
        self.assertEqual(window.tabs.currentIndex(), 4)
        self.assertIn("3 of 4 sets generated", window.generated_state.text())
        self.assertFalse(generated_complete(self.folder, self.exam, settings | {"qr_horizontal": "Right"}))
