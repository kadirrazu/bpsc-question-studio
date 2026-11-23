"""Room Bundles window and its background worker."""

from dataclasses import asdict
import json
from pathlib import Path
from time import perf_counter

from PySide6.QtCore import QThread, Qt, QUrl, Signal
from PySide6.QtGui import QColor, QDesktopServices, QFontDatabase
from PySide6.QtWidgets import (
    QCheckBox, QColorDialog, QDialog, QFileDialog, QFormLayout, QGridLayout, QHBoxLayout,
    QLabel, QLineEdit, QMessageBox, QPlainTextEdit, QPushButton, QScrollArea, QTabWidget, QVBoxLayout, QWidget,
)

from .room_bundles import (BundleOptions, COVER_FIELDS, PAPER_SIZES, ROOM_CODE_POSITIONS,
                           ROOM_CODE_SCOPES, build_bundles)
from .appearance import icon
from .controls import Choice, CountInput, DecimalInput
from .file_io import save_json


def button(text, action, primary=False, symbol=None):
    control = QPushButton(text)
    control.clicked.connect(action)
    if symbol:
        control.setIcon(icon(symbol, light=primary))
    if primary:
        control.setObjectName("bundlePrimary")
    return control


def subtitle(text):
    label = QLabel(text)
    label.setWordWrap(True)
    return label


class BundleTask(QThread):
    progress = Signal(str)
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, operation, parent):
        super().__init__(parent)
        self.operation = operation

    def run(self):
        try:
            self.succeeded.emit(self.operation(self.progress.emit))
        except Exception as error:
            self.failed.emit(str(error))


class RoomBundlesDialog(QDialog):
    def __init__(self, app_folder, parent=None, *, set_files=None, workbook_path=None):
        super().__init__(parent)
        self.setWindowTitle("Room Print Packs")
        self.resize(1080, 760)
        self.task = None
        self.settings_path = app_folder / "Config" / "room-bundles.json"
        self.setStyleSheet("""
            QLabel#bundleTitle { background: #342451; color: white; font-size: 25px;
                                font-weight: 700; padding: 15px; border-radius: 8px; }
            QPushButton#bundlePrimary { background: #62458B; color: white; border-color: #62458B; }
            QPushButton#bundlePrimary:disabled { background: #EAE5F3; color: #5F5572; border-color: #DAD0EB; }
            QPlainTextEdit#bundleConsole { background: #171326; color: #F4EFFC;
                border: 1px solid #594777; border-radius: 8px; padding: 10px; font-size: 13px; }
            QLabel#consoleHeading { color: #372553; font-weight: 700; font-size: 15px; }
            QLabel#bundleStatus { color: #483A60; background: #EBE5F4; padding: 7px 12px; border-radius: 6px; }
            QLabel#bundleStatus[state="success"] { color: #115B3B; background: #DFF3E8; font-weight: 700; }
            QLabel#bundleStatus[state="warning"] { color: #674300; background: #FFF1D5; font-weight: 700; }
            QLabel#bundleStatus[state="error"] { color: #841F34; background: #FFE4E9; font-weight: 700; }
            QTabWidget#bundleTabs::pane { background: #F5F0FB; border-top: 1px solid #CDBEE0; }
            QTabBar::tab { background: #E9E0F4; color: #493860; border: 1px solid #D1C2E4;
                font-size: 14px; font-weight: 700; padding: 9px 16px; margin-right: 6px; }
            QTabBar::tab:selected { background: #62458B; color: white; border-color: #62458B; }
            QTabBar::tab:!selected:hover { background: #DED0F0; }
            QCheckBox::indicator:checked:enabled { background-color: #62458B; border-color: #62458B; }
        """)
        defaults = {}
        settings_error = ""
        if self.settings_path.is_file():
            try:
                defaults = json.loads(self.settings_path.read_text("utf-8"))
                if not isinstance(defaults, dict):
                    raise ValueError("Expected settings saved by Room Print Packs.")
            except (ValueError, OSError) as error:
                defaults = {}
                settings_error = str(error)
        bundle = BundleOptions()
        self.output = None
        self.settings_snapshot = defaults
        main = QVBoxLayout(self)
        title = QLabel("Room Print Packs")
        title.setObjectName("bundleTitle")
        main.addWidget(title)
        main.addWidget(subtitle("Select a center workbook and the original question PDFs. Each room starts with Set 1, then repeats the selected order until question_count is reached."))
        self.form = QTabWidget()
        self.form.setObjectName("bundleTabs")
        self.tab_icons = []
        self.form.currentChanged.connect(self.update_tab_icons)
        main.addWidget(self.form, 3)
        form = self.section("Source Files and Paper", "folder")
        self.workbook = self.path_picker(form, "Center Workbook (.xlsx)", str(workbook_path) if workbook_path else defaults.get("workbook", ""), "Excel Workbook (*.xlsx)")
        default_output = defaults.get("output_base") or app_folder / "Room Outputs"
        self.output_base = self.path_picker(form, "Output Base Folder", str(default_output))
        self.set_count = self.spin(len(set_files) if set_files else defaults.get("set_count", 4), 1, 25)
        form.addRow("Number of Sets", self.set_count)
        self.set_fields = []
        fields = QGridLayout()
        form.addRow("Question PDFs", fields)
        source_paths = set_files if set_files is not None else defaults.get("set_files", [])
        for index in range(self.set_count.maximum()):
            field = QLineEdit(source_paths[index] if index < len(source_paths) else "")
            field.setPlaceholderText(f"Select Set {index + 1} PDF")
            choose = button(f"Set {index + 1}", lambda checked=False, entry=field: self.choose(entry, "PDF Files (*.pdf)"), symbol="pdf")
            fields.addWidget(choose, index, 0)
            fields.addWidget(field, index, 1)
            self.set_fields.append((choose, field))
        self.set_count.valueChanged.connect(self.show_sets)
        self.show_sets(self.set_count.value())
        self.paper = Choice()
        self.paper.addItems(PAPER_SIZES)
        self.paper.setCurrentText(defaults.get("paper", bundle.paper))
        form.addRow("Output Paper Size", self.paper)
        self.duplex = QCheckBox("Add a blank after odd-page question papers for duplex printing")
        self.duplex.setChecked(defaults.get("duplex", bundle.duplex))
        form.addRow(self.duplex)
        form = self.section("Cover Page", "templates")
        self.cover = QCheckBox("Include a Cover Page")
        self.cover.setChecked(defaults.get("cover", bundle.cover))
        form.addRow(self.cover)
        self.cover_title = QLineEdit(defaults.get("cover_title", bundle.cover_title))
        form.addRow("Cover Page Heading", self.cover_title)
        self.cover_blanks = self.spin(defaults.get("cover_blanks", bundle.cover_blanks))
        form.addRow("Blank Pages After Cover", self.cover_blanks)
        self.cover_fields = self.page_fields(form, "Cover Table Fields", defaults.get("cover_fields", bundle.cover_fields))
        self.cover_uppercase = QCheckBox("Uppercase Cover Title and Table Text")
        self.cover_uppercase.setChecked(defaults.get("cover_uppercase", bundle.cover_uppercase))
        form.addRow(self.cover_uppercase)
        self.cover_credit = QCheckBox("Show Software Credit on the Cover")
        self.cover_credit.setChecked(defaults.get("cover_credit", bundle.cover_credit))
        form.addRow(self.cover_credit)
        self.cover_timestamp = QCheckBox("Show Timestamp on the Cover")
        self.cover_timestamp.setChecked(defaults.get("cover_timestamp", bundle.cover_timestamp))
        form.addRow(self.cover_timestamp)
        self.footer_size = DecimalInput()
        self.footer_size.setRange(5, 20)
        self.footer_size.setDecimals(1)
        self.footer_size.setSingleStep(0.5)
        self.footer_size.setValue(defaults.get("footer_size", bundle.footer_size))
        form.addRow("Footer Font Size (pt)", self.footer_size)
        self.footer_opacity = DecimalInput()
        self.footer_opacity.setRange(0, 1)
        self.footer_opacity.setDecimals(2)
        self.footer_opacity.setSingleStep(0.05)
        self.footer_opacity.setValue(defaults.get("footer_opacity", bundle.footer_opacity))
        form.addRow("Footer Opacity (0–1)", self.footer_opacity)
        form.addRow(subtitle("Footer size and opacity apply to both pages."))
        form = self.section("End Page", "footer")
        self.end = QCheckBox("Include an End Page")
        self.end.setChecked(defaults.get("end", bundle.end))
        form.addRow(self.end)
        self.end_blanks = self.spin(defaults.get("end_blanks", bundle.end_blanks))
        form.addRow("Blank Pages Before End Note", self.end_blanks)
        self.end_note = QPlainTextEdit(defaults.get("end_note", bundle.end_note))
        self.end_note.setFixedHeight(78)
        form.addRow("End Note on Last Page", self.end_note)
        self.end_fields = self.page_fields(form, "End Page Table Fields", defaults.get("end_fields", bundle.end_fields))
        self.end_uppercase = QCheckBox("Uppercase End Note and Table Text")
        self.end_uppercase.setChecked(defaults.get("end_uppercase", bundle.end_uppercase))
        form.addRow(self.end_uppercase)
        self.end_timestamp = QCheckBox("Show Timestamp on the End Page")
        self.end_timestamp.setChecked(defaults.get("end_timestamp", bundle.end_timestamp))
        form.addRow(self.end_timestamp)
        self.end_credit = QCheckBox("Show Software Credit on the End Page")
        self.end_credit.setChecked(defaults.get("end_credit", bundle.end_credit))
        form.addRow(self.end_credit)
        for control in (self.cover_title, self.cover_blanks, self.cover_credit, self.cover_timestamp,
                        self.cover_uppercase, *self.cover_fields.values()):
            control.setEnabled(self.cover.isChecked())
            self.cover.toggled.connect(control.setEnabled)
        for control in (self.end_blanks, self.end_note, self.end_timestamp, self.end_credit,
                        self.end_uppercase, *self.end_fields.values()):
            control.setEnabled(self.end.isChecked())
            self.end.toggled.connect(control.setEnabled)
        form = self.section("Room Code", "footer")
        self.print_room_code = QCheckBox("Print the room_code from the center workbook")
        self.print_room_code.setChecked(defaults.get("print_room_code", bundle.print_room_code))
        form.addRow(self.print_room_code)
        self.room_code_position = Choice()
        self.room_code_position.addItems(ROOM_CODE_POSITIONS)
        self.room_code_position.setCurrentText(defaults.get("room_code_position", bundle.room_code_position))
        form.addRow("Placement", self.room_code_position)
        self.room_code_scope = Choice()
        self.room_code_scope.addItems(ROOM_CODE_SCOPES)
        self.room_code_scope.setCurrentText(defaults.get("room_code_scope", bundle.room_code_scope))
        form.addRow("Print on", self.room_code_scope)
        saved_format = defaults.get("room_code_format", bundle.room_code_format)
        self.room_code_format = QLineEdit(bundle.room_code_format if saved_format == "Room code: {code}" else saved_format)
        self.room_code_format.setMaxLength(100)
        form.addRow("Text Format", self.room_code_format)
        self.room_code_size = DecimalInput()
        self.room_code_size.setDecimals(1)
        self.room_code_size.setRange(6, 24)
        self.room_code_size.setSingleStep(0.5)
        self.room_code_size.setValue(defaults.get("room_code_size", bundle.room_code_size))
        form.addRow("Font Size (pt)", self.room_code_size)
        self.room_code_color = QLineEdit(defaults.get("room_code_color", bundle.room_code_color))
        self.room_code_color.setMaxLength(7)
        color_row = QHBoxLayout()
        color_row.addWidget(self.room_code_color, 1)
        choose_color = button("Choose Color", self.choose_room_color, symbol="palette")
        color_row.addWidget(choose_color)
        form.addRow("Font Color", color_row)
        self.room_code_opacity = DecimalInput()
        self.room_code_opacity.setRange(0, 1)
        self.room_code_opacity.setSingleStep(0.05)
        self.room_code_opacity.setValue(defaults.get("room_code_opacity", bundle.room_code_opacity))
        form.addRow("Opacity", self.room_code_opacity)
        form.addRow(subtitle("Use {code} for the eight-character code. Select a clear margin on the question pages."))
        for control in (self.room_code_position, self.room_code_scope, self.room_code_format, self.room_code_size,
                        self.room_code_color, choose_color, self.room_code_opacity):
            control.setEnabled(self.print_room_code.isChecked())
            self.print_room_code.toggled.connect(control.setEnabled)
        console_line = QHBoxLayout()
        console_title = QLabel("Generation Console")
        console_title.setObjectName("consoleHeading")
        console_line.addWidget(console_title)
        console_line.addStretch()
        self.status = QLabel("Ready")
        self.status.setObjectName("bundleStatus")
        console_line.addWidget(self.status)
        main.addLayout(console_line)
        self.console = QPlainTextEdit(readOnly=True)
        self.console.setObjectName("bundleConsole")
        self.console.setFont(QFontDatabase.systemFont(QFontDatabase.SystemFont.FixedFont))
        self.console.setMinimumHeight(136)
        self.console.setMaximumBlockCount(5000)
        self.console.setPlaceholderText("Room-by-room progress will appear here.")
        main.addWidget(self.console, 2)
        if settings_error:
            self.console.appendPlainText(f"Could not read saved room settings: {settings_error}")
        actions = QHBoxLayout()
        self.generate = button("Generate Room Files", self.start, True, "export")
        self.open_folder = button("Open Output Folder", lambda: QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.output))), symbol="folder")
        self.open_folder.setEnabled(False)
        self.close_button = button("Close", self.reject)
        self.save_button = button("Save Config Defaults", self.save_defaults, symbol="save")
        self.save_button.setToolTip("Save paths, paper size, duplex, cover, end-page and room code settings.")
        for control in (self.generate, self.open_folder, self.save_button, self.close_button):
            actions.addWidget(control)
        actions.addStretch()
        main.addLayout(actions)

    def section(self, title, symbol):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        page = QWidget()
        scroll.setWidget(page)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)
        self.tab_icons.append(symbol)
        self.form.addTab(scroll, icon(symbol), title)
        form = QFormLayout(page)
        form.setContentsMargins(15, 14, 16, 14)
        form.setHorizontalSpacing(15)
        form.setVerticalSpacing(8)
        return form

    def update_tab_icons(self, current):
        for index, symbol in enumerate(self.tab_icons):
            self.form.setTabIcon(index, icon(symbol, light=index == current))

    def page_fields(self, form, title, selected):
        grid = QGridLayout()
        fields = {}
        for index, (key, label) in enumerate(COVER_FIELDS.items()):
            option = QCheckBox(label)
            option.setChecked(key in selected)
            fields[key] = option
            grid.addWidget(option, index // 2, index % 2)
        form.addRow(title, grid)
        return fields

    def path_picker(self, form, label, default, extension=None):
        row = QHBoxLayout()
        entry = QLineEdit(default)
        row.addWidget(entry, 1)
        row.addWidget(button("Browse", lambda: self.choose(entry, extension), symbol="folder"))
        form.addRow(label, row)
        return entry

    def choose(self, field, extension):
        if extension:
            path, _ = QFileDialog.getOpenFileName(self, "Select File", field.text(), extension)
        else:
            path = QFileDialog.getExistingDirectory(self, "Select Folder", field.text())
        if path:
            field.setText(path)

    def spin(self, value, minimum=0, maximum=30):
        control = CountInput()
        control.setRange(minimum, maximum)
        control.setValue(value)
        return control

    def choose_room_color(self):
        color = QColorDialog.getColor(QColor(self.room_code_color.text()), self, "Room Code Color")
        if color.isValid():
            self.room_code_color.setText(color.name())

    def show_sets(self, count):
        for index, (choose, entry) in enumerate(self.set_fields):
            choose.setVisible(index < count)
            entry.setVisible(index < count)

    def options(self):
        return BundleOptions(
            paper=self.paper.currentText(), cover=self.cover.isChecked(), cover_blanks=self.cover_blanks.value(),
            cover_fields=tuple(key for key, check in self.cover_fields.items() if check.isChecked()),
            end=self.end.isChecked(), end_blanks=self.end_blanks.value(), end_note=self.end_note.toPlainText(),
            duplex=self.duplex.isChecked(),
            end_fields=tuple(key for key, check in self.end_fields.items() if check.isChecked()),
            cover_title=self.cover_title.text().strip(), cover_credit=self.cover_credit.isChecked(),
            cover_timestamp=self.cover_timestamp.isChecked(), end_timestamp=self.end_timestamp.isChecked(),
            end_credit=self.end_credit.isChecked(), footer_size=self.footer_size.value(),
            footer_opacity=self.footer_opacity.value(),
            cover_uppercase=self.cover_uppercase.isChecked(), end_uppercase=self.end_uppercase.isChecked(),
            print_room_code=self.print_room_code.isChecked(), room_code_position=self.room_code_position.currentText(),
            room_code_scope=self.room_code_scope.currentText(), room_code_format=self.room_code_format.text().strip(),
            room_code_size=self.room_code_size.value(), room_code_color=self.room_code_color.text().strip(),
            room_code_opacity=self.room_code_opacity.value(),
        )

    def save_defaults(self, *, automatic=False):
        config = {
            **asdict(self.options()), "set_count": self.set_count.value(),
            "output_base": self.output_base.text().strip(), "workbook": self.workbook.text().strip(),
            "set_files": [field.text().strip() for _, field in self.set_fields[:self.set_count.value()]],
        }
        config["cover_fields"] = list(config["cover_fields"])
        config["end_fields"] = list(config["end_fields"])
        if automatic and config == self.settings_snapshot:
            return
        try:
            save_json(self.settings_path, config)
        except OSError as error:
            if automatic:
                self.console.appendPlainText(f"Could not save configuration: {error}")
            else:
                QMessageBox.warning(self, "Could Not Save Configuration", str(error))
            return
        self.settings_snapshot = config
        if not automatic:
            self.console.appendPlainText("Room Print Packs configuration saved.")

    def set_status(self, text, state=""):
        self.status.setText(text)
        self.status.setProperty("state", state)
        self.status.style().unpolish(self.status)
        self.status.style().polish(self.status)

    def start(self):
        workbook = self.workbook.text().strip()
        folder = self.output_base.text().strip()
        sources = [field.text().strip() for _, field in self.set_fields[:self.set_count.value()]]
        if not workbook or not folder or not all(sources):
            QMessageBox.warning(self, "Missing Inputs", "Select the center workbook, every question PDF and the output base folder.")
            return
        if self.generate.text().startswith("Regenerate"):
            result = QMessageBox.question(self, "Regenerate Room Files",
                                          "Regenerate all rows marked DO in a new time-stamped folder?",
                                          QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                          QMessageBox.StandardButton.No)
            if result != QMessageBox.StandardButton.Yes:
                return
        options = self.options()
        self.console.clear()
        self.save_defaults(automatic=True)
        self.set_status("Generating…")
        self.console.appendPlainText("Reading center workbook…")
        for control in (self.form, self.generate, self.close_button, self.open_folder, self.save_button):
            control.setEnabled(False)
        task = BundleTask(lambda progress: build_bundles(
            Path(workbook), [Path(file) for file in sources], Path(folder), options, progress), self)
        self.task = task
        queued = Qt.ConnectionType.QueuedConnection
        task.progress.connect(self.console.appendPlainText, queued)
        task.succeeded.connect(self.completed, queued)
        task.failed.connect(self.failed, queued)
        task.finished.connect(self.worker_finished, queued)
        self.started_at = perf_counter()
        task.start()

    def completed(self, result):
        self.output, done, failed, copies = result
        state = "warning" if failed or not done else "success"
        heading = "Completed with errors" if failed else ("Completed" if done else "No rooms selected")
        unit = "question" if copies == 1 else "questions"
        summary = f"{done} rooms saved  |  {failed} failed  |  {copies} {unit}"
        self.set_status(f"{heading} · {summary}", state)
        self.console.appendPlainText(f"\n{heading.upper()} · {summary}")
        self.log_elapsed()
        if done:
            self.generate.setText("Regenerate Room Files")

    def failed(self, message):
        self.set_status("Stopped · review the console", "error")
        self.console.appendPlainText(f"\nSTOPPED · {message}")
        self.log_elapsed()

    def log_elapsed(self):
        minutes, seconds = divmod(round(perf_counter() - self.started_at, 2), 60)
        self.console.appendPlainText(f"Total time · {int(minutes)}m {seconds:05.2f}s")

    def worker_finished(self):
        self.task.deleteLater()
        self.task = None
        for control in (self.form, self.generate, self.close_button, self.save_button):
            control.setEnabled(True)
        self.open_folder.setEnabled(bool(self.output and self.output.is_dir()))

    def reject(self):
        if not self.task:
            self.save_defaults(automatic=True)
            super().reject()
