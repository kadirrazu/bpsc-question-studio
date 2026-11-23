"""Desktop workflow and common configuration."""

import sys
from pathlib import Path
from time import monotonic

import fitz
from PySide6.QtCore import QDate, QLockFile, QRect, QRegularExpression, QSize, QThread, QTimer, Qt, QUrl, Signal
from PySide6.QtGui import QColor, QDesktopServices, QFontDatabase, QFontMetrics, QImage, QPixmap, QRegularExpressionValidator
from PySide6.QtWidgets import (
    QApplication, QButtonGroup, QCheckBox, QColorDialog, QComboBox, QDialog, QDoubleSpinBox,
    QFileDialog, QFormLayout, QFrame, QGridLayout, QHBoxLayout,
    QHeaderView, QInputDialog, QLabel, QLineEdit, QListView, QListWidget, QListWidgetItem, QMainWindow, QMessageBox, QProgressBar, QPushButton,
    QScrollArea, QStackedWidget, QStyledItemDelegate, QTabWidget, QTableWidget, QTableWidgetItem, QTextBrowser, QVBoxLayout, QWidget,
)

from .access import DEFAULT_EMAIL, initialize_access, matches
from .appearance import THEME, icon
from .controls import Choice, CountInput, DateInput, DecimalInput
from .documents import open_word, validate_template
from .project import (
    EXAM_TYPES, PATH_SETTINGS, QR_FIELDS, VERSION, create_exam, delete_exam, exam_folders,
    load_settings, load_units, mapping_rows, post_name, read_exam, save_json,
)
from .workflow import (
    confirm_zero, export_booklets, generate_mapping, generated_numbers, import_source,
    mapping_files, review_status, save_mapping_exports, save_template,
)

STEPS = (
    ("Exam Details", "details"), ("Question Mapping", "mapping"), ("Set 0 Source", "upload"),
    ("Review Set 0", "review"), ("Generate Sets", "export"),
)


def button(text, action, primary=False, symbol=None, nav=False):
    control = QPushButton(text)
    if primary or nav:
        control.setObjectName("primary" if primary else "navButton")
    if symbol:
        control.setIcon(icon(symbol, light=primary or nav))
        control.setIconSize(QSize(17, 17))
    control.clicked.connect(action)
    return control


def local_path(path):
    QDesktopServices.openUrl(QUrl.fromLocalFile(str(path.resolve())))


def pick(parent, title, extension=None):
    if extension:
        path, _ = QFileDialog.getOpenFileName(parent, title, "", extension)
        return path
    return QFileDialog.getExistingDirectory(parent, title)


def subtitle(text):
    label = QLabel(text)
    label.setObjectName("subtitle")
    label.setWordWrap(True)
    return label


def row(layout, widgets):
    h = QHBoxLayout()
    h.setSpacing(9)
    for widget in widgets:
        h.addWidget(widget)
    h.addStretch()
    layout.addLayout(h)


def card(parent):
    panel = QFrame()
    panel.setObjectName("card")
    contents = QVBoxLayout(panel)
    contents.setContentsMargins(23, 21, 23, 23)
    contents.setSpacing(14)
    parent.addWidget(panel)
    return contents


def style_status(label, text, state="info"):
    label.setText(text)
    label.setProperty("state", state)
    label.style().unpolish(label)
    label.style().polish(label)


class SavedExamDelegate(QStyledItemDelegate):
    def sizeHint(self, option, index):
        hint = super().sizeHint(option, index)
        width = max(64, self.parent().viewport().width() - 44)
        text = index.data(Qt.ItemDataRole.DisplayRole) or ""
        flags = Qt.TextFlag.TextWordWrap | Qt.TextFlag.TextWrapAnywhere
        height = QFontMetrics(option.font).boundingRect(QRect(0, 0, width, 10000), flags, text).height()
        return QSize(0, max(hint.height(), height + 24))


class LoginDialog(QDialog):
    def __init__(self, credentials):
        super().__init__()
        self.credentials = credentials
        self.setWindowTitle("Sign In to BPSC Question Studio")
        self.setWindowIcon(icon())
        self.setMinimumWidth(440)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(22, 20, 22, 20)
        layout.setSpacing(12)
        heading = QLabel("BPSC Question Studio")
        heading.setObjectName("stepTitle")
        layout.addWidget(heading)
        layout.addWidget(subtitle("Sign in with the email and password for this installation."))
        form = QFormLayout()
        self.email = QLineEdit(DEFAULT_EMAIL)
        self.password = QLineEdit()
        self.password.setEchoMode(QLineEdit.EchoMode.Password)
        self.password.setFocus()
        form.addRow("Email", self.email)
        form.addRow("Password", self.password)
        layout.addLayout(form)
        self.message = subtitle("")
        self.message.setObjectName("status")
        self.message.hide()
        layout.addWidget(self.message)
        submit = button("Sign In", self.sign_in, True, "check")
        submit.setDefault(True)
        row(layout, [submit, button("Exit", self.reject)])

    def sign_in(self):
        try:
            if matches(self.credentials, self.password.text(), self.email.text()):
                self.password.clear()
                self.accept()
                return
            message = "The email or password is incorrect."
        except (OSError, ValueError) as error:
            message = str(error)
        self.password.clear()
        self.password.setFocus()
        style_status(self.message, message, "error")
        self.message.show()


class SettingsDialog(QDialog):
    def __init__(self, owner):
        super().__init__(owner)
        self.owner = owner
        self.fields = {}
        self.setWindowTitle("Common Settings")
        self.resize(1020, 550)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(15)
        self.navigation = QGridLayout()
        self.navigation.setSpacing(6)
        self.navigation_buttons = QButtonGroup(self)
        layout.addLayout(self.navigation)
        self.tabs = QStackedWidget()
        self.navigation_buttons.idClicked.connect(self.tabs.setCurrentIndex)
        layout.addWidget(self.tabs)
        locations = self.tab("Locations", "folder")
        self.path_field(locations, "template_root", "Template Base Folder")
        self.path_field(locations, "output_root", "Exam Output Base Folder")
        templates = self.tab("Templates", "templates")
        self.path_field(templates, "header", "First-Page Header", "Word Document (*.docx)")
        self.path_field(templates, "footer", "Last-Page Footer", "Word Document (*.docx)")
        self.check_field(templates, "use_default_header", "Use Default Header When the Field Is Empty")
        self.check_field(templates, "use_default_footer", "Use Default Footer When the Field Is Empty")
        templates.addRow(subtitle(
            "Place DEFAULT_HEADER.docx and DEFAULT_FOOTER.docx in Config/Template under the Template Base Folder. "
            "A selected custom template takes precedence."
        ))
        tags = subtitle(
            "<b>Optional Magic Tags</b><br>"
            "<code>[[SET]]</code> · Set number<br>"
            "<code>[[CODE]]</code> · Code name of the question<br>"
            "<code>[[QR]]</code> · QR image<br><br>"
            "Set the QR image size and position in QR Content. Headers and footers can be saved without tags. Runners accept the two text tags. "
            "Older templates and runner tags still work."
        )
        tags.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        templates.addRow(tags)
        self.choice(templates, "tag_language", "Set Tag Language (Global)", ["Bangla", "English"])
        templates.addRow(subtitle("Applies to set-number tags in templates and to the set number in QR content. "
                                   "Choose the runner’s language separately on the Runner page."))
        angles = [("Normal", 0), ("90° Clockwise", 90), ("270° Clockwise", 270)]
        for key, label in (("set", "Set Number"), ("code", "Code Name")):
            self.choice(templates, f"{key}_rotation", f"{label} Rotation", angles)
            self.choice(templates, f"{key}_horizontal", f"{label} Horizontal Alignment",
                        ["Template", "Left", "Center", "Right"])
            self.choice(templates, f"{key}_vertical", f"{label} Cell Alignment",
                        ["Template", "Top", "Center", "Bottom"])
            rotation = self.fields[f"{key}_rotation"]
            for control in (self.fields[f"{key}_horizontal"], self.fields[f"{key}_vertical"]):
                rotation.currentIndexChanged.connect(
                    lambda _, field=control, angle=rotation: field.setEnabled(bool(angle.currentData()))
                )
                control.setEnabled(bool(rotation.currentData()))
        templates.addRow(subtitle(
            "Alignment overrides apply to rotated tags. Horizontal alignment moves the paragraph. "
            "Cell alignment moves the contents of that cell, including any other tags it contains."
        ))
        pages = self.tab("Question Pages", "page-layout")
        self.choice(pages, "paper", "Paper Size", ["Legal", "A4"])
        margins = QHBoxLayout()
        for side in ("top", "bottom", "left", "right"):
            group = QVBoxLayout()
            group.addWidget(QLabel(side.title()))
            group.addWidget(self.number(side, 0.1, 4))
            margins.addLayout(group)
        pages.addRow("Margins (in)", margins)
        self.number_field(pages, "header_gap", "Header Gap (in)", 0, 3)
        self.number_field(pages, "footer_gap", "Footer Gap (in)", 0, 3)
        self.choice(pages, "serial_suffix", "Question Serial Suffix",
                    [("Dot (.)", "."), ("Closing Parenthesis", ")"), ("Dari (।)", "।"),
                     ("None", ""), ("Keep Original", "Original")])
        self.font_field(pages, "font", "Generated Text Font")
        self.number_field(pages, "font_size", "Generated Text Size (pt)", 5, 32, step=0.5)
        mapping = self.tab("Mapping Export", "mapping")
        self.choice(mapping, "mapping_paper", "Paper Size", ["Legal", "A4", "Letter"])
        self.choice(mapping, "mapping_orientation", "Orientation", ["Portrait", "Landscape"])
        pdf_margins = QHBoxLayout()
        for side in ("top", "right", "bottom", "left"):
            group = QVBoxLayout()
            group.addWidget(QLabel(side.title()))
            group.addWidget(self.number(f"mapping_margin_{side}", 0, 2))
            pdf_margins.addLayout(group)
        mapping.addRow("PDF Margins (in)", pdf_margins)
        mapping.addRow(subtitle("Paper size and orientation apply to Excel and PDF. PDF margins apply only to the mapping PDF."))
        qr = self.tab("QR Content", "qr")
        qr.addRow(subtitle("Choose what each set’s QR code contains. Clear all fields to omit the QR image."))
        qr_options = QGridLayout()
        self.qr_fields = {}
        for index, (key, label) in enumerate(QR_FIELDS.items()):
            box = QCheckBox(label)
            box.setChecked(key in owner.settings["qr_fields"])
            self.qr_fields[key] = box
            qr_options.addWidget(box, index // 2, index % 2)
        qr.addRow(qr_options)
        self.text_field(qr, "credit", "Custom QR Text")
        self.fields["credit"].setMaxLength(500)
        self.number_field(qr, "qr_size", "QR Width (in)", 0.35, 2)
        self.choice(qr, "qr_horizontal", "Horizontal Alignment", ["Template", "Left", "Center", "Right"])
        self.choice(qr, "qr_vertical", "Cell Alignment", ["Template", "Top", "Center", "Bottom"])
        self.number_field(qr, "qr_inset", "Horizontal Inset (in)", -1, 1)
        self.number_field(qr, "qr_gap", "Vertical Gap (in)", -1, 1)
        qr.addRow(subtitle(
            "Choose Template to follow the tag’s paragraph and cell. Put the QR tag on its own line "
            "to apply an inset or gap. Use a negative value to move in the opposite direction."
        ))
        runner = self.tab("Runner", "runner")
        self.text_field(runner, "runner_text", "Runner Text")
        self.fields["runner_text"].setPlaceholderText("BPSC  •  Set [[SET]]")
        self.choice(runner, "runner_tag_language", "Set Tag Language (Runner)", ["Bangla", "English"])
        runner.addRow(subtitle("The language sets the digits in [[SET]]. Enter the rest of the runner text as you want it printed."))
        sides = QHBoxLayout()
        self.sides = {}
        for side in ("top", "bottom", "left", "right"):
            control = QCheckBox(side.title())
            control.setChecked(side in owner.settings["sides"])
            self.sides[side] = control
            sides.addWidget(control)
        runner.addRow("Print on", sides)
        self.number_field(runner, "runner_offset", "Distance from Edge (in)", 0, 3)
        self.number_field(runner, "runner_interval", "Gap Between Repeats (in)", 0.01, 12)
        self.appearance_fields(runner, "runner")
        footer = self.tab("Footer Text", "footer")
        footer.addRow(subtitle("Style the unique booklet code and page label printed on every page."))
        self.choice(footer, "footer_language", "Set and Page Language", ["English", "Bangla"])
        self.appearance_fields(footer, "footer")
        footer.addRow(subtitle("Unique Alphanumeric Code for Each Set (left side of the footer)"))
        self.check_field(footer, "print_code_enabled", "Print the 8-character unique code")
        self.appearance_fields(footer, "print_code", show_case=False)
        engine = self.tab("Renderer", "settings")
        self.path_field(engine, "renderer", "LibreOffice Program", "LibreOffice (soffice*);;All Files (*)")
        engine.addRow(subtitle("The app looks for an installed copy of LibreOffice. Browse to its executable if it is installed in a different location."))
        actions = QHBoxLayout()
        actions.addStretch()
        actions.addWidget(button("Cancel", self.reject))
        actions.addWidget(button("Save Settings", self.save, True, "save"))
        layout.addLayout(actions)

    def tab(self, title, symbol):
        page = QScrollArea()
        page.setFrameShape(QFrame.Shape.NoFrame)
        page.setWidgetResizable(True)
        content = QWidget()
        layout = QVBoxLayout(content)
        layout.setContentsMargins(6, 15, 7, 5)
        form = QFormLayout()
        form.setHorizontalSpacing(14)
        form.setVerticalSpacing(11)
        layout.addLayout(form)
        layout.addStretch()
        page.setWidget(content)
        index = self.tabs.addWidget(page)
        control = QPushButton(title, icon=icon(symbol), checkable=True)
        control.setObjectName("settingsSection")
        self.navigation_buttons.addButton(control, index)
        self.navigation.addWidget(control, index // 4, index % 4)
        if index == 0:
            control.setChecked(True)
        return form

    def choice(self, layout, key, title, choices):
        control = Choice()
        for choice in choices:
            display, value = choice if isinstance(choice, tuple) else (choice, choice)
            control.addItem(display, value)
        control.setCurrentIndex(control.findData(self.owner.settings[key]))
        self.fields[key] = control
        layout.addRow(title, control)

    def text_field(self, layout, key, title):
        field = QLineEdit(str(self.owner.settings[key]))
        self.fields[key] = field
        layout.addRow(title, field)

    def check_field(self, layout, key, title):
        field = QCheckBox(title)
        field.setChecked(self.owner.settings[key])
        self.fields[key] = field
        layout.addRow(field)

    def font_field(self, layout, key, label):
        family = QLineEdit(self.owner.settings[key])
        self.fields[key] = family
        controls = QHBoxLayout()
        controls.addWidget(family, 1)

        def choose_font():
            current = family.text().strip() or self.fields["font"].text().strip()
            families = sorted({*QFontDatabase.families(), current} - {""}, key=str.casefold)
            selected, ok = QInputDialog.getItem(self, "Choose Font", "Font Family",
                                                families, families.index(current) if current else 0, False)
            if ok:
                family.setText(selected)

        controls.addWidget(button("Choose Font", choose_font, symbol="font"))
        layout.addRow(label, controls)
        return family

    def appearance_fields(self, layout, kind, *, show_case=True):
        family = self.font_field(layout, f"{kind}_font", "Font")
        size = self.number(f"{kind}_size", 0 if kind == "footer" else 5, 32, step=0.5)
        if kind == "footer":
            def update_inherited_font(name):
                family.setPlaceholderText(f"Use Generated Text Font ({name})")

            update_inherited_font(self.fields["font"].text())
            self.fields["font"].textChanged.connect(update_inherited_font)
            size.setSpecialValueText("Use Generated Text Size")
        layout.addRow("Font Size (pt)", size)
        color_key = f"{kind}_color"
        color = QLineEdit(self.owner.settings[color_key])
        color.setMaxLength(7)
        color.setValidator(QRegularExpressionValidator(QRegularExpression("#[0-9A-Fa-f]{6}"), color))
        self.fields[color_key] = color
        controls = QHBoxLayout()
        controls.addWidget(color, 1)

        def choose_color():
            selected = QColorDialog.getColor(QColor(color.text()), self)
            if selected.isValid():
                color.setText(selected.name().upper())

        controls.addWidget(button("Choose Color", choose_color, symbol="palette"))
        layout.addRow("Text Color", controls)
        self.number_field(layout, f"{kind}_opacity", "Opacity (0–1)", 0, 1)
        if show_case:
            self.check_field(layout, f"{kind}_uppercase", "Use All Capital Letters")

    def number(self, key, low, high, *, step=0.05):
        field = DecimalInput()
        field.setRange(low, high)
        field.setDecimals(2)
        field.setSingleStep(step)
        field.setValue(self.owner.settings[key])
        self.fields[key] = field
        return field

    def number_field(self, layout, key, title, low, high, *, step=0.05):
        layout.addRow(title, self.number(key, low, high, step=step))

    def path_field(self, layout, key, title, extension=None):
        controls = QHBoxLayout()
        field = QLineEdit(str(self.owner.settings[key]))
        field.setReadOnly(True)
        self.fields[key] = field
        controls.addWidget(field, 1)

        def choose():
            path = pick(self, title, extension)
            if path:
                field.setText(path)

        controls.addWidget(button("Browse", choose, symbol="folder"))
        if extension:
            controls.addWidget(button("Clear", field.clear))
        layout.addRow(title, controls)

    def save(self):
        settings = {}
        for key, field in self.fields.items():
            if isinstance(field, QDoubleSpinBox):
                settings[key] = field.value()
            elif isinstance(field, QComboBox):
                settings[key] = field.currentData()
            elif isinstance(field, QCheckBox):
                settings[key] = field.isChecked()
            else:
                settings[key] = field.text().strip()
        if not settings["font"]:
            QMessageBox.warning(self, "Generated Text Font", "Enter a font name or use Choose Font.")
            return
        for kind in ("runner", "footer", "print_code"):
            key = f"{kind}_color"
            if not self.fields[key].hasAcceptableInput():
                QMessageBox.warning(self, "Text Color", "Enter a six-digit hex color starting with # or use Choose Color.")
                return
            settings[key] = settings[key].upper()
        settings["sides"] = [side for side, box in self.sides.items() if box.isChecked()]
        settings["qr_fields"] = [key for key, box in self.qr_fields.items() if box.isChecked()]
        if "text" in settings["qr_fields"] and not settings["credit"]:
            QMessageBox.warning(self, "QR Content", "Enter custom QR text or clear the Custom Text checkbox.")
            return
        try:
            root = Path(settings["template_root"])
            for kind in ("header", "footer"):
                if settings[kind]:
                    validate_template(open_word(settings[kind]))
            (root / "Config" / "Template").mkdir(parents=True, exist_ok=True)
            Path(settings["output_root"]).mkdir(parents=True, exist_ok=True)
            for kind in ("header", "footer"):
                if settings[kind]:
                    settings[kind] = save_template(root, kind, settings[kind])
            for key in PATH_SETTINGS:
                if settings[key]:
                    settings[key] = str(self.owner.app_folder / settings[key])
            portable = {**settings, "settings_version": VERSION}
            for key in PATH_SETTINGS:
                if portable[key]:
                    path = Path(portable[key])
                    if path.is_relative_to(self.owner.app_folder):
                        portable[key] = str(path.relative_to(self.owner.app_folder))
            save_json(self.owner.settings_path, portable)
        except (ValueError, OSError) as error:
            QMessageBox.warning(self, "Could Not Save Settings", str(error))
            return
        self.owner.settings = settings
        self.accept()


class Task(QThread):
    progress = Signal(int, int, str)
    succeeded = Signal(object)
    failed = Signal(str)

    def __init__(self, operation, parent, report=False):
        super().__init__(parent)
        self.operation = operation
        self.report = report

    def run(self):
        try:
            result = self.operation(self.progress.emit) if self.report else self.operation()
            self.succeeded.emit(result)
        except Exception as error:
            self.failed.emit(str(error))


class Preview(QDialog):
    def __init__(self, parent, path):
        super().__init__(parent)
        self.document = fitz.open(path)
        self.index = 0
        self.setWindowTitle("PDF Preview")
        self.resize(760, 800)
        v = QVBoxLayout(self)
        actions = QHBoxLayout()
        actions.addWidget(button("Previous Page", lambda: self.show_page(self.index - 1), symbol="prev"))
        self.position = QLabel()
        actions.addWidget(self.position)
        actions.addWidget(button("Next Page", lambda: self.show_page(self.index + 1), symbol="next"))
        actions.addStretch()
        actions.addWidget(button("Open in PDF Viewer", lambda: local_path(path), symbol="pdf"))
        v.addLayout(actions)
        scroll = QScrollArea()
        self.image = QLabel()
        scroll.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        scroll.setWidget(self.image)
        v.addWidget(scroll)
        self.show_page(0)

    def show_page(self, index):
        if 0 <= index < len(self.document):
            self.index = index
            self.position.setText(f"{index + 1} / {len(self.document)}")
            pixmap = self.document[index].get_pixmap(matrix=fitz.Matrix(1.03, 1.03), alpha=False)
            image = QImage(pixmap.samples, pixmap.width, pixmap.height, pixmap.stride,
                           QImage.Format.Format_RGB888).copy()
            self.image.setPixmap(QPixmap.fromImage(image))
            self.image.adjustSize()

    def done(self, result):
        self.document.close()
        super().done(result)


class Window(QMainWindow):
    def __init__(self, app_folder, credentials=None):
        super().__init__()
        self.app_folder = app_folder
        self.credentials = credentials or initialize_access(app_folder)
        self.units = load_units(app_folder / "Config" / "units.json")
        self.settings_path = app_folder / "Config" / "settings.json"
        self.settings = load_settings(self.settings_path)
        for key in PATH_SETTINGS:
            if self.settings[key]:
                self.settings[key] = str(app_folder / self.settings[key])
        self.settings["template_root"] = self.settings["template_root"] or str(app_folder)
        self.settings["output_root"] = self.settings["output_root"] or str(app_folder / "Exams")
        self.folder = self.exam = self.task = None
        self.setWindowTitle("BPSC Question Studio")
        self.setWindowIcon(icon())
        self.resize(1190, 730)
        self.setMinimumSize(980, 620)
        shell = QWidget()
        layout = QHBoxLayout(shell)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)
        self.setCentralWidget(shell)
        self.sidebar_panel = self.sidebar()
        layout.addWidget(self.sidebar_panel)

        canvas = QVBoxLayout()
        canvas.setContentsMargins(23, 22, 23, 18)
        canvas.setSpacing(14)
        layout.addLayout(canvas, 1)
        canvas.addWidget(self.banner())
        self.tabs = QTabWidget()
        self.tabs.tabBar().setUsesScrollButtons(True)
        self.tabs.tabBar().setExpanding(False)
        canvas.addWidget(self.tabs, 1)
        self.exam_page()
        self.mapping_page()
        self.source_page()
        self.review_page()
        self.output_page()
        self.status = QLabel("Create a new exam or open a saved exam.")
        self.status.setObjectName("status")
        self.status.setTextFormat(Qt.TextFormat.PlainText)
        self.status.setWordWrap(True)
        canvas.addWidget(self.status)
        self.progress_row = QWidget()
        self.progress_row.setObjectName("fieldGroup")
        progress_layout = QHBoxLayout(self.progress_row)
        progress_layout.setContentsMargins(1, 0, 1, 0)
        self.progress_bar = QProgressBar()
        self.progress_bar.setAccessibleName("Question set generation progress")
        progress_layout.addWidget(self.progress_bar, 1)
        self.elapsed = QLabel()
        progress_layout.addWidget(self.elapsed)
        self.progress_timer = QTimer(self)
        self.progress_timer.setInterval(500)
        self.progress_timer.timeout.connect(self.update_elapsed)
        self.progress_row.hide()
        canvas.addWidget(self.progress_row)
        self.update_saved()
        self.refresh()

    def sidebar(self):
        sidebar = QWidget()
        sidebar.setObjectName("sidebar")
        sidebar.setFixedWidth(224)
        nav = QVBoxLayout(sidebar)
        nav.setSpacing(10)
        nav.setContentsMargins(14, 27, 14, 18)
        heading = QLabel("Set Generation · Saved Exams")
        heading.setObjectName("sideTitle")
        nav.addWidget(heading)
        nav.addWidget(button("New Exam", self.new_exam, symbol="plus", nav=True))
        self.saved = QListWidget()
        self.saved.setObjectName("examList")
        self.saved.setAccessibleName("Saved Exams")
        self.saved.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.saved.setWordWrap(True)
        self.saved.setTextElideMode(Qt.TextElideMode.ElideNone)
        self.saved.setResizeMode(QListView.ResizeMode.Adjust)
        self.saved.setItemDelegate(SavedExamDelegate(self.saved))
        self.saved.itemActivated.connect(self.open_exam)
        nav.addWidget(self.saved, 1)
        self.open_saved = button("Open Selected Exam", lambda: self.open_exam(self.saved.currentItem()), symbol="folder", nav=True)
        self.delete_saved = button("Delete Selected Exam", self.remove_exam, symbol="delete", nav=True)
        self.delete_saved.setProperty("destructive", True)

        def selection_changed(item, _previous):
            self.open_saved.setEnabled(item is not None)
            self.delete_saved.setEnabled(item is not None)

        self.saved.currentItemChanged.connect(selection_changed)
        nav.addWidget(self.open_saved)
        nav.addWidget(self.delete_saved)
        nav.addWidget(button("Common Settings", self.configure, symbol="settings", nav=True))
        nav.addWidget(button("Help", self.help, symbol="help", nav=True))
        nav.addWidget(button("About", self.about, symbol="about", nav=True))
        modules = QLabel("Roomwise File Generation")
        modules.setObjectName("moduleTitle")
        nav.addWidget(modules)
        nav.addWidget(button("Center File Wizard", self.center_file_wizard, symbol="mapping", nav=True))
        nav.addWidget(button("Room Print Packs", self.room_bundles, symbol="pdf", nav=True))
        return sidebar

    def center_file_wizard(self):
        from .center_wizard import CenterFileWizard
        wizard = CenterFileWizard(self)
        if wizard.exec() and wizard.output:
            self.room_bundles(workbook_path=wizard.output)

    def room_bundles(self, *, workbook_path=None):
        from .room_dialog import RoomBundlesDialog
        set_files = None
        if self.exam and self.folder:
            out = self.folder / "GENERATED_SETS"
            saved = self.exam.get("generated", {})
            set_files = []
            for number in range(1, self.exam["sets"] + 1):
                names = saved.get(str(number), [])
                pdf = next((out / name for name in names
                            if Path(name).suffix.lower() == ".pdf" and (out / name).is_file()), None)
                set_files.append(str(pdf) if pdf else "")
        RoomBundlesDialog(self.app_folder, self, set_files=set_files, workbook_path=workbook_path).exec()

    def banner(self):
        hero = QWidget()
        hero.setObjectName("hero")
        line = QHBoxLayout(hero)
        line.setContentsMargins(21, 16, 21, 16)
        line.setSpacing(13)
        logo = QLabel()
        logo.setPixmap(icon().pixmap(QSize(51, 51)))
        line.addWidget(logo)
        text = QVBoxLayout()
        text.setSpacing(1)
        title = QLabel("BPSC Question Studio")
        title.setObjectName("title")
        text.addWidget(title)
        self.context = QLabel()
        self.context.setObjectName("heroContext")
        self.context.setTextFormat(Qt.TextFormat.PlainText)
        self.context.setWordWrap(True)
        text.addWidget(self.context)
        line.addLayout(text, 1)
        return hero

    def page(self, description):
        index = self.tabs.count()
        title, symbol = STEPS[index]
        scroll = QScrollArea()
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setWidgetResizable(True)
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 11, 0, 4)
        contents = card(layout)
        heading = QLabel(title)
        heading.setObjectName("stepTitle")
        contents.addWidget(heading)
        contents.addWidget(subtitle(description))
        layout.addStretch()
        scroll.setWidget(panel)
        self.tabs.addTab(scroll, icon(symbol), title)
        return contents

    def exam_page(self):
        contents = self.page("Enter the exam details and choose how many question sets to create.")
        grid = QGridLayout()
        grid.setHorizontalSpacing(22)
        grid.setVerticalSpacing(7)
        grid.setColumnStretch(0, 1)
        grid.setColumnStretch(1, 1)
        self.unit = Choice()
        self.reset_units()
        self.inputs = {}
        fields = (
            ("post_code", "Post Code (optional)", "260015"),
            ("title", "Post Name", "Assistant Registrar"),
            ("post_grade", "Post Grade (optional)", "9"),
            ("code_name", "Code Name of the Question", "ইছামতি"),
        )
        entries = [("Unit", self.unit)]
        for key, title, hint in fields:
            field = QLineEdit()
            field.setPlaceholderText(hint)
            if key == "post_code":
                field.setMaxLength(6)
                field.setValidator(QRegularExpressionValidator(QRegularExpression("[0-9]{0,6}"), field))
            elif key == "post_grade":
                field.setValidator(QRegularExpressionValidator(QRegularExpression("[0-9]*"), field))
            self.inputs[key] = field
            entries.append((title, field))
        self.exam_type = Choice()
        self.exam_type.addItems(list(EXAM_TYPES))
        self.date = DateInput(QDate.currentDate())
        self.date.setDisplayFormat("dd MMM yyyy")
        self.date.setCalendarPopup(True)
        self.date.setAccessibleName("Exam Date")
        self.count = CountInput()
        self.count.setRange(1, 500)
        self.count.setValue(100)
        self.sets = CountInput()
        self.sets.setRange(1, 25)
        self.sets.setValue(4)
        entries.extend([
            ("Exam Type", self.exam_type), ("Exam Date", self.date),
        ])
        for index, (title, field) in enumerate(entries):
            y, x = (index // 2) * 2, index % 2
            label = QLabel(title)
            label.setBuddy(field)
            grid.addWidget(label, y, x)
            grid.addWidget(field, y + 1, x)
        counts = QGridLayout()
        counts.setHorizontalSpacing(12)
        counts.setVerticalSpacing(grid.verticalSpacing())
        for index, (text, field) in enumerate((("Question Count", self.count), ("Total Set Count", self.sets))):
            label = QLabel(text)
            label.setBuddy(field)
            counts.addWidget(label, 0, index)
            counts.addWidget(field, 1, index)
            counts.setColumnStretch(index, 1)
        grid.addLayout(counts, 6, 1, 2, 1)
        contents.addLayout(grid)
        self.create_button = button("Create Exam", self.create_exam, True, "plus")
        self.folder_button = button("Open Exam Folder", lambda: local_path(self.folder), symbol="folder")
        row(contents, [self.create_button, self.folder_button])

    def reset_units(self):
        self.unit.clear()
        self.unit.addItem("Select Unit", "")
        for unit in self.units:
            self.unit.addItem(unit, unit)

    def mapping_page(self):
        contents = self.page("Assign each question a new position and save the mapping for printing.")
        self.mapping_info = subtitle("")
        self.mapping_info.setObjectName("status")
        contents.addWidget(self.mapping_info)
        self.make_map = button("Generate Mapping", self.generate_mapping, True, "mapping")
        self.map_view = button("View Mapping", self.view_mapping, symbol="review")
        row(contents, [self.make_map, self.map_view])
        separator = QFrame()
        separator.setObjectName("separator")
        separator.setFixedHeight(1)
        contents.addWidget(separator)
        self.xlsx = button("Open Excel", lambda: self.open_mapping(0), symbol="mapping")
        self.pdf_map = button("Open Mapping PDF", lambda: self.open_mapping(1), symbol="pdf")
        self.reexport = button("Save Mapping Files", self.save_mapping, symbol="export")
        row(contents, [self.xlsx, self.pdf_map, self.reexport])

    def source_page(self):
        contents = self.page("Select the Word document with the original question table.")
        self.source_label = subtitle("")
        self.source_label.setObjectName("status")
        self.source_label.setTextFormat(Qt.TextFormat.PlainText)
        self.source_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
        contents.addWidget(self.source_label)
        self.source_button = button("Choose Set 0 DOCX", self.choose_source, True, "upload")
        row(contents, [self.source_button])

    def review_page(self):
        contents = self.page("Generate the Set 0 PDF, check its layout, then confirm it to unlock the other sets.")
        self.zero_state = subtitle("")
        self.zero_state.setObjectName("status")
        contents.addWidget(self.zero_state)
        self.make_zero = button("Generate Set 0 PDF", self.generate_zero, True, "export")
        self.open_zero = button("Open Set 0 PDF", self.preview_zero, symbol="pdf")
        self.confirm_button = button("Confirm Set 0", self.approve_zero, symbol="check")
        row(contents, [self.make_zero, self.open_zero, self.confirm_button])

    def output_page(self):
        contents = self.page("Choose the sets and file format. Print saved PDFs from your PDF viewer.")
        self.generated_state = subtitle("")
        self.generated_state.setObjectName("status")
        contents.addWidget(self.generated_state)
        self.select_set = Choice()
        self.select_set.setAccessibleName("Set Selection")
        self.format = Choice()
        self.format.setAccessibleName("File Format")
        self.format.addItems(["DOCX and PDF", "DOCX Only", "PDF Only"])
        self.format.setCurrentText("PDF Only")
        row(contents, [self.select_set, self.format])
        self.export_button = button("Generate Sets", self.export, True, "export")
        self.output_folder = button("Open Generated Files", self.open_output, symbol="folder")
        row(contents, [self.export_button, self.output_folder])

    def help(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("Help")
        dialog.resize(650, 460)
        layout = QVBoxLayout(dialog)
        heading = QLabel("How to Prepare an Exam")
        heading.setObjectName("stepTitle")
        layout.addWidget(heading)
        instructions = QTextBrowser()
        instructions.setOpenExternalLinks(False)
        instructions.setHtml(
            "<ol>"
            "<li><b>Create an exam.</b> Select the unit and enter the post name and the question’s code name. "
            "Set the question count and the number of sets you want. Enter 4 for Set 1 through Set 4.</li>"
            "<li><b>Create the mapping.</b> Open Question Mapping and click Generate Mapping. "
            "The app saves both Excel and PDF copies. Use View Mapping to check the order.</li>"
            "<li><b>Choose Set 0.</b> In Set 0 Source, upload a Word file with two columns. "
            "Put serial numbers in the first column and the questions in the second.</li>"
            "<li><b>Review Set 0.</b> Generate the PDF, open it and check the layout. "
            "Click Confirm Set 0 when it is ready.</li>"
            "<li><b>Generate the other sets.</b> Choose the sets and file format, then click Generate Sets. "
            "Open the saved PDF in your PDF viewer to print it.</li></ol>"
            "<p>Use Common Settings to change the template, margins, QR content or output folder. "
            "If question PDF generation cannot find LibreOffice, install it or select its executable under Renderer.</p>"
        )
        layout.addWidget(instructions, 1)
        row(layout, [button("Close", dialog.accept)])
        dialog.exec()

    def remove_exam(self):
        item = self.saved.currentItem()
        if item is None:
            return
        folder = item.data(Qt.ItemDataRole.UserRole)
        try:
            exam = read_exam(folder)
            password, accepted = QInputDialog.getText(
                self, "Delete Exam",
                f"Delete {post_name(exam)} and its source, mapping and generated files? "
                "This cannot be undone. Enter the access password to delete it.",
                QLineEdit.EchoMode.Password,
            )
            if not accepted:
                return
            if not matches(self.credentials, password):
                self.error("The access password is incorrect.")
                return
            delete_exam(Path(self.settings["output_root"]), folder)
            if self.folder == folder:
                self.new_exam()
            self.update_saved()
            style_status(self.status, "Exam deleted.", "success")
        except (ValueError, OSError) as error:
            self.error(error)

    def about(self):
        dialog = QMessageBox(self)
        dialog.setWindowTitle("About")
        dialog.setIconPixmap(icon().pixmap(QSize(62, 62)))
        dialog.setText(f"<b>BPSC Question Studio</b><br>Version {VERSION}")
        dialog.setInformativeText("Development and Maintenance<br><b>Md. Abdul Kadir</b><br>Programmer, BPSC")
        dialog.exec()

    def configure(self):
        before = self.settings["output_root"]
        if SettingsDialog(self).exec():
            if before != self.settings["output_root"]:
                self.new_exam()
                self.update_saved()
            style_status(self.status, "Common settings saved.", "success")
            self.refresh()

    def update_saved(self):
        self.saved.clear()
        self.open_saved.setEnabled(False)
        self.delete_saved.setEnabled(False)
        for folder in exam_folders(Path(self.settings["output_root"])):
            try:
                exam = read_exam(folder)
            except (OSError, ValueError):
                continue
            lines = [exam["title"]]
            grade = exam.get("post_grade")
            if grade not in (None, ""):
                lines.append(f"Post Grade: {grade}")
            lines.append(f"{exam['unit']}  ·  {exam['date']}")
            item = QListWidgetItem("\n".join(lines))
            item.setData(Qt.ItemDataRole.UserRole, folder)
            item.setToolTip(f"{item.text()}\n{folder}")
            self.saved.addItem(item)
            if folder == self.folder:
                self.saved.setCurrentItem(item)

    def refresh(self, go=None):
        exam = self.exam or {}
        mapped = bool(exam.get("mapping"))
        source = exam.get("source", "")
        source_exists = bool(source and (self.folder / "SOURCE" / "ZERO_SET.docx").is_file())
        files = mapping_files(self.folder, exam) if mapped else []
        mapping_complete = bool(files and all(path.is_file() for path in files))
        ready = mapping_complete and source_exists
        approved, pdf = review_status(self.folder, self.exam, self.settings) if ready else (False, None)
        completed_sets = len(generated_numbers(self.folder, exam)) if approved else 0
        done = approved and completed_sets == exam["sets"]
        exists = bool(exam)
        context = "Question mapping, booklet review and set generation"
        if exists:
            details = [post_name(exam)]
            if exam.get("post_code"):
                details.append(f"Post Code: {exam['post_code']}")
            if exam.get("unit"):
                details.append(exam["unit"])
            details.extend((exam.get("exam_type", "NON CADRE").upper(), exam["date"]))
            context = "  ·  ".join(details)
        self.context.setText(context)
        self.create_button.setVisible(not exists)
        self.folder_button.setVisible(exists)
        for field in (*self.inputs.values(), self.unit, self.date, self.exam_type, self.count, self.sets):
            field.setEnabled(not exists)
        self.make_map.setText("Regenerate Mapping" if mapped else "Generate Mapping")
        for control in (self.map_view, self.reexport):
            control.setEnabled(mapped)
        self.xlsx.setEnabled(bool(files and files[0].exists()))
        self.pdf_map.setEnabled(bool(files and files[1].exists()))
        if mapped:
            summary = f"{exam['count']} questions · {exam['sets']} sets · "
            style_status(self.mapping_info,
                         summary + ("Excel and PDF saved" if mapping_complete else "Save the Excel and PDF files"),
                         "success" if mapping_complete else "warning")
        else:
            style_status(self.mapping_info, "Each row gives the new positions for one Set 0 question.")
        if source_exists:
            validated = self.folder / "SOURCE" / "ZERO_SET.docx"
            style_status(self.source_label,
                         f"Validated · {exam['count']} questions\nOriginal file: {source}\nSaved copy: {validated}",
                         "success")
        else:
            style_status(self.source_label, "Use a two-column table with serial numbers and questions.")
        self.make_zero.setEnabled(ready)
        self.open_zero.setEnabled(pdf is not None)
        self.confirm_button.setEnabled(pdf is not None and not approved)
        style_status(self.zero_state,
                     "Set 0 confirmed. You can generate the other sets." if approved else
                     "Open the saved PDF, then confirm Set 0." if pdf else
                     "Generate Set 0 again. The inputs changed or a review file is missing." if exam.get("review") else
                     "Generate Set 0 to begin review.",
                     "success" if approved else "info")
        self.export_button.setText("Regenerate Sets" if done else "Generate Sets")
        if approved:
            missing = exam["sets"] - completed_sets
            message = f"{completed_sets} of {exam['sets']} sets generated."
            if missing and exam.get("generated"):
                message += " Generate the missing sets to complete this step."
            style_status(self.generated_state, message, "success" if done else "info")
        else:
            style_status(self.generated_state, "Confirm Set 0 to generate the other sets.")
        self.output_folder.setEnabled(exists and (self.folder / "GENERATED_SETS").is_dir())
        selection = self.select_set.currentData()
        self.select_set.clear()
        self.select_set.addItem("All New Sets", None)
        for number in range(1, exam.get("sets", 0) + 1):
            self.select_set.addItem(f"Set {number}", number)
        old_index = self.select_set.findData(selection)
        if old_index >= 0:
            self.select_set.setCurrentIndex(old_index)
        for index, (enabled, completed) in enumerate(zip(
            (True, exists, mapping_complete, ready, approved),
            (exists, mapping_complete, source_exists, approved, done),
        )):
            title, symbol = STEPS[index]
            self.tabs.setTabEnabled(index, enabled)
            self.tabs.setTabIcon(index, icon("check" if completed else symbol))
            self.tabs.setTabToolTip(index, f"{title} · {'Completed' if completed else 'Available' if enabled else 'Locked'}")
        if go is not None and self.tabs.isTabEnabled(go):
            self.tabs.setCurrentIndex(go)

    def choose_source(self):
        filename = pick(self, "Choose Set 0 DOCX", "Word Document (*.docx)")
        if not filename:
            return
        try:
            self.exam = import_source(self.folder, self.exam, filename)
            style_status(self.status, f"Set 0 validated · {self.exam['count']} questions", "success")
            self.refresh(3)
        except (ValueError, OSError) as error:
            self.error(error)

    def new_exam(self):
        self.folder = self.exam = None
        self.saved.setCurrentItem(None)
        self.reset_units()
        for field in self.inputs.values():
            field.clear()
        self.inputs["post_code"].setMaxLength(6)
        self.exam_type.setCurrentIndex(0)
        self.date.setDate(QDate.currentDate())
        self.count.setValue(100)
        self.sets.setValue(4)
        style_status(self.status, "Enter the details for a new exam.")
        self.refresh(0)

    def open_exam(self, item):
        if item is None:
            return
        try:
            folder = item.data(Qt.ItemDataRole.UserRole)
            exam = read_exam(folder)
        except (ValueError, OSError) as error:
            self.error(error)
            return
        self.folder, self.exam = folder, exam
        post_code = str(exam.get("post_code") or "")
        self.inputs["post_code"].setMaxLength(max(6, len(post_code)))
        for key, field in self.inputs.items():
            value = exam.get(key)
            field.setText("" if value is None else str(value))
        self.reset_units()
        if self.unit.findData(exam["unit"]) < 0:
            self.unit.addItem(exam["unit"], exam["unit"])
        self.unit.setCurrentIndex(self.unit.findData(exam["unit"]))
        self.date.setDate(QDate.fromString(exam["date"], "yyyy-MM-dd"))
        self.exam_type.setCurrentText(exam.get("exam_type", "NON CADRE").upper())
        self.count.setValue(exam["count"])
        self.sets.setValue(exam["sets"])
        style_status(self.status, f"Opened {post_name(exam)}")
        go = 1
        if exam["mapping"] and all(path.is_file() for path in mapping_files(folder, exam)):
            go = 2
            if exam.get("source") and (folder / "SOURCE" / "ZERO_SET.docx").is_file():
                go = 4 if review_status(folder, exam, self.settings)[0] else 3
        self.refresh(go)

    def create_exam(self):
        try:
            self.folder, self.exam = create_exam(
                Path(self.settings["output_root"]), self.unit.currentData(),
                self.inputs["post_code"].text().strip(), self.inputs["title"].text().strip(),
                self.inputs["code_name"].text().strip(),
                self.count.value(), self.sets.value(), self.date.date().toString("yyyy-MM-dd"),
                self.exam_type.currentText(), grade=self.inputs["post_grade"].text().strip() or None,
            )
            style_status(self.status, "Exam created. Generate the question mapping to continue.", "success")
            self.update_saved()
            self.refresh(1)
        except (ValueError, OSError) as error:
            self.error(error)

    def generate_mapping(self):
        if self.exam["mapping"] and not self.confirm(
            "Regenerate Mapping", "Create new question orders and print codes? Previous files will stay in the exam folder.",
        ):
            return
        self.work("Generating the question mapping…", lambda: generate_mapping(
            Path(self.settings["output_root"]), self.folder, self.exam, self.settings,
        ), self.mapping_generated)

    def mapping_generated(self, exam):
        self.exam = exam
        self.refresh()
        self.saved_export(mapping_files(self.folder, exam), "Mapping Saved",
                          "Excel and PDF mapping files are ready. Open them here or continue to Set 0 Source.")

    def save_mapping(self):
        self.work("Saving the mapping files…", lambda: save_mapping_exports(
            self.folder, self.exam, self.settings,
        ), self.mapping_resaved)

    def mapping_resaved(self, paths):
        self.refresh()
        self.saved_export(paths, "Mapping Saved", "Excel and PDF mapping files are ready.")

    def open_mapping(self, format_index):
        path = mapping_files(self.folder, self.exam)[format_index]
        if not path.exists():
            self.error("Save the mapping files again.")
        elif format_index == 0:
            local_path(path)
        else:
            Preview(self, path).exec()

    def view_mapping(self):
        window = QDialog(self)
        window.setWindowTitle("Question Mapping")
        window.resize(730, 540)
        layout = QVBoxLayout(window)
        layout.addWidget(subtitle("Set 0 shows the original question numbers. The other columns show their new positions."))
        try:
            rows = mapping_rows(self.exam)
        except ValueError as error:
            self.error(error)
            return
        table = QTableWidget(len(rows), len(rows[0]))
        table.setObjectName("mappingTable")
        table.verticalHeader().hide()
        table.setHorizontalHeaderLabels([f"Set {n}" for n in range(len(rows[0]))])
        table.horizontalHeader().setMinimumSectionSize(65)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        table.setAlternatingRowColors(True)
        for r, line in enumerate(rows):
            table.setRowHeight(r, 29)
            for c, number in enumerate(line):
                item = QTableWidgetItem(str(number))
                item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                table.setItem(r, c, item)
        layout.addWidget(table)
        window.exec()

    def update_elapsed(self):
        minutes, seconds = divmod(int(monotonic() - self.progress_started), 60)
        self.elapsed.setText(f"Elapsed {minutes:02d}:{seconds:02d}")

    def show_progress(self, current, total, message):
        self.progress_bar.setRange(0, total)
        self.progress_bar.setValue(current)
        self.progress_bar.setFormat("%v / %m")
        style_status(self.status, message)

    def work(self, text, function, callback, *, report=False):
        if self.task:
            return
        style_status(self.status, text)
        for control in (self.tabs, self.sidebar_panel):
            control.setEnabled(False)
        if report:
            self.progress_started = monotonic()
            self.progress_bar.setRange(0, 0)
            self.progress_row.show()
            self.update_elapsed()
            self.progress_timer.start()
        task = Task(function, self, report)
        self.task = task

        def unlock():
            self.progress_timer.stop()
            self.progress_row.hide()
            for control in (self.tabs, self.sidebar_panel):
                control.setEnabled(True)

        def successful(result):
            unlock()
            callback(result)

        def failed(message):
            unlock()
            self.refresh()
            self.error(message)

        def end():
            self.task = None
            task.deleteLater()

        task.succeeded.connect(successful, Qt.ConnectionType.QueuedConnection)
        task.failed.connect(failed, Qt.ConnectionType.QueuedConnection)
        task.progress.connect(self.show_progress, Qt.ConnectionType.QueuedConnection)
        task.finished.connect(end, Qt.ConnectionType.QueuedConnection)
        task.start()

    def saved_export(self, files, title, message):
        style_status(self.status, f"{message} · {files[0].parent}", "success")
        QMessageBox.information(self, title, message)

    def confirm(self, title, text):
        return QMessageBox.question(self, title, text, QMessageBox.StandardButton.Yes |
                                    QMessageBox.StandardButton.No, QMessageBox.StandardButton.No) == QMessageBox.StandardButton.Yes

    def generate_zero(self):
        approved, _ = review_status(self.folder, self.exam, self.settings)
        if approved and not self.confirm(
            "Generate Set 0 Again", "Generating a new Set 0 PDF requires a new confirmation.",
        ):
            return
        self.work("Generating Set 0 for review…",
                  lambda progress: export_booklets(self.folder, self.exam, self.settings, self.app_folder,
                                                   [0], ["docx", "pdf"], progress=progress),
                  self.zero_generated, report=True)

    def zero_generated(self, files):
        self.refresh()
        self.saved_export(files, "Set 0 PDF Ready", "The Set 0 PDF is ready. Open it and check the layout before confirming.")

    def preview_zero(self):
        _, path = review_status(self.folder, self.exam, self.settings)
        if path:
            Preview(self, path).exec()
        else:
            self.error("The Set 0 PDF has changed. Generate it again before confirming.")
            self.refresh()

    def approve_zero(self):
        if not self.confirm("Confirm Set 0", "Have you reviewed and approved the Set 0 PDF?"):
            return
        try:
            confirm_zero(self.folder, self.exam, self.settings)
            style_status(self.status, "Set 0 approved. Select the sets and formats to generate.", "success")
            self.refresh(4)
        except (ValueError, OSError) as error:
            self.error(error)
            self.refresh()

    def export(self):
        choice = self.select_set.currentData()
        numbers = list(range(1, self.exam["sets"] + 1)) if choice is None else [choice]
        formats = (["docx", "pdf"], ["docx"], ["pdf"])[self.format.currentIndex()]
        self.work("Generating the selected question sets…",
                  lambda progress: export_booklets(self.folder, self.exam, self.settings, self.app_folder,
                                                   numbers, formats, progress=progress),
                  lambda files: self.sets_generated(files, len(numbers)), report=True)

    def sets_generated(self, files, count):
        self.refresh()
        self.saved_export(files, "Generation Complete",
                          f"Generated {count} {'set' if count == 1 else 'sets'} ({len(files)} files).")

    def open_output(self):
        local_path(self.folder / "GENERATED_SETS")

    def error(self, error):
        style_status(self.status, "Review the message and try again.", "error")
        QMessageBox.warning(self, "Could Not Complete the Step", str(error))

    def closeEvent(self, event):
        if self.task:
            event.ignore()
        else:
            super().closeEvent(event)


def main():
    if sys.platform == "win32":
        from ctypes import windll
        windll.shell32.SetCurrentProcessExplicitAppUserModelID("BPSC.QuestionStudio")
    app = QApplication(sys.argv)
    app.setApplicationName("BPSC Question Studio")
    app.setApplicationDisplayName("BPSC Question Studio")
    app.setStyle("Fusion")
    app.setStyleSheet(THEME)
    app.setWindowIcon(icon())
    app_folder = Path(sys.executable if getattr(sys, "frozen", False) else __file__).resolve().parent
    if not getattr(sys, "frozen", False):
        app_folder = app_folder.parent
    lock = QLockFile(str(app_folder / ".bpsc-question-studio.lock"))
    if not lock.tryLock(0):
        QMessageBox.warning(None, "BPSC Question Studio",
                            "Open one copy of the app at a time from a folder you can write to.")
        return
    try:
        credentials = initialize_access(app_folder)
        if LoginDialog(credentials).exec() != QDialog.DialogCode.Accepted:
            return
        window = Window(app_folder, credentials)
    except (OSError, ValueError) as error:
        QMessageBox.warning(None, "Could Not Open BPSC Question Studio", str(error))
        return
    window.show()
    sys.exit(app.exec())
