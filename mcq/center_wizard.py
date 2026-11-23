"""Interactive preview and export for derived center workbooks."""

from pathlib import Path

from PySide6.QtCore import QSize, Qt
from PySide6.QtWidgets import (
    QAbstractItemView, QCheckBox, QDialog, QFileDialog, QFormLayout, QFrame, QHBoxLayout,
    QHeaderView, QLabel, QLineEdit, QListWidget, QListWidgetItem, QMessageBox,
    QPushButton, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget,
)

from .appearance import icon
from .center_files import OUTPUT_FIELDS, ROOM_CODE, derive_center_rows, read_center_file, save_center_file
from .controls import Choice

WIZARD_STYLE = """
QDialog#centerWizard { background: #F4F1F9; }
QFrame#centerHero { background: #312346; border-radius: 13px; }
QLabel#centerTitle { color: #FFFFFF; font-size: 28px; font-weight: 750; }
QLabel#centerIntro { color: #ECE4F7; }
QFrame#centerPanel { background: white; border: 1px solid #DDD4EA; border-radius: 11px; }
QLabel#centerHeading { color: #372553; font-size: 16px; font-weight: 750; }
QLabel#centerHint { color: #544766; font-size: 13px; }
QLabel#centerStatus { background: #ECE6F5; color: #493761; font-weight: 650;
    padding: 7px 10px; border-radius: 6px; }
QLabel#centerStatus[state="error"] { background: #FCE8EC; color: #852036; }
QLabel#centerStatus[state="ready"] { background: #E8F5ED; color: #1C583C; }
QListWidget#centerColumns { border: 0; padding: 1px; outline: 0; }
QListWidget#centerColumns::item { padding: 4px 6px; border-radius: 5px; }
QListWidget#centerColumns::item:selected { background: #EDE4F7; color: #372553; }
QTableWidget#centerPreview { background: #FFFFFF; alternate-background-color: #F6F3FA;
    border: 1px solid #E5DDEE; border-radius: 7px; padding: 0; }
QTableWidget#centerPreview QHeaderView::section { background: #ECE5F5; color: #372553; padding: 8px; }
QPushButton#centerPrimary { background: #62458B; color: white; border-color: #62458B; }
QPushButton#centerPrimary:hover { background: #503471; }
QPushButton#centerPrimary:disabled { background: #ECE6F3; color: #625770; border-color: #DBD1E6; }
"""

COLUMN_LABELS = {"c_code": "Center Code", "c_name": "Center Name", "building": "Building",
                 "floor": "Floor", "room_no": "Room Number", "no_of_cand": "Candidates",
                 "room_serial": "Room Serial", "question_count": "Question Count", ROOM_CODE: "Room Code"}


def wizard_label(text, name):
    label = QLabel(text)
    label.setObjectName(name)
    label.setWordWrap(True)
    return label


def wizard_panel(parent, name="centerPanel"):
    panel = QFrame()
    panel.setObjectName(name)
    layout = QVBoxLayout(panel)
    layout.setContentsMargins(16, 12, 16, 12)
    layout.setSpacing(7)
    parent.addWidget(panel)
    return layout


class CenterFileWizard(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Center File Wizard")
        self.setObjectName("centerWizard")
        self.setStyleSheet(WIZARD_STYLE)
        self.resize(1100, 740)
        self.source = self.output = None
        self.input_data = b""
        self.rows = []
        self.derived = {}
        self.include_code = QCheckBox("Add a unique 8-character room code")
        self.auto_do = QCheckBox("Fill instruction with DO")
        self.count = Choice()
        self.count.addItem("Same as number of candidates", 1)
        self.count.addItem("Round up to a multiple of 4", 4)
        self.separator = QLineEdit("-")
        self.separator.setMaxLength(8)
        self.components = QListWidget()
        self.components.setObjectName("centerColumns")
        self.components.setAccessibleName("File-name components in order")
        self.components.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.components.setToolTip("Check components to include. Drag to change the file-name order.")
        self.components.setMaximumHeight(150)
        for field in ("room_serial", "c_code", "room_no", "no_of_cand", "c_name", "building", "floor",
                      "question_count", ROOM_CODE):
            item = QListWidgetItem(COLUMN_LABELS[field])
            item.setData(Qt.ItemDataRole.UserRole, field)
            item.setToolTip(field)
            item.setCheckState(Qt.CheckState.Checked if field in {"room_serial", "c_code", "room_no"}
                               else Qt.CheckState.Unchecked)
            self.components.addItem(item)
        self.preview = QTableWidget()
        self.preview.setObjectName("centerPreview")
        self.preview.setAccessibleName("Derived Center File preview")
        self.preview.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.preview.setAlternatingRowColors(True)
        self.preview.setShowGrid(False)
        self.preview.verticalHeader().hide()
        self.preview.verticalHeader().setDefaultSectionSize(31)
        self.preview.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.message = wizard_label("Open a center workbook to preview the derived columns.", "centerStatus")
        self.save_button = QPushButton("Generate Center File")
        self.save_button.setObjectName("centerPrimary")
        self.save_button.setIcon(icon("export", light=True))
        self.save_button.setEnabled(False)
        self.save_button.clicked.connect(self.save_file)
        browse = QPushButton("Open Source Workbook")
        browse.setIcon(icon("folder"))
        browse.clicked.connect(self.open_file)
        self.filename = wizard_label("No workbook selected", "centerHint")
        self.filename.setTextFormat(Qt.TextFormat.PlainText)
        inputs = QHBoxLayout()
        inputs.addWidget(browse)
        inputs.addWidget(self.filename, 1)
        form = QFormLayout()
        form.addRow("Question Count", self.count)
        form.addRow("File-name Separator", self.separator)
        form.addRow(self.auto_do)
        form.addRow(self.include_code)
        left = QWidget()
        left.setObjectName("fieldGroup")
        left.setLayout(form)
        config = QHBoxLayout()
        config.setSpacing(14)
        settings = wizard_panel(config)
        settings.addWidget(wizard_label("Generation Rules", "centerHeading"))
        settings.addWidget(left)
        file_components = wizard_panel(config)
        file_components.addWidget(wizard_label("File-name Components", "centerHeading"))
        file_components.addWidget(wizard_label("Check the fields to include and drag them into order.", "centerHint"))
        file_components.addWidget(self.components)
        close = QPushButton("Close")
        close.clicked.connect(self.reject)
        buttons = QHBoxLayout()
        buttons.addStretch()
        buttons.addWidget(self.save_button)
        buttons.addWidget(close)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 16, 18, 16)
        layout.setSpacing(11)
        hero = wizard_panel(layout, "centerHero")
        heading = QHBoxLayout()
        mark = QLabel()
        mark.setPixmap(icon("mapping", light=True).pixmap(QSize(25, 25)))
        heading.addWidget(mark)
        heading.addWidget(wizard_label("Center File Wizard", "centerTitle"), 1)
        hero.addLayout(heading)
        hero.addWidget(wizard_label("Prepare a center workbook for Room Print Packs.", "centerIntro"))
        source_panel = wizard_panel(layout)
        source_panel.addWidget(wizard_label("Source Workbook", "centerHeading"))
        source_panel.addLayout(inputs)
        layout.addLayout(config)
        preview_panel = wizard_panel(layout)
        preview_panel.addWidget(wizard_label("Live Preview", "centerHeading"))
        preview_panel.addWidget(self.message)
        preview_panel.addWidget(self.preview, 1)
        layout.setStretchFactor(preview_panel.parentWidget(), 1)
        layout.addLayout(buttons)
        self.count.currentIndexChanged.connect(self.refresh_preview)
        self.separator.textChanged.connect(self.refresh_preview)
        self.auto_do.toggled.connect(self.refresh_preview)
        self.include_code.toggled.connect(self.refresh_preview)
        self.components.itemChanged.connect(self.refresh_preview)
        self.components.model().rowsMoved.connect(self.refresh_preview)
        self.components.model().rowsRemoved.connect(self.refresh_preview)

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Center Workbook", "", "Excel Workbook (*.xlsx)")
        if not path:
            return
        try:
            source = Path(path)
            data = source.read_bytes()
            name, rows = read_center_file(data)
        except Exception as error:
            QMessageBox.warning(self, "Could Not Read Workbook", str(error))
            return
        self.source, self.input_data, self.sheet_name, self.rows = source, data, name, rows
        self.filename.setText(str(source))
        self.refresh_preview()

    def refresh_preview(self):
        if not self.rows:
            return
        self.preview.setRowCount(0)
        self.save_button.setEnabled(False)
        fields = [self.components.item(i).data(Qt.ItemDataRole.UserRole) for i in range(self.components.count())
                  if self.components.item(i).checkState() == Qt.CheckState.Checked]
        include = self.include_code.isChecked()
        try:
            self.derived = derive_center_rows(self.rows, self.count.currentData(), fields,
                                              self.separator.text(), self.auto_do.isChecked(), include)
        except ValueError as error:
            self.preview_status(str(error), "error")
            return
        headers = ("Excel Row", *(COLUMN_LABELS.get(field, field.replace("_", " ").title()) for field in
                    ("c_code", "room_no", "no_of_cand", *OUTPUT_FIELDS, *((ROOM_CODE,) if include else ()))))
        sample = self.rows[:12]
        self.preview.setColumnCount(len(headers))
        self.preview.setHorizontalHeaderLabels(headers)
        self.preview.setRowCount(len(sample))
        for index, row in enumerate(sample):
            values = (row.excel_row, row.values["c_code"], row.values["room_no"], row.candidates,
                      *self.derived[row.excel_row])
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value))
                if column in (0, 3, 4, 5, 7, 8, 9):
                    item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
                self.preview.setItem(index, column, item)
        self.preview_status(f"Showing {len(sample)} of {len(self.rows)} rooms. The source workbook stays unchanged.", "ready")
        self.save_button.setEnabled(True)

    def preview_status(self, message, state):
        self.message.setText(message)
        self.message.setProperty("state", state)
        self.message.style().unpolish(self.message)
        self.message.style().polish(self.message)

    def save_file(self):
        suggested = self.source.with_name(self.source.stem + "_DERIVED.xlsx")
        name, _ = QFileDialog.getSaveFileName(self, "Save Derived Center File", str(suggested),
                                              "Excel Workbook (*.xlsx)")
        if not name:
            return
        destination = Path(name)
        if destination.suffix.lower() != ".xlsx":
            destination = destination.with_name(destination.name + ".xlsx")
            if destination.exists() and QMessageBox.question(
                self, "Replace Workbook?", f"Replace {destination.name}?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            ) != QMessageBox.StandardButton.Yes:
                return
        try:
            save_center_file(self.source, self.input_data, destination, self.sheet_name,
                             self.derived, self.include_code.isChecked())
        except (OSError, ValueError) as error:
            QMessageBox.warning(self, "Could Not Save Workbook", str(error))
            return
        self.output = destination
        self.accept()
