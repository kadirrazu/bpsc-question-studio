"""Value controls that let wheel events scroll their enclosing form."""

from PySide6.QtWidgets import QComboBox, QDateEdit, QDoubleSpinBox, QSpinBox


class NoWheel:
    def wheelEvent(self, event):
        event.ignore()


class Choice(NoWheel, QComboBox):
    pass


class DateInput(NoWheel, QDateEdit):
    pass


class CountInput(NoWheel, QSpinBox):
    def __init__(self):
        super().__init__()
        self.lineEdit().setTextMargins(10, 0, 0, 0)


class DecimalInput(NoWheel, QDoubleSpinBox):
    def __init__(self):
        super().__init__()
        self.lineEdit().setTextMargins(10, 0, 0, 0)
