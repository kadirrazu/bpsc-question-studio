"""Shared application icon and desktop styles."""

from pathlib import Path

from PySide6.QtGui import QIcon

ASSETS = Path(__file__).resolve().parents[1] / "assets"


def icon(name="studio", light=False):
    file = "studio.png" if name == "studio" else name + ("-light" if light else "") + ".svg"
    return QIcon(str(ASSETS / file))


THEME = """
QWidget { color: #182C48; background: #F3F6FC; font-size: 14px; }
QLabel { background: transparent; }
QWidget#sidebar { background: #142840; }
QWidget#hero { border-radius: 15px;
 background: qlineargradient(x1:0,y1:0,x2:1,y2:1,stop:0 #183758,stop:1 #234D7D); }
QLabel#title { color: white; font-size: 31px; font-weight: 750; }
QLabel#heroContext { color: #D2E1F3; font-size: 13px; }
QLabel#moduleTitle { color: #C9D9EE; font-size: 13px; font-weight: 700; }
QLabel#sideTitle { color: #FFFFFF; font-size: 15px; font-weight: 700; }
QLabel#stepTitle { font-size: 27px; font-weight: 750; }
QLabel#subtitle { color: #4C6280; font-size: 14px; }
QLabel#status { font-size: 13px; color: #27496C; padding: 12px 14px;
 background: #E6EFFA; border: 1px solid #D8E6F7; border-radius: 9px; }
QLabel#status[state="success"] { color: #12583B; background: #E8F7EE; border-color: #B9DFC8; font-weight: 700; }
QLabel#status[state="warning"] { color: #704914; background: #FFF7E5; border-color: #E8D6A6; font-weight: 650; }
QLabel#status[state="error"] { color: #8F252D; background: #FDEFF0; border-color: #EDC0C5; font-weight: 700; }
QFrame#separator { background: #DCE5F0; }
QProgressBar { background: #E4EDF8; border: 1px solid #C5D8EF; border-radius: 7px;
 color: #122C47; font-weight: 700; text-align: center; min-height: 19px; }
QProgressBar::chunk { background: #79ADE4; border-radius: 6px; }
QTableWidget#mappingTable { background-color: #F8F9FA; alternate-background-color: #EEF0F3; color: #182C48; }
QTableWidget#mappingTable QHeaderView::section { background: #E2E6EC; color: #182C48; }
QFrame#card { background: white; border: 1px solid #DEE7F3; border-radius: 13px; }
QWidget#fieldGroup { background: transparent; }
QTabWidget::pane { border: 0; background: #F3F6FC; }
QTabBar::tab { background: #E6EEF9; color: #2B4D72; border: 1px solid #D5E2F2;
 border-radius: 8px; margin: 0 5px 2px 0; padding: 10px 10px; font-weight: 650; }
QTabBar::tab:selected { background: white; border: 1px solid #548BD0; color: #174475; }
QTabBar::tab:disabled { background: #EDF1F7; color: #78889C; border-color: #E0E7F1; }
QGroupBox { background: white; border: 1px solid #DEE7F3; border-radius: 11px;
 margin-top: 13px; padding: 9px 12px 7px; font-weight: 650; }
QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 5px; }
QLineEdit, QDateEdit, QSpinBox, QDoubleSpinBox, QComboBox, QListWidget, QTableWidget {
 background: white; border: 1px solid #CBD9EB; border-radius: 7px; padding: 7px;
 selection-background-color: #265E9B; selection-color: white; }
QLineEdit:focus, QDateEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus, QComboBox:focus {
 border: 1px solid #2D70BA; }
QLineEdit:disabled, QDateEdit:disabled, QSpinBox:disabled, QComboBox:disabled {
 color: #4C6280; background: #F0F4F9; }
QComboBox, QDateEdit { padding-right: 32px; }
QComboBox::drop-down, QDateEdit::drop-down { subcontrol-origin: border; width: 28px; border: none; }
QComboBox::down-arrow { image: url("@ASSETS@/chevron.svg"); width: 14px; height: 14px; }
QDateEdit::down-arrow { image: url("@ASSETS@/calendar.svg"); width: 16px; height: 16px; }
QSpinBox, QDoubleSpinBox { padding: 0 31px 0 2px; min-height: 36px; }
QSpinBox QLineEdit, QDoubleSpinBox QLineEdit { padding: 0; border: none; background: transparent; }
QSpinBox::up-button, QDoubleSpinBox::up-button, QSpinBox::down-button, QDoubleSpinBox::down-button {
 subcontrol-origin: border; width: 26px; height: 17px; right: 1px;
 background: #EAF1FC; border-left: 1px solid #CBD9EB; }
QSpinBox::up-button, QDoubleSpinBox::up-button { subcontrol-position: top right; top: 1px;
 border-top-right-radius: 6px; border-bottom: 1px solid #CBD9EB; }
QSpinBox::down-button, QDoubleSpinBox::down-button { subcontrol-position: bottom right; bottom: 1px;
 border-bottom-right-radius: 6px; }
QSpinBox::up-button:hover, QDoubleSpinBox::up-button:hover, QSpinBox::down-button:hover, QDoubleSpinBox::down-button:hover {
 background: #D6E5FA; }
QSpinBox::up-arrow, QDoubleSpinBox::up-arrow { image: url("@ASSETS@/plus.svg"); width: 12px; height: 12px; }
QSpinBox::down-arrow, QDoubleSpinBox::down-arrow { image: url("@ASSETS@/minus.svg"); width: 12px; height: 12px; }
QCheckBox { spacing: 9px; min-height: 24px; }
QCheckBox::indicator { width: 18px; height: 18px; background: white;
 border: 1px solid #7289A8; border-radius: 5px; }
QCheckBox::indicator:hover, QCheckBox::indicator:focus { border-color: #2467AB; }
QCheckBox::indicator:checked { background: #265E9B; border-color: #265E9B;
 image: url("@ASSETS@/checkbox-check.svg"); }
QCheckBox::indicator:disabled { background: #E5ECF5; border-color: #AEBCCF; }
QCheckBox::indicator:checked:disabled { background: #71849F; border-color: #71849F; }
QPushButton { background: #FFFFFF; color: #214569; border: 1px solid #C8D8EB;
 border-radius: 8px; padding: 9px 12px; font-weight: 650; }
QPushButton:hover { background: #EFF5FE; border-color: #8DB3D9; }
QPushButton:focus, QTabBar::tab:focus { border: 1px solid #387BC5; }
QPushButton#settingsSection { text-align: left; padding: 7px 9px; }
QPushButton#settingsSection:checked { background: #E6EFFA; border-color: #548BD0; color: #174475; }
QPushButton#primary { background: #2563A3; border-color: #2563A3; color: white; }
QPushButton#primary:hover { background: #1C5189; }
QPushButton:disabled, QPushButton#primary:disabled { background: #EBF0F6; color: #74859B; border-color: #DFE7F1; }
QPushButton#navButton { text-align: left; background: #1B3756; border-color: #2C4A6C; color: white; }
QPushButton#navButton:hover { background: #25476A; }
QPushButton#navButton[destructive="true"]:hover { background: #672E3B; border-color: #9B5F6D; }
QPushButton#navButton:disabled { color: #9BADBF; background: #23364C; border-color: #2C4259; }
QListWidget#examList { background: #19324E; color: #EFF5FF; border-color: #2B4768; }
QListWidget#examList::item { padding: 8px; border-radius: 6px; }
QListWidget#examList::item:selected { background: #2E5C90; color: white; }
QHeaderView::section { background: #E5EEF9; border: 0; padding: 6px; font-weight: 650; }
QCalendarWidget QWidget#qt_calendar_navigationbar { background: #DFEBF9; }
""".replace("@ASSETS@", ASSETS.as_posix())
