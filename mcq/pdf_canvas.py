"""A point-based PDF canvas shared by the printing tools."""

from PySide6.QtCore import QBuffer, QIODevice, QMarginsF, QSizeF
from PySide6.QtGui import QPageLayout, QPageSize, QPainter, QPdfWriter


def margin_painter(width, height, resolution=72):
    buffer = QBuffer()
    if not buffer.open(QIODevice.OpenModeFlag.WriteOnly):
        raise ValueError("Could not open a PDF overlay buffer.")
    writer = QPdfWriter(buffer)
    writer.setResolution(resolution)
    writer.setPageSize(QPageSize(QSizeF(width, height), QPageSize.Unit.Point))
    writer.setPageMargins(QMarginsF(0, 0, 0, 0), QPageLayout.Unit.Point)
    painter = QPainter(writer)
    if not painter.isActive():
        raise ValueError("Could not draw the PDF margins.")
    return buffer, writer, painter
