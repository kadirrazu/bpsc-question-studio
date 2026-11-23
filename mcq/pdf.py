"""Render the same DOCX through Writer, then apply PDF-only margin text."""

import os
import re
import shutil
import subprocess
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from tempfile import TemporaryDirectory

import fitz
from PySide6.QtGui import QColor, QFont, QFontDatabase, QFontMetricsF, QGuiApplication

from .documents import text_tag_values
from .pdf_canvas import margin_painter
from .project import footer_parts, label, text_appearance
from .file_io import write_bytes

MARGIN_DPI = 144  # A half-point font size occupies an integer number of pixels.
MARGIN_SCALE = MARGIN_DPI / 72


def renderer_path(configured, app_folder):
    candidates = [
        Path(configured) if configured else None,
        app_folder / "LibreOffice" / "App" / "libreoffice" / "program" / "soffice.com",
        app_folder / "LibreOffice" / "program" / "soffice.com",
        app_folder / "LibreOffice" / "program" / "soffice.exe",
        app_folder / "LibreOffice" / "program" / "soffice",
        Path(os.environ.get("ProgramFiles", "C:/Program Files")) / "LibreOffice/program/soffice.com",
        Path(os.environ.get("ProgramFiles(x86)", "C:/Program Files (x86)")) / "LibreOffice/program/soffice.com",
    ]
    for candidate in candidates:
        if candidate and candidate.is_file():
            if candidate.name.lower() in ("soffice.exe", "libreoffice.exe"):
                console = candidate.with_name("soffice.com")
                if console.is_file():
                    return console
            if candidate.name.lower() == "libreofficeportable.exe":
                for name in ("soffice.com", "soffice.exe"):
                    direct = candidate.parent / "App" / "libreoffice" / "program" / name
                    if direct.is_file():
                        return direct
                raise ValueError("Select LibreOffice's program/soffice.com in Common Settings → Renderer.")
            return candidate
    return shutil.which("soffice.com") or shutil.which("libreoffice") or shutil.which("soffice")


def render_words(sources, out, settings, app_folder):
    renderer = renderer_path(settings["renderer"], app_folder)
    if not renderer:
        raise ValueError("Install LibreOffice or select its executable under Common Settings → Renderer.")
    out.mkdir(parents=True, exist_ok=True)
    if len(sources) >= 4 and (os.cpu_count() or 1) >= 4:
        midpoint = (len(sources) + 1) // 2
        with ThreadPoolExecutor(max_workers=2) as pool:
            first = pool.submit(render_batch, renderer, sources[:midpoint], out)
            second = pool.submit(render_batch, renderer, sources[midpoint:], out)
            return first.result() + second.result()
    return render_batch(renderer, sources, out)


def render_batch(renderer, sources, out):
    # A private profile prevents a leftover Writer process or recovery state
    # from taking over a subsequent conversion.
    with TemporaryDirectory(prefix="writer-", dir=out.parent, ignore_cleanup_errors=True) as work:
        profile = Path(work) / "profile"
        converted = Path(work) / "pdf"
        converted.mkdir()
        config = profile / "user" / "registrymodifications.xcu"
        write_bytes(config, (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<oor:items xmlns:oor="http://openoffice.org/2001/registry">'
            '<item oor:path="/org.openoffice.Office.Common/Load/Document">'
            '<prop oor:name="LoadPrinter" oor:op="fuse"><value>false</value></prop>'
            '</item></oor:items>'
        ).encode("utf-8"))
        command = [
            str(renderer), f"-env:UserInstallation={profile.resolve().as_uri()}",
            "--headless", "--invisible", "--nologo", "--nodefault",
            "--nofirststartwizard", "--norestore", "--convert-to", "pdf:writer_pdf_Export",
            "--outdir", str(converted.resolve()), *(str(source.resolve()) for source in sources),
        ]
        environment = dict(os.environ, SAL_DISABLE_PRINTERLIST="1",
                           SAL_DISABLE_SYNCHRONOUS_PRINTER_DETECTION="1", SAL_DISABLESKIA="1")
        # Writer may start a child process on Windows. File-backed output avoids
        # waiting for an inherited pipe to close after a timeout.
        log_path = Path(work) / "writer.log"
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        with log_path.open("wb") as log:
            try:
                process = subprocess.Popen(command, stdin=subprocess.DEVNULL, stdout=log,
                                           stderr=subprocess.STDOUT, env=environment, creationflags=flags)
            except OSError as error:
                raise ValueError(f"Could not start LibreOffice at {renderer}. Check Common Settings → Renderer. {error}") from error
            try:
                returncode = process.wait(timeout=60 + 20 * len(sources))
            except subprocess.TimeoutExpired as error:
                if os.name == "nt":
                    try:
                        subprocess.run(["taskkill", "/PID", str(process.pid), "/T", "/F"],
                                       stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL,
                                       stderr=subprocess.DEVNULL, creationflags=flags, timeout=5, check=False)
                    except (OSError, subprocess.TimeoutExpired):
                        pass
                if process.poll() is None:
                    process.kill()
                process.wait(timeout=5)
                detail = renderer_output(log_path)
                hint = ("Check that Common Settings → Renderer points to LibreOffice's program/soffice.com. "
                        "Open the source DOCX in LibreOffice Writer to check for a prompt.")
                raise ValueError(f"LibreOffice timed out while converting the question paper. {hint}\n{renderer}\n{detail}".strip()) from error
        expected = [converted / (source.stem + ".pdf") for source in sources]
        if returncode or any(not file.is_file() or not file.stat().st_size for file in expected):
            detail = renderer_output(log_path)
            raise ValueError(f"LibreOffice could not generate every PDF.\n{renderer}\n{detail}".strip())
        return [file.replace(out / file.name) for file in expected]


def renderer_output(path):
    with path.open("rb") as log:
        log.seek(max(0, path.stat().st_size - 2048))
        return log.read().decode(errors="replace").strip()


def paint_runners(painter, width, height, settings, text):
    metrics = painter.fontMetrics()
    offset = settings["runner_offset"] * MARGIN_DPI
    interval = settings["runner_interval"] * MARGIN_DPI
    text_width = metrics.horizontalAdvance(text)
    for side in settings["sides"]:
        if offset + metrics.height() > settings[side] * MARGIN_DPI:
            raise ValueError(f"The runner does not fit in the {side} margin.")
        painter.save()
        if side == "right":
            painter.translate(width, 0)
            painter.rotate(90)
        elif side == "left":
            painter.translate(0, height)
            painter.rotate(-90)
        vertical = side in ("left", "right")
        start = settings["top" if vertical else "left"] * MARGIN_DPI
        end = (height if vertical else width) - settings["bottom" if vertical else "right"] * MARGIN_DPI
        if side == "left":
            start, end = height - end, height - start
        if side == "bottom":
            footer_font, _ = margin_text_style(settings, "footer")
            footer_metrics = QFontMetricsF(footer_font, painter.device())
            prefix, separator = footer_parts(25, settings)
            page = label(999) if settings["footer_language"] == "Bangla" else "999"
            if settings["print_code_enabled"] and settings["print_code_opacity"] > 0:
                code_font, _ = margin_text_style(settings, "print_code")
                start += QFontMetricsF(code_font, painter.device()).horizontalAdvance("WWWWWWWW") + 10 * MARGIN_SCALE
            if settings["footer_opacity"] > 0:
                end -= footer_metrics.horizontalAdvance(f"{prefix}{page}{separator}{page}") + 10 * MARGIN_SCALE
        if text_width > end - start:
            painter.restore()
            raise ValueError(f"The runner text is too long for the {side} side.")
        baseline = height - offset - metrics.descent() if side == "bottom" else offset + metrics.ascent()
        x = start
        while x + text_width <= end:
            painter.drawText(round(x), round(baseline), text)
            x += text_width + interval
        painter.restore()


def margin_text_style(settings, kind):
    family, size, color, opacity = text_appearance(settings, kind)
    font = QFont(family)
    font.setPointSizeF(size)
    if kind in ("runner", "print_code") and family == "Courier New":
        font.setStyleHint(QFont.StyleHint.Monospace)
    pen = QColor(color)
    pen.setAlphaF(opacity)
    return font, pen


def footer_baseline(painter, height, settings):
    return round(height - min(0.2, settings["bottom"] / 3) * MARGIN_DPI - painter.fontMetrics().descent())


def runner_overlay(width, height, settings, text, identifier):
    buffer, writer, painter = margin_painter(width, height, MARGIN_DPI)
    width, height = width * MARGIN_SCALE, height * MARGIN_SCALE
    try:
        if text:
            font, color = margin_text_style(settings, "runner")
            painter.setFont(font)
            painter.setPen(color)
            paint_runners(painter, width, height, settings, text)
        if identifier:
            font, color = margin_text_style(settings, "print_code")
            painter.setFont(font)
            painter.setPen(color)
            painter.drawText(round(settings["left"] * MARGIN_DPI), footer_baseline(painter, height, settings), identifier)
    finally:
        painter.end()
    del painter, writer
    buffer.close()
    return bytes(buffer.data())


def uses_footer_overlay(settings):
    return (settings["footer_language"] == "Bangla" or settings["footer_opacity"] < 1
            or (settings["print_code_enabled"] and settings["print_code_opacity"] > 0))


def paging_pdf(width, height, settings, number, pages):
    buffer, writer, painter = margin_painter(width, height, MARGIN_DPI)
    width, height = width * MARGIN_SCALE, height * MARGIN_SCALE
    page_labels = []
    digits = label if settings["footer_language"] == "Bangla" else str
    prefix, separator = footer_parts(number, settings)
    try:
        font, color = margin_text_style(settings, "footer")
        painter.setFont(font)
        painter.setPen(color)
        for index in range(1, pages + 1):
            if index > 1:
                writer.newPage()
            page_label = f"{prefix}{digits(index)}{separator}{digits(pages)}"
            page_labels.append(page_label)
            x = width - settings["right"] * MARGIN_DPI - painter.fontMetrics().horizontalAdvance(page_label)
            painter.drawText(round(x), footer_baseline(painter, height, settings), page_label)
    finally:
        painter.end()
    del painter, writer
    buffer.close()
    pdf = fitz.open(stream=bytes(buffer.data()), filetype="pdf")
    if settings["footer_language"] == "Bangla":
        # Keep Bengali text searchable when the font maps ligatures to glyph IDs.
        for page, text in zip(pdf, page_labels):
            encoded = text.encode("utf-16-be").hex().upper().encode("ascii")
            for xref in page.get_contents():
                stream = pdf.xref_stream(xref)
                pdf.update_stream(xref, b"/Span << /ActualText <FEFF" + encoded + b"> >> BDC\n"
                                  + stream + b"\nEMC\n")
    return pdf


def finish_booklet_pdf(source, target, exam, number, settings):
    text = settings["runner_text"].strip() if settings["sides"] and settings["runner_opacity"] else ""
    footer_overlay = uses_footer_overlay(settings) and settings["footer_opacity"] > 0
    code = (exam["records"][number]["code"]
            if settings["print_code_enabled"] and settings["print_code_opacity"] > 0 else None)
    if not text and not footer_overlay and not code:
        shutil.move(source, target)
        return
    values = text_tag_values(number, exam["code_name"], settings["runner_tag_language"])
    text = re.sub("|".join(re.escape(tag) for tag in values), lambda match: values[match.group()], text)
    if settings["runner_uppercase"]:
        text = text.upper()
    with fitz.open(source) as document:
        first = document[0].rect
        if text or code:
            overlay_bytes = runner_overlay(first.width, first.height, settings, text, code)
            with fitz.open(stream=overlay_bytes, filetype="pdf") as overlay:
                for page in document:
                    page.show_pdf_page(page.rect, overlay, 0, overlay=True)
        if footer_overlay:
            with paging_pdf(first.width, first.height, settings, number, len(document)) as paging:
                for index, page in enumerate(document):
                    page.show_pdf_page(page.rect, paging, index, overlay=True)
        document.save(target, garbage=2)


def check_pdf_fonts(fonts):
    if QGuiApplication.instance() is None:
        return
    available = {name.casefold() for name in QFontDatabase.families()}
    missing = sorted(font for font in fonts if font.casefold() not in available)
    if missing:
        raise ValueError("Install the fonts used by the document before exporting PDF: " + ", ".join(missing))
