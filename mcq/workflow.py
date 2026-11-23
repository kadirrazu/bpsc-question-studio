"""Filesystem operations shared by the UI and the regression checks."""

from datetime import datetime
from hashlib import sha256
import json
import ntpath
import os
from pathlib import Path
from tempfile import TemporaryDirectory

from .documents import declared_fonts, make_booklet, open_word, question_rows, validate_template
from .mapping import export_mapping_pdf, export_mapping_xlsx
from .pdf import check_pdf_fonts, finish_booklet_pdf, render_words, uses_footer_overlay
from .project import DEFAULTS, DHAKA, commit_mapping, file_prefix, post_name, prepare_mapping, save_exam, template_path, validate_mapping
from .file_io import write_bytes


def import_source(folder, exam, source):
    question_rows(open_word(source), exam["count"])
    target = folder / "SOURCE" / "ZERO_SET.docx"
    write_bytes(target, Path(source).read_bytes())
    updated = {**exam, "source": Path(source).name}
    save_exam(folder, updated)
    return updated


def save_template(root, kind, source):
    validate_template(open_word(source))
    target = root / "Config" / "Template" / (kind.upper() + ".docx")
    write_bytes(target, Path(source).read_bytes())
    return str(target.resolve())


def mapping_files(folder, exam):
    name = f"{file_prefix(exam)}-MAPPING-{exam['records'][0]['code']}"
    return [folder / "MAPPING" / f"{name}.{kind}" for kind in ("xlsx", "pdf")]


def save_mapping_exports(folder, exam, settings):
    files = mapping_files(folder, exam)
    with TemporaryDirectory(prefix="mcq-mapping-") as temporary:
        scratch = Path(temporary)
        staged = [scratch / file.name for file in files]
        export_mapping_xlsx(exam, staged[0], settings)
        export_mapping_pdf(exam, staged[1], settings)
        originals = {path: path.read_bytes() for path in files if path.is_file()}
        published = []
        try:
            for source, target in zip(staged, files):
                write_bytes(target, source.read_bytes())
                published.append(target)
        except OSError:
            for path in published:
                if path in originals:
                    write_bytes(path, originals[path])
                else:
                    path.unlink(missing_ok=True)
            raise
    return files


def generate_mapping(data, folder, exam, settings):
    updated = prepare_mapping(data, exam)
    files = save_mapping_exports(folder, updated, settings)
    try:
        commit_mapping(folder, updated)
    except OSError:
        for file in files:
            file.unlink(missing_ok=True)
        raise
    return updated


def digest(path):
    return sha256(path.read_bytes()).hexdigest()


def _review_payload(folder, exam, settings):
    layout = {key: value for key, value in settings.items()
              if not key.startswith("mapping_") and key not in {
                  "output_root", "template_root", "header", "footer", "renderer", "settings_version",
                  "use_default_header", "use_default_footer",
              }}
    if not settings["print_code_enabled"]:
        for key in ("print_code_font", "print_code_size", "print_code_color", "print_code_opacity"):
            layout.pop(key, None)
    runner_active = bool(settings["sides"] and settings["runner_text"].strip())
    for key in ("tag_language", "runner_tag_language", "qr_inset", "qr_gap", "set_rotation", "code_rotation",
                "footer_font", "footer_size", "footer_color",
                "footer_opacity", "footer_uppercase", "runner_font", "runner_size",
                "runner_color", "runner_opacity", "runner_uppercase"):
        if layout.get(key) == DEFAULTS[key] or (key.startswith("runner_") and not runner_active):
            layout.pop(key, None)
    if runner_active:
        layout["runner_style_version"] = 1
    if settings["footer_language"] == "Bangla":
        layout["footer_format_version"] = 1
    for kind in ("qr", "set", "code"):
        for axis in ("horizontal", "vertical"):
            key = f"{kind}_{axis}"
            if layout.get(key) == "Template" or (kind != "qr" and not settings[f"{kind}_rotation"]):
                layout.pop(key, None)
    for template in ("header", "footer"):
        path = template_path(settings, template)
        layout[template] = digest(path) if path else ""
    payload = {"layout": layout, "renderer": settings.get("renderer", ""), "mapping": exam["mapping"],
               "codes": [record["code"] for record in exam["records"]],
               "code_name": exam["code_name"], "source": digest(folder / "SOURCE" / "ZERO_SET.docx")}
    if "post" in settings.get("qr_fields", []) and exam.get("post_grade") not in (None, ""):
        payload["qr_post"] = post_name(exam)
    return payload


def _signature(payload):
    return sha256(json.dumps(payload, sort_keys=True, ensure_ascii=False).encode()).hexdigest()


def _current_signature(payload):
    renderer = payload["renderer"]
    if renderer:
        renderer = (ntpath.normcase(ntpath.normpath(renderer))
                    if os.name == "nt" or ntpath.splitdrive(renderer)[0]
                    else os.path.normpath(renderer))
    layout = {key: int(value) if type(value) is float and value.is_integer() else value
              for key, value in payload["layout"].items()}
    return _signature({**payload, "layout": layout, "renderer": renderer, "schema": 2})


def review_signature(folder, exam, settings):
    return _current_signature(_review_payload(folder, exam, settings))


def _review_matches(review, folder, exam, settings):
    if not review.get("signature"):
        return False
    payload = _review_payload(folder, exam, settings)
    if review.get("schema") == 2:
        return _current_signature(payload) == review["signature"]

    # v0.3.1 could hash forward slashes from the file picker and backslashes after a Windows restart.
    if "qr_post" in payload:
        payload["qr_post"] = f"{exam['title']} ({exam['post_grade']})"
    renderer = payload["renderer"]
    spellings = {renderer}
    if os.name == "nt" or ntpath.splitdrive(renderer)[0]:
        spellings.update((renderer.replace("\\", "/"), renderer.replace("/", "\\")))
    return any(_signature({**payload, "renderer": spelling}) == review["signature"]
               for spelling in spellings)


def review_status(folder, exam, settings):
    review = exam.get("review") or {}
    if not review.get("file"):
        return False, None
    output = folder / "GENERATED_SETS"
    pdf = output / review["file"]
    try:
        is_current = (all((output / name).is_file() for name in review.get("files", [pdf.name]))
                      and digest(pdf) == review["pdf_digest"]
                      and _review_matches(review, folder, exam, settings))
    except (OSError, KeyError):
        return False, None
    if not is_current:
        return False, None
    return bool(review.get("confirmed_at")), pdf


def confirm_zero(folder, exam, settings):
    approved, pdf = review_status(folder, exam, settings)
    if not pdf:
        raise ValueError("Generate and review the current Set 0 PDF before confirming.")
    if not approved:
        updated = {**exam, "review": {**exam["review"],
                   "confirmed_at": datetime.now(DHAKA).isoformat(timespec="seconds")}}
        save_exam(folder, updated)
        exam.update(updated)


def generated_numbers(folder, exam):
    outputs = exam.get("generated", {})
    output = folder / "GENERATED_SETS"
    return [number for number in range(1, exam["sets"] + 1)
            if (names := outputs.get(str(number))) and all((output / name).is_file() for name in names)]


def generated_complete(folder, exam, settings):
    approved, _ = review_status(folder, exam, settings)
    return approved and len(generated_numbers(folder, exam)) == exam["sets"]


def export_booklets(folder, exam, settings, app_folder, numbers, formats, *, progress=None):
    current = 0
    render_pdf = "pdf" in formats
    total = 2 + len(numbers) + (1 + len(numbers) if render_pdf else 0)

    def report(message, advance=False):
        nonlocal current
        current += int(advance)
        if progress is not None:
            progress(current, total, message)

    report("Checking the question source, templates and settings…")
    validate_mapping(exam)
    if not numbers or any(number not in range(exam["sets"] + 1) for number in numbers):
        raise ValueError("Select a valid set to export.")
    if 0 in numbers and numbers != [0]:
        raise ValueError("Generate Set 0 by itself and confirm its PDF before generating the other sets.")
    is_zero = numbers == [0]
    signature = review_signature(folder, exam, settings)
    if is_zero and "pdf" not in formats:
        raise ValueError("The Set 0 review requires a PDF.")
    if not is_zero and not review_status(folder, exam, settings)[0]:
        raise ValueError("Review and confirm the latest Set 0 PDF before exporting the other sets.")
    source = folder / "SOURCE" / "ZERO_SET.docx"
    output = folder / "GENERATED_SETS"
    if "pdf" in formats:
        font_names = declared_fonts(source) | {settings["font"]}
        if settings["footer_font"]:
            font_names.add(settings["footer_font"])
        for kind in ("header", "footer"):
            path = template_path(settings, kind)
            if path:
                font_names.update(declared_fonts(path))
        check_pdf_fonts(font_names)
    now = datetime.now(DHAKA).isoformat(timespec="seconds")
    report("Input checks complete.", advance=True)
    records = [
        {**record, "timestamp": record["timestamp"] or now} if number in numbers else record
        for number, record in enumerate(exam["records"])
    ]
    working_exam = {**exam, "records": records}
    staged = []
    pdf_jobs = []
    output_files = {}
    with TemporaryDirectory(prefix="mcq-sets-", ignore_cleanup_errors=True) as temporary:
        scratch = Path(temporary)
        for number in numbers:
            report(f"Preparing the Word document for Set {number}…")
            code = exam["records"][number]["code"]
            name = f"{file_prefix(exam)}-SET-{number:02d}-{code}"
            docx = scratch / (name + ".docx")
            render_source = docx
            if "pdf" in formats and uses_footer_overlay(settings):
                render_source = scratch / ("print-" + name + ".docx")
            make_booklet(source, working_exam, number, settings, docx,
                         pdf_target=render_source if render_source != docx else None,
                         keep_docx="docx" in formats)
            output_files[str(number)] = []
            if "docx" in formats:
                staged.append(docx)
                output_files[str(number)].append(docx.name)
            if "pdf" in formats:
                final_pdf = scratch / (name + ".pdf")
                pdf_jobs.append((number, render_source, final_pdf))
                output_files[str(number)].append(final_pdf.name)
            report(f"Set {number} document prepared.", advance=True)
        if pdf_jobs:
            report("LibreOffice is converting the question papers to PDF…")
            converted = render_words([job[1] for job in pdf_jobs], scratch / "rendered", settings, app_folder)
            report("LibreOffice conversion complete.", advance=True)
            for (number, _, final_pdf), rendered in zip(pdf_jobs, converted):
                report(f"Applying the PDF margins and page labels for Set {number}…")
                finish_booklet_pdf(rendered, final_pdf, exam, number, settings)
                staged.append(final_pdf)
                report(f"Set {number} PDF complete.", advance=True)
        report("Saving the generated files…")
        if review_signature(folder, exam, settings) != signature:
            raise ValueError("The source or templates changed during generation. Review Set 0 again.")
        for item in staged:
            write_bytes(output / item.name, item.read_bytes())
    if is_zero:
        review = exam.get("review") or {}
        working_exam["generated"] = (exam.get("generated", {})
                                     if _review_matches(review, folder, exam, settings) else {})
        pdf = next(output / item.name for item in staged if item.suffix == ".pdf")
        working_exam["review"] = {
            "file": pdf.name, "files": output_files["0"], "schema": 2,
            "signature": signature, "pdf_digest": digest(pdf),
            "confirmed_at": "",
        }
    else:
        previous = exam.get("generated", {})
        working_exam["generated"] = {
            **previous, **{number: sorted(set(previous.get(number, []) + names))
                           for number, names in output_files.items()},
        }
    save_exam(folder, working_exam)
    exam.update(working_exam)
    report("Question files saved.", advance=True)
    return [output / name for names in output_files.values() for name in names]
