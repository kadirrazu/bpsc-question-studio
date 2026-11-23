"""Local exams, saved permutations and their stable print identifiers."""

import json
import math
import re
import secrets
import shutil
import string
from datetime import date as system_date, timedelta, timezone
from pathlib import Path
from .file_io import save_json

DHAKA = timezone(timedelta(hours=6))
RANDOM = secrets.SystemRandom()
BENGALI = str.maketrans("0123456789", "০১২৩৪৫৬৭৮৯")
VERSION = "0.5.1"
EXAM_TYPES = {"NON CADRE": "NC", "CADRE": "CADRE"}
DEFAULT_UNITS = [*(f"Unit {number:02d}" for number in range(1, 21)),
                 "Non Cadre (Exam)", "Cadre (Exam)", "Others"]
QR_FIELDS = {
    "set": "Set Number", "post": "Post Name", "post_code": "Post Code",
    "unit": "Unit", "date": "Exam Date", "timestamp": "Timestamp",
    "booklet": "Unique Booklet Code", "text": "Custom Text",
}
PATH_SETTINGS = ("header", "footer", "renderer", "output_root", "template_root")
DEFAULTS = {
    "paper": "Legal", "top": 0.5, "bottom": 0.5, "left": 0.5, "right": 0.5,
    "header_gap": 0.1, "footer_gap": 0.1, "qr_size": 0.75,
    "font": "Arial", "font_size": 9, "runner_text": "",
    "tag_language": "Bangla", "runner_tag_language": "Bangla",
    "sides": [], "runner_offset": 0.16, "runner_interval": 1.75,
    "credit": "Bangladesh Public Service Commission (BPSC)",
    "header": "", "footer": "", "renderer": "",
    "template_root": "", "output_root": "",
    "mapping_paper": "Legal", "mapping_orientation": "Portrait",
    **{f"mapping_margin_{side}": 0.5 for side in ("top", "right", "bottom", "left")},
    "qr_horizontal": "Template", "qr_vertical": "Template",
    "qr_inset": 0, "qr_gap": 0,
    "set_rotation": 0, "code_rotation": 0,
    "set_horizontal": "Template", "set_vertical": "Template",
    "code_horizontal": "Template", "code_vertical": "Template",
    "use_default_header": True, "use_default_footer": True,
    "runner_font": "Courier New", "runner_size": 10, "runner_color": "#353535",
    "runner_opacity": 0.8, "runner_uppercase": False,
    "footer_font": "", "footer_size": 0, "footer_color": "#6B7280",
    "footer_opacity": 1, "footer_uppercase": False,
    "print_code_enabled": True, "print_code_font": "Courier New", "print_code_size": 8,
    "print_code_color": "#000000", "print_code_opacity": 0.9,
    "footer_language": "English", "serial_suffix": ".",
    "qr_fields": ["set", "timestamp", "text"],
}


def load_settings(path):
    saved = json.loads(path.read_text("utf-8")) if path.exists() else {}
    settings_version = saved.pop("settings_version", None)
    if settings_version is None and saved.get("qr_size") == 0.65:
        saved["qr_size"] = DEFAULTS["qr_size"]
    return {**DEFAULTS, **saved}


def template_path(settings, kind, base=None):
    """Resolve the custom template before looking for an enabled default."""
    if settings[kind]:
        return Path(settings[kind])
    if not settings[f"use_default_{kind}"]:
        return None
    root = settings.get("template_root") or base
    if root:
        root = Path(root)
        name = f"DEFAULT_{kind.upper()}.docx"
        for path in (root / "Config" / "Template" / name, root / name):
            if path.is_file():
                return path
    return None


def text_appearance(settings, kind):
    return (settings[f"{kind}_font"] or settings["font"],
            settings[f"{kind}_size"] or settings["font_size"],
            settings[f"{kind}_color"], settings[f"{kind}_opacity"])


def footer_parts(number, settings):
    if settings["footer_language"] == "Bangla":
        return f"সেট - {label(number)}, পৃষ্ঠা ", "/"
    if settings["footer_uppercase"]:
        return f"SET {number} / PAGE ", " OF "
    return f"Set {number} / Page ", " of "


def load_units(path):
    if not path.exists():
        save_json(path, DEFAULT_UNITS)
    entries = json.loads(path.read_text("utf-8-sig"))
    if not isinstance(entries, list) or not entries or any(
        not isinstance(unit, str) or not unit.strip() for unit in entries
    ):
        raise ValueError("Config/units.json must contain a list of unit names.")
    units = [unit.strip() for unit in entries]
    if len(set(units)) != len(units):
        raise ValueError("Remove duplicate unit names from Config/units.json.")
    return units


def label(number):
    return str(number).translate(BENGALI)


def set_number_text(number, language):
    return label(number) if language == "Bangla" else str(number)


def make_mapping(count, sets):
    original = list(range(1, count + 1))
    if math.factorial(min(count, 15)) - (count > 1) < sets:
        raise ValueError("There are too few questions for that many different sets.")
    orders = [original]
    cycle = RANDOM.sample(original, count)
    inverse = {question: index for index, question in enumerate(cycle)}
    shifts = RANDOM.sample(range(1, count), count - 1)
    for shift in shifts[:sets]:
        orders.append([cycle[(inverse[question] + shift) % count] for question in original])
    # Beyond the number of questions, some repeated positions are unavoidable.
    # Keep the orders unique and prefer candidates with fewer reused positions.
    occupied = [{order[index] for order in orders} for index in range(count)]
    for order in orders[1:]:
        for _ in range(2 * count):
            a, b = RANDOM.sample(range(count), 2)
            left, right = order[a], order[b]
            if right not in occupied[a] and left not in occupied[b]:
                occupied[a].remove(left)
                occupied[a].add(right)
                occupied[b].remove(right)
                occupied[b].add(left)
                order[a], order[b] = right, left
    seen = set(map(tuple, orders)) if count > 1 else set()
    while len(orders) <= sets:
        candidates = []
        while not candidates:
            candidates = [order for _ in range(24)
                          if (order := tuple(RANDOM.sample(original, count))) not in seen]
        order = min(candidates, key=lambda candidate: sum(
            question in occupied[index] for index, question in enumerate(candidate)
        ))
        seen.add(order)
        for index, question in enumerate(order):
            occupied[index].add(question)
        orders.append(list(order))
    return orders


def generate_metadata(sets, used_codes):
    alphabet = string.ascii_uppercase + string.digits
    records = []
    for _ in range(sets + 1):
        while True:
            code = "".join(secrets.choice(alphabet) for _ in range(8))
            if code not in used_codes:
                used_codes.add(code)
                records.append({"code": code, "timestamp": ""})
                break
    return records


def read_exam(folder):
    return json.loads((folder / "exam.json").read_text("utf-8"))


def save_exam(folder, exam):
    save_json(folder / "exam.json", exam)


def exam_folders(data):
    folders = [file.parent for prefix in EXAM_TYPES.values()
               for file in data.glob(f"{prefix}_*/*/*/exam.json")]
    return sorted(
        folders,
        key=lambda path: path.stat().st_mtime, reverse=True,
    )


def safe_name(text):
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f\s]+', '_', text.strip()).strip(' ._').upper()[:75]
    reserved = {"CON", "PRN", "AUX", "NUL", *(f"COM{n}" for n in range(1, 10)),
                *(f"LPT{n}" for n in range(1, 10))}
    if safe.split(".")[0].upper() in reserved:
        safe = "_" + safe
    return safe or "_"


def create_exam(data, unit, post_code, title, code_name, count, sets, date=None, exam_type="NON CADRE", *, grade=None):
    if not all((unit, title, code_name)):
        raise ValueError("Select a unit and enter the post name and the question’s code name.")
    if post_code and not re.fullmatch(r"[0-9]{6}", post_code):
        raise ValueError("Enter a six-digit post code or leave the field blank.")
    if grade is not None and (not re.fullmatch(r"[0-9]+", str(grade)) or int(grade) < 1):
        raise ValueError("Enter a positive whole number for the post grade or leave it blank.")
    today = system_date.today()
    date = date or today.isoformat()
    try:
        system_date.fromisoformat(date)
    except ValueError as error:
        raise ValueError("Use an exam date in YYYY-MM-DD format.") from error
    exam_type = exam_type.strip().upper()
    if exam_type not in EXAM_TYPES:
        raise ValueError("Select a valid exam type.")
    unit_folder = safe_name(unit)
    if unit_folder.isdecimal():
        unit_folder = "UNIT_" + unit_folder
    folder = data / f"{EXAM_TYPES[exam_type]}_{today.year}" / unit_folder / f"{date}-{safe_name(title)}"
    suffix = post_code or "NO_POST_CODE"
    alternate = folder.with_name(f"{folder.name}-{suffix}")
    duplicate = "An exam already exists for this date, post and post code. Open it from Saved Exams."
    if (alternate / "exam.json").exists():
        raise ValueError(duplicate)
    if (folder / "exam.json").is_file():
        saved = read_exam(folder)
        if str(saved.get("post_code") or "") == post_code:
            raise ValueError(duplicate)
        folder = alternate
    folder.mkdir(parents=True, exist_ok=True)
    exam = {
        "unit": unit, "post_code": post_code, "title": title, "post_grade": int(grade) if grade is not None else None,
        "code_name": code_name, "count": count, "sets": sets, "date": date, "year": today.year, "exam_type": exam_type,
        "mapping": [], "records": [], "source": "",
    }
    path = folder / "exam.json"
    try:
        manifest = path.open("x", encoding="utf-8")
    except FileExistsError as error:
        raise ValueError(duplicate) from error
    try:
        with manifest:
            json.dump(exam, manifest, indent=2, ensure_ascii=False)
    except OSError:
        path.unlink(missing_ok=True)
        raise
    return folder, exam


def prepare_mapping(data, exam):
    orders = make_mapping(exam["count"], exam["sets"])
    used = {
        record["code"]
        for other in exam_folders(data)
        for record in read_exam(other)["records"]
    }
    for other in exam_folders(data):
        for previous in (other / "MAPPING").glob("mapping-*.json"):
            used.update(row["code"] for row in json.loads(previous.read_text("utf-8"))["records"])
    return {**exam, "mapping": orders, "records": generate_metadata(exam["sets"], used),
            "review": {}, "generated": {}}


def commit_mapping(folder, updated):
    save_json(folder / "MAPPING" / f"mapping-{updated['records'][0]['code']}.json", {
        "columns": list(range(updated["sets"] + 1)),
        "rows": mapping_rows(updated),
        "records": updated["records"],
    })
    save_exam(folder, updated)
    return updated


def validate_mapping(exam):
    original = list(range(1, exam["count"] + 1))
    orders = exam["mapping"]
    if len(orders) != exam["sets"] + 1 or not orders or orders[0] != original:
        raise ValueError("The saved mapping is incomplete. Regenerate it before continuing.")
    if any(sorted(order) != original for order in orders):
        raise ValueError("The saved mapping contains a repeated or missing question.")
    if exam["count"] > 1 and len(set(map(tuple, orders))) != len(orders):
        raise ValueError("The saved mapping contains duplicate set orders.")
    codes = [record["code"] for record in exam["records"]]
    if len(codes) != len(orders) or len(set(codes)) != len(codes):
        raise ValueError("The saved mapping has missing or duplicate booklet codes.")


def mapping_rows(exam):
    validate_mapping(exam)
    positions = [
        {question: number for number, question in enumerate(order, 1)}
        for order in exam["mapping"][1:]
    ]
    return [[number, *(column[number] for column in positions)]
            for number in range(1, exam["count"] + 1)]


def post_name(exam):
    title, grade = exam["title"], exam.get("post_grade")
    return f"{title} (Grade: {grade})" if grade not in (None, "") else title


def file_prefix(exam):
    post = re.sub(r'[^\w-]', '_', exam["post_code"], flags=re.UNICODE)
    return f"EXAM-{post[:30].upper()}" if post else "EXAM"


def delete_exam(data, folder):
    root = data.resolve()
    target = folder.resolve()
    if (not target.is_relative_to(root) or len(target.relative_to(root).parts) != 3
            or target != root / folder.relative_to(data) or folder not in exam_folders(data)):
        raise ValueError("This folder is not a saved exam in the configured location.")
    shutil.rmtree(target)
