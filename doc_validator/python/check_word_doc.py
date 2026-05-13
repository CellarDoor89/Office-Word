"""Проверка оформления Word-документа по «Памятке».

Использует python-docx (читает .docx как ZIP-архив, не требует Word).

    python check_word_doc.py "путь/к/файлу.docx"

Без аргументов — берёт первый .docx в текущей папке.
"""
from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import List

from docx import Document
from docx.shared import Emu
from docx.enum.text import WD_LINE_SPACING

# ---------- параметры из Памятки ----------
MARGIN_LEFT_MM = 30
MARGIN_TOP_MM = 20
MARGIN_BOTTOM_MM = 20
MARGIN_RIGHT_MM = 10
MM_TOL = 0.5
CM_TOL = 0.05
FONT_REQ = "Times New Roman"
INDENT_REQ_CM = 1.25
SIZES_OK = {13, 14}

EMU_PER_MM = 36000
EMU_PER_CM = 360000


def emu_to_mm(v) -> float:
    if v is None:
        return 0.0
    return round(int(v) / EMU_PER_MM, 2)


def emu_to_cm(v) -> float:
    if v is None:
        return 0.0
    return round(int(v) / EMU_PER_CM, 3)


def check_margins(doc, issues: List[str]) -> None:
    s = doc.sections[0]
    pairs = [
        ("Левое",   emu_to_mm(s.left_margin),   MARGIN_LEFT_MM),
        ("Верхнее", emu_to_mm(s.top_margin),    MARGIN_TOP_MM),
        ("Нижнее",  emu_to_mm(s.bottom_margin), MARGIN_BOTTOM_MM),
        ("Правое",  emu_to_mm(s.right_margin),  MARGIN_RIGHT_MM),
    ]
    for name, actual, required in pairs:
        if abs(actual - required) > MM_TOL:
            issues.append(f"{name} поле {actual} мм (требуется {required} мм).")


def _run_font(run) -> str | None:
    name = run.font.name
    if name:
        return name
    # fallback в style
    try:
        return run.style.font.name
    except Exception:
        return None


def check_fonts_sizes_spacing_indent(doc, issues: List[str]) -> None:
    bad_fonts: set[str] = set()
    bad_sizes: set[str] = set()
    bad_spacing = False
    body_paras = 0
    bad_indent = 0

    for p in doc.paragraphs:
        text = p.text.strip()
        if len(text) <= 1:
            continue

        # шрифт + размер
        for run in p.runs:
            name = _run_font(run)
            if name and name != FONT_REQ:
                bad_fonts.add(name)
            sz = run.font.size
            if sz is not None:
                pt = sz.pt
                if pt not in SIZES_OK:
                    bad_sizes.add(f"{pt:g}")

        # межстрочный
        pf = p.paragraph_format
        rule = pf.line_spacing_rule
        ls = pf.line_spacing
        ok = False
        if rule in (WD_LINE_SPACING.SINGLE, WD_LINE_SPACING.ONE_POINT_FIVE):
            ok = True
        elif rule == WD_LINE_SPACING.MULTIPLE and ls is not None:
            if 0.95 <= float(ls) <= 1.55:
                ok = True
        elif rule is None and ls is None:
            ok = True  # дефолт стиля — считаем нормальным
        elif ls is not None and 0.95 <= float(ls) <= 1.55:
            ok = True
        if not ok:
            bad_spacing = True

        # отступ — только в больших абзацах не-заголовков
        style_name = (p.style.name or "").lower()
        if len(text) > 80 and "heading" not in style_name and "заголовок" not in style_name:
            body_paras += 1
            ind = pf.first_line_indent
            ind_cm = emu_to_cm(ind) if ind is not None else 0.0
            if abs(ind_cm - INDENT_REQ_CM) > CM_TOL:
                bad_indent += 1

    if bad_fonts:
        issues.append(f"Используется не {FONT_REQ}: {', '.join(sorted(bad_fonts))}.")
    if bad_sizes:
        issues.append(f"Размер шрифта вне 13–14 пт: {', '.join(sorted(bad_sizes))}.")
    if bad_spacing:
        issues.append("Межстрочный интервал вне диапазона 1.0–1.5.")
    if body_paras and bad_indent:
        issues.append(
            f"Абзацный отступ ≠ 1.25 см в {bad_indent} из {body_paras} абзацев основного текста."
        )


def check_page_numbers(doc, issues: List[str]) -> None:
    # python-docx не считает страницы. Используем эвристику: документ длиннее ~40 абзацев.
    if len(doc.paragraphs) < 40:
        return
    found_page_field = False
    centered = False
    decor = False
    for section in doc.sections:
        for hdr in (section.header, section.first_page_header):
            xml = hdr._element.xml if hdr is not None else ""
            if "PAGE" in xml and 'instr' in xml.lower():
                found_page_field = True
            # упрощённо: ищем выравнивание по центру в первом параграфе заголовка
            for p in hdr.paragraphs:
                if p.alignment is not None and int(p.alignment) == 1:
                    centered = True
                t = p.text
                if "-" in t or "стр" in t.lower() or "с." in t.lower():
                    decor = True
    if not found_page_field:
        issues.append("Похоже на многостраничный документ, но номер страницы в колонтитуле не найден.")
    else:
        if not centered:
            issues.append("Номер страницы должен быть по центру верхнего поля.")
        if decor:
            issues.append("В номере страницы есть лишние символы (тире / «стр.» / «с.»).")


def check_executor(doc, issues: List[str]) -> None:
    full = "\n".join(p.text for p in doc.paragraphs)
    if not re.search(r"8\(\d{3,5}\)\s?\d", full):
        issues.append("Не найдена отметка об исполнителе (телефон в формате 8(код)номер).")


def check_restricted_mark(doc, issues: List[str]) -> None:
    full = "\n".join(p.text for p in doc.paragraphs).lower()
    if "для служебного пользования" in full:
        if not re.search(r"экз\.?\s*№", full):
            issues.append(
                "Указано «Для служебного пользования», но не найден номер экземпляра («Экз. № …»)."
            )


def validate(path: Path) -> List[str]:
    doc = Document(str(path))
    issues: List[str] = []
    check_margins(doc, issues)
    check_fonts_sizes_spacing_indent(doc, issues)
    check_page_numbers(doc, issues)
    check_executor(doc, issues)
    check_restricted_mark(doc, issues)
    return issues


def main() -> int:
    if len(sys.argv) > 1:
        path = Path(sys.argv[1])
    else:
        candidates = sorted(Path.cwd().glob("*.docx"))
        if not candidates:
            print("Укажите путь к .docx или положите файл в текущую папку.")
            return 2
        path = candidates[0]

    if not path.exists():
        print(f"Файл не найден: {path}")
        return 2

    issues = validate(path)

    print(f"\nДокумент: {path}")
    print("-" * 60)
    if not issues:
        print("OK. Все автоматические проверки пройдены.")
        return 0
    for i, msg in enumerate(issues, 1):
        print(f"{i}. {msg}")
    print(f"\nНайдено замечаний: {len(issues)}")
    return 1


if __name__ == "__main__":
    sys.exit(main())
