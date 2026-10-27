# -*- coding: utf-8 -*-
"""
build_reports.py — сборка 4 отчётов по кейсам дисциплины
«Искусственный интеллект, большие данные и этика» (ГОСТ 7.32-2017,
клон оформления ВКР Синергии).

Структура отчёта: титульный лист (РЕШЕНИЕ КЕЙСА № N) -> СОДЕРЖАНИЕ
(ручное, с точечными лидерами) -> ВВЕДЕНИЕ -> разделы 1–3 -> ВЫВОДЫ ->
СПИСОК ИСТОЧНИКОВ. Содержимое берётся из cases/*/content.py.

Режимы:
    python3 tools/build_reports.py                 # сборка всех docx
    python3 tools/build_reports.py --case 1        # только кейс 1
    python3 tools/build_reports.py --update-toc    # проставить фактические
        # номера страниц в СОДЕРЖАНИИ по сконвертированным PDF (itog/*.pdf)
        # и пересобрать docx

Типовой цикл: build -> convert_pdf.sh -> build --update-toc ->
convert_pdf.sh. Файлы складываются в cases/*/output/ и копируются
в itog/ под финальными именами сдачи.
"""

import argparse
import importlib.util
import json
import re
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "templates"))

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, \
    WD_TAB_LEADER  # noqa: E402
from docx.oxml import OxmlElement  # noqa: E402
from docx.oxml.ns import qn  # noqa: E402
from docx.shared import Cm, Pt  # noqa: E402

import synergy_docx as sd  # noqa: E402

DISCIPLINE = "Искусственный интеллект, большие данные и этика"
GROUP = "ДМЭ-101цэии"
STUDENT = "Шатов Илья Андреевич"
YEAR = 2026
WORK_TYPE = "РЕШЕНИЕ КЕЙСА"

TOC_TITLE = "СОДЕРЖАНИЕ"
USABLE_WIDTH_CM = 16.4

CASES = [
    {"slug": "case-01-fair-hr", "number": 1,
     "file": "ДМЭ-101ЦЭИИ.Шатов.Кейс_1_Беспристрастный_HR"},
    {"slug": "case-02-medical-assistant", "number": 2,
     "file": "ДМЭ-101ЦЭИИ.Шатов.Кейс_2_Медицинский_ассистент"},
    {"slug": "case-03-ai-code-copyright", "number": 3,
     "file": "ДМЭ-101ЦЭИИ.Шатов.Кейс_3_Код_без_автора"},
    {"slug": "case-04-deepfake-biometrics", "number": 4,
     "file": "ДМЭ-101ЦЭИИ.Шатов.Кейс_4_Deepfake_проверка"},
]

PAGEMAP = ROOT / "itog" / "_pagemap.json"


def load_content(case):
    path = ROOT / "cases" / case["slug"] / "content.py"
    spec = importlib.util.spec_from_file_location("case_content", path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


# ------------------------------------------------------------------
# СОДЕРЖАНИЕ (ручное, с точечными лидерами)
# ------------------------------------------------------------------

def toc_entry(doc, text, page, bold=False):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.5
    pf.space_after = Pt(0)
    pf.tab_stops.add_tab_stop(Cm(USABLE_WIDTH_CM), WD_TAB_ALIGNMENT.RIGHT,
                              WD_TAB_LEADER.DOTS)
    r = p.add_run(text)
    r.bold = bold
    r.font.name = sd.TNR
    r.font.size = Pt(14)
    r2 = p.add_run("\t" + str(page))
    r2.bold = bold
    r2.font.name = sd.TNR
    r2.font.size = Pt(14)
    return p


def toc_title(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pf = p.paragraph_format
    pf.page_break_before = True
    pf.first_line_indent = Cm(0)
    pf.line_spacing = 1.0
    pf.space_after = Pt(12)
    r = p.add_run(TOC_TITLE)
    r.bold = True
    r.font.name = sd.TNR
    r.font.size = Pt(16)
    return p


def load_pagemap():
    if PAGEMAP.exists():
        return json.loads(PAGEMAP.read_text(encoding="utf-8"))
    return {}


def norm(s):
    return re.sub(r"\s+", " ", s).strip().lower()


def render_blocks(r, blocks):
    for block in blocks:
        kind = block[0]
        if kind == "p":
            r.paragraph(block[1])
        elif kind == "pi":
            r.paragraph(block[1], italic=True)
        elif kind == "bullets":
            for item in block[1]:
                r.paragraph("– " + item)
        elif kind == "numbered":
            for i, item in enumerate(block[1], 1):
                r.paragraph("%d. %s" % (i, item))
        elif kind == "table":
            r.table(block[1]["rows"], caption=block[1]["caption"],
                    header=True)
        elif kind == "conclusion":
            r.conclusion(block[1])
        else:
            raise ValueError("Неизвестный блок: %r" % (kind,))


def build_case(case, pagemap):
    c = load_content(case)
    r = sd.SynergyReport(
        discipline=DISCIPLINE,
        topic=c.CASE_TITLE,
        lab_number=case["number"],
        work_type=WORK_TYPE,
        student=STUDENT,
        group=GROUP,
        teacher="",
        year=YEAR,
    )

    # пустой абзац-разрыв из __init__ не нужен: СОДЕРЖАНИЕ получит
    # page_break_before
    tail = r.doc.paragraphs[-1]
    if not tail.text.strip():
        tail._p.getparent().remove(tail._p)

    # --- СОДЕРЖАНИЕ ---
    toc_title(r.doc)
    entries = [("ВВЕДЕНИЕ", True)]
    entries += [(t, True) for t in c.SECTION_TITLES]
    entries += [("ВЫВОДЫ", True), ("СПИСОК ИСТОЧНИКОВ", True)]
    for text, _bold in entries:
        page = pagemap.get(norm(text), "0")
        toc_entry(r.doc, text, page)

    # --- ВВЕДЕНИЕ ---
    r.heading("ВВЕДЕНИЕ", 1)
    for para in c.INTRO:
        r.paragraph(para)

    # --- разделы ---
    for section in c.SECTIONS:
        r.heading(section["title"], 1)
        render_blocks(r, section["blocks"])

    # --- ВЫВОДЫ ---
    r.heading("ВЫВОДЫ", 1)
    for para in c.CONCLUSIONS:
        r.paragraph(para)

    # --- СПИСОК ИСТОЧНИКОВ ---
    r.heading("СПИСОК ИСТОЧНИКОВ", 1)
    for i, ref in enumerate(c.REFERENCES, 1):
        p = r.doc.add_paragraph(style="Body Text")
        p.add_run("%d. %s" % (i, ref))
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.left_indent = Cm(0.75)

    # --- финализация (как apply_title_page.py) ---
    add_page_numbers(r.doc)
    sd.sanitize_fonts(r.doc)
    enforce_alignment(r.doc)

    out_case = ROOT / "cases" / case["slug"] / "output" / (case["file"] + ".docx")
    out_case.parent.mkdir(parents=True, exist_ok=True)
    r.save(str(out_case))

    itog = ROOT / "itog" / (case["file"] + ".docx")
    itog.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(out_case, itog)
    return out_case, itog


def add_page_numbers(doc):
    section = doc.sections[0]
    section.different_first_page_header_footer = True
    sect_pr = section._sectPr
    for el in sect_pr.findall(qn("w:pgNumType")):
        sect_pr.remove(el)
    pg = OxmlElement("w:pgNumType")
    pg.set(qn("w:start"), "1")
    sect_pr.append(pg)

    footer = section.footer
    footer.is_linked_to_previous = False
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    for run in list(p.runs):
        run._r.getparent().remove(run._r)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run()
    run.font.name = sd.TNR
    run.font.size = Pt(14)
    f1 = OxmlElement("w:fldChar")
    f1.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    f2 = OxmlElement("w:fldChar")
    f2.set(qn("w:fldCharType"), "end")
    run._r.append(f1)
    run._r.append(instr)
    run._r.append(f2)

    first = section.first_page_footer
    first.is_linked_to_previous = False
    fp = first.paragraphs[0] if first.paragraphs else first.add_paragraph()
    for run in list(fp.runs):
        run._r.getparent().remove(run._r)


def enforce_alignment(doc):
    n = 0
    for p in doc.paragraphs:
        sid = p.style.style_id if p.style is not None else ""
        if sid in ("BodyText", "FirstParagraph") and p.text.strip():
            pf = p.paragraph_format
            pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            if pf.first_line_indent is None or pf.first_line_indent.cm == 0:
                pf.first_line_indent = Cm(1.25)
            pf.line_spacing = 1.5
            n += 1
    return n


# ------------------------------------------------------------------
# --update-toc: фактические номера страниц по PDF
# ------------------------------------------------------------------

def update_toc_from_pdf(pdf_path, titles):
    from pypdf import PdfReader
    pages = [norm(p.extract_text() or "") for p in PdfReader(pdf_path).pages]

    toc_page = 0
    for i, ptext in enumerate(pages, start=1):
        if "содержание" in ptext:
            toc_page = i
            break

    def find_page(title):
        t = norm(title)
        for i, ptext in enumerate(pages, start=1):
            if i <= toc_page:
                continue
            lines = [norm(x) for x in re.split(r"[\n\r]+", ptext)]
            if t in lines:
                return i
        for i, ptext in enumerate(pages, start=1):
            if i <= toc_page:
                continue
            if t in ptext:
                return i
        return None

    result = {}
    for title in titles:
        page = find_page(title)
        if page is not None:
            result[norm(title)] = page
    return result


def collect_titles(c):
    return ["ВВЕДЕНИЕ"] + list(c.SECTION_TITLES) + \
        ["ВЫВОДЫ", "СПИСОК ИСТОЧНИКОВ"]


def update_toc(only=None):
    """Номера страниц — по-кейсные (общие ключи «ВВЕДЕНИЕ» и т.д.
    у кейсов разные, поэтому один общий pagemap перезаписывал бы их).

    Возвращает {case_number: {norm_title: page}}; каждый найденный
    кейс сразу пересобирается со своими номерами (раскладка после
    пересборки не меняется — заменяются только цифры в строках)."""
    all_maps = {}
    missing = []
    for case in CASES:
        if only and case["number"] not in only:
            continue
        pdf = ROOT / "itog" / (case["file"] + ".pdf")
        if not pdf.exists():
            missing.append(pdf.name)
            continue
        c = load_content(case)
        found = update_toc_from_pdf(pdf, collect_titles(c))
        all_maps[case["number"]] = found
        build_case(case, found)
        not_found = [t for t in collect_titles(c) if norm(t) not in found]
        if not_found:
            missing.extend("%s: %s" % (case["file"], t) for t in not_found)
    PAGEMAP.write_text(json.dumps(all_maps, ensure_ascii=False, indent=1),
                       encoding="utf-8")
    return all_maps, missing


def load_case_pagemap(case):
    """По-кейсная карта страниц из файла; без файла — пустая."""
    if PAGEMAP.exists():
        data = json.loads(PAGEMAP.read_text(encoding="utf-8"))
        if isinstance(data, dict) and "page" not in data:
            sub = data.get(str(case["number"]))
            if isinstance(sub, dict):
                return sub
    return {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--case", type=int, choices=[c["number"] for c in CASES],
                    help="собрать один кейс по номеру")
    ap.add_argument("--update-toc", action="store_true",
                    help="обновить СОДЕРЖАНИЕ по PDF из itog/ и пересобрать")
    args = ap.parse_args()

    only = {args.case} if args.case else None

    if args.update_toc:
        all_maps, missing = update_toc(only)
        print("pagemap:", json.dumps(all_maps, ensure_ascii=False))
        if missing:
            print("ВНИМАНИЕ, не найдено в PDF:", "; ".join(missing))
        return  # кейсы уже пересобраны со своими номерами

    for case in CASES:
        if only and case["number"] not in only:
            continue
        out_case, out_itog = build_case(case, load_case_pagemap(case))
        print("OK:", out_itog.relative_to(ROOT),
              "(%s)" % out_case.parent.name)


if __name__ == "__main__":
    main()
