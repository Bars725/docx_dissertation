"""Фикстуры: образцы .docx, имитирующие файлы из Word / LibreOffice / Google Docs."""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

import pytest
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Mm, Pt

from dissertation import ooxml


def _split_runs(paragraph, pieces, bold_first=False):
    for i, piece in enumerate(pieces):
        r = paragraph.add_run(piece)
        if i == 0 and bold_first:
            r.bold = True


@pytest.fixture
def word_like_template(tmp_path) -> Path:
    """Образец «как из Word»: шрифты темы в docDefaults, оформление в стилях,
    «Глава N.» в заголовках, номер страницы вверху, [N] в списке литературы,
    раздробленные runs и немного прямого форматирования."""
    doc = Document()
    s = doc.sections[0]
    s.page_width, s.page_height = Mm(210), Mm(297)
    s.left_margin, s.right_margin, s.top_margin, s.bottom_margin = Mm(25), Mm(10), Mm(20), Mm(20)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"          # python-docx пишет только ascii/hAnsi
    normal.font.size = Pt(13)
    pf = normal.paragraph_format
    pf.line_spacing = 2.0
    pf.first_line_indent = Cm(1.0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    h1 = doc.styles["Heading 1"]
    h1.font.size = Pt(16)
    h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    h1.paragraph_format.first_line_indent = Cm(0)
    h1.paragraph_format.page_break_before = True
    h2 = doc.styles["Heading 2"]
    h2.font.size = Pt(14)

    hp = s.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ooxml.append_field(hp, "PAGE")
    s.different_first_page_header_footer = True

    doc.add_paragraph("МИНИСТЕРСТВО НАУКИ").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("ДИССЕРТАЦИЯ").alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_heading("ВВЕДЕНИЕ", level=1)
    long = ("Актуальность темы исследования обусловлена необходимостью разработки "
            "методов, которые позволят повысить качество оформления научных работ. ")
    doc.add_paragraph(long * 2)
    doc.add_heading("Глава 1. Обзор литературы", level=1)
    p = doc.add_paragraph()
    _split_runs(p, ["Раздроблен", "ный ", "абзац: " + long * 2], bold_first=True)
    doc.add_heading("1.1. Подтема", level=2)
    doc.add_paragraph(long * 3)
    doc.add_heading("Глава 2. Методы", level=1)
    doc.add_paragraph(long * 2)
    doc.add_heading("СПИСОК ЛИТЕРАТУРЫ", level=1)
    doc.add_paragraph("[1] Иванов И. И. Оформление диссертаций. — М. : Наука, 2020. — 200 с.")
    doc.add_paragraph("[2] Петров П. П. Стандарты. — СПб. : Питер, 2021. — 150 с.")
    path = tmp_path / "word_like.docx"
    doc.save(str(path))
    return path


@pytest.fixture
def placeholder_doc(tmp_path) -> Path:
    """Плейсхолдеры, разрезанные на runs, — в теле, в таблице и в колонтитуле."""
    doc = Document()
    p = doc.add_paragraph()
    for piece in ("Тема: {{TI", "TLE", "}} — конец"):
        p.add_run(piece)
    p.runs[0].bold = True
    cell = doc.add_table(rows=1, cols=1).cell(0, 0)
    cp = cell.paragraphs[0]
    cp.add_run("{{AUT")
    cp.add_run("HOR}}")
    hp = doc.sections[0].header.paragraphs[0]
    hp.add_run("Город {{CI")
    hp.add_run("TY}}")
    # абзац с табуляцией внутри «плейсхолдера» — заменяться не должен
    tp = doc.add_paragraph()
    tp.add_run("{{TI")
    tab_run = tp.add_run()
    tab_run._r.append(OxmlElement("w:tab"))
    tp.add_run("TLE}}")
    path = tmp_path / "placeholders.docx"
    doc.save(str(path))
    return path


def soffice() -> str | None:
    return shutil.which("soffice") or shutil.which("libreoffice")


def libreoffice_convert(src: Path, fmt: str, outdir: Path) -> Path:
    """Пересохранить файл через LibreOffice (имитация «открыли и сохранили в LO»)."""
    outdir.mkdir(parents=True, exist_ok=True)
    subprocess.run([soffice(), "--headless", "--norestore", f"-env:UserInstallation=file://{outdir}/lo_profile",
                    "--convert-to", fmt, "--outdir", str(outdir), str(src)],
                   check=True, capture_output=True, timeout=180)
    ext = fmt.split(":")[0]
    return outdir / f"{src.stem}.{ext}"


def w_attr(el, tag, attr):
    child = el.find(qn(tag))
    return None if child is None else child.get(qn(attr))
