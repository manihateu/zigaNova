"""Сборка отчётов DOCX из Markdown-файлов.

Поддерживается ограниченное подмножество Markdown: заголовки, абзацы,
маркированные списки, таблицы, блоки кода, изображения и простое
внутреннее форматирование (**жирный**, *курсив*, `код`).
Титульная страница задаётся front matter между двумя строками ---.
"""
from pathlib import Path
import re
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = 'Times New Roman'
BODY_SIZE = Pt(12)


def set_run_font(run, size=BODY_SIZE, bold=None, italic=None, name=FONT):
    run.font.name = name
    run.font.size = size
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    for attr in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rfonts.set(qn(attr), name)


def add_page_number(paragraph):
    run = paragraph.add_run()
    fld_begin = OxmlElement('w:fldChar')
    fld_begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'PAGE'
    fld_end = OxmlElement('w:fldChar')
    fld_end.set(qn('w:fldCharType'), 'end')
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)
    set_run_font(run, Pt(11))


def shade(cell, color='EDEDED'):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tc_pr.append(shd)


def parse_front_matter(text):
    lines = text.splitlines()
    meta = {}
    if lines and lines[0].strip() == '---':
        try:
            end = next(i for i, line in enumerate(lines[1:], 1) if line.strip() == '---')
        except StopIteration:
            return {}, text
        for line in lines[1:end]:
            if ':' in line:
                key, value = line.split(':', 1)
                meta[key.strip()] = value.strip()
        return meta, '\n'.join(lines[end + 1:])
    return meta, text


def add_title_page(doc, meta):
    def centered(text, size, bold=False, space_after=0):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(space_after)
        r = p.add_run(text)
        set_run_font(r, size, bold=bold)
        return p

    centered(meta.get('ministry', 'Министерство науки и высшего образования Российской Федерации'),
             Pt(12))
    centered(meta.get('university', 'Ульяновский государственный технический университет'), Pt(12))
    centered(meta.get('department', 'Кафедра «Информационные системы»'), Pt(12), space_after=24)
    for _ in range(3):
        doc.add_paragraph()
    centered(meta.get('doctype', 'ОТЧЁТ'), Pt(20), bold=True, space_after=6)
    centered(meta.get('course', 'по лабораторному практикуму'), Pt(14), space_after=6)
    centered(meta.get('course_name', '«Стохастические модели и оценки»'), Pt(14), space_after=24)
    centered(meta.get('title', 'Лабораторная работа'), Pt(18), bold=True, space_after=6)
    centered(meta.get('variant', ''), Pt(14), space_after=36)
    for _ in range(2):
        doc.add_paragraph()

    info = [
        ('Выполнил:', meta.get('author', '____________________________')),
        ('Группа:', meta.get('group', '____________________________')),
        ('Проверил:', meta.get('teacher', '____________________________')),
    ]
    for label, value in info:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p.paragraph_format.space_after = Pt(6)
        r1 = p.add_run(label + ' ')
        set_run_font(r1, Pt(12))
        r2 = p.add_run(value)
        set_run_font(r2, Pt(12))

    for _ in range(5):
        doc.add_paragraph()
    centered(meta.get('city', 'Ульяновск'), Pt(12), space_after=0)
    centered(meta.get('year', '2026'), Pt(12))
    doc.add_page_break()


INLINE = re.compile(r'(\*\*.+?\*\*|\*[^*]+?\*|`[^`]+?`|\$[^$]+?\$)')


def add_inline(paragraph, text, size=BODY_SIZE, base_bold=False, base_italic=False):
    for chunk in INLINE.split(text):
        if not chunk:
            continue
        bold, italic, mono = base_bold, base_italic, False
        value = chunk
        if chunk.startswith('**') and chunk.endswith('**') and len(chunk) > 4:
            value, bold = chunk[2:-2], True
        elif chunk.startswith('$') and chunk.endswith('$') and len(chunk) > 2:
            value, italic = chunk[1:-1], True
        elif chunk.startswith('*') and chunk.endswith('*') and len(chunk) > 2:
            value, italic = chunk[1:-1], True
        elif chunk.startswith('`') and chunk.endswith('`') and len(chunk) > 2:
            value, mono = chunk[1:-1], True
        run = paragraph.add_run(value)
        set_run_font(run, size, bold=bold, italic=italic,
                     name='Consolas' if mono else FONT)


def add_body_paragraph(doc, text, style=None, size=BODY_SIZE, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
                       space_after=Pt(6), first_line=Cm(1.25)):
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_after = space_after
    p.paragraph_format.line_spacing = 1.5
    if first_line is not None:
        p.paragraph_format.first_line_indent = first_line
    add_inline(p, text, size)
    return p


def add_heading(doc, text, level):
    sizes = {1: Pt(16), 2: Pt(14), 3: Pt(13)}
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14 if level <= 2 else 10)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    if level == 1:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text)
    set_run_font(r, sizes.get(level, Pt(12)), bold=True,
                 italic=(level >= 3))
    return p


def add_table(doc, rows):
    if not rows:
        return
    cols = max(len(r) for r in rows)
    table = doc.add_table(rows=0, cols=cols)
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        cells = table.add_row().cells
        for j in range(cols):
            value = row[j] if j < len(row) else ''
            cell = cells[j]
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(2)
            add_inline(p, value, Pt(10), base_bold=(i == 0))
            if i == 0:
                shade(cell)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def add_image(doc, path, caption):
    if not path.exists():
        p = doc.add_paragraph()
        r = p.add_run(f'[не найден рисунок: {path}]')
        set_run_font(r, Pt(10), italic=True)
        return
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(2)
    run = p.add_run()
    run.add_picture(str(path), width=Cm(16))
    if caption:
        cap = doc.add_paragraph()
        cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        cap.paragraph_format.space_after = Pt(10)
        r = cap.add_run(caption)
        set_run_font(r, Pt(10), italic=True)


def convert(md_path, out_path):
    text = md_path.read_text(encoding='utf-8')
    meta, body = parse_front_matter(text)
    doc = Document()

    section = doc.sections[0]
    section.left_margin = Cm(3)
    section.right_margin = Cm(1.5)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)

    normal = doc.styles['Normal']
    normal.font.name = FONT
    normal.font.size = BODY_SIZE
    normal.element.rPr.rFonts.set(qn('w:eastAsia'), FONT)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_page_number(footer)

    add_title_page(doc, meta)

    lines = body.splitlines()
    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if not stripped:
            i += 1
            continue

        if stripped == '$$':
            i += 1
            block = []
            while i < len(lines) and lines[i].strip() != '$$':
                block.append(lines[i])
                i += 1
            i += 1
            for formula in block:
                if not formula.strip():
                    continue
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(2)
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                r = p.add_run(formula.strip())
                set_run_font(r, Pt(12), name='Cambria Math')
            doc.add_paragraph().paragraph_format.space_after = Pt(0)
            continue

        if stripped.startswith('```'):
            i += 1
            block = []
            while i < len(lines) and not lines[i].strip().startswith('```'):
                block.append(lines[i])
                i += 1
            i += 1
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.left_indent = Cm(0.75)
            p.paragraph_format.space_after = Pt(8)
            p.paragraph_format.line_spacing = 1.0
            r = p.add_run('\n'.join(block))
            set_run_font(r, Pt(10), name='Consolas')
            continue

        heading = re.match(r'^(#{1,4})\s+(.*)$', stripped)
        if heading:
            add_heading(doc, heading.group(2).strip(), len(heading.group(1)))
            i += 1
            continue

        image = re.match(r'^!\[(.*?)\]\((.*?)\)\s*$', stripped)
        if image:
            rel = image.group(2).strip()
            path = (md_path.parent / rel).resolve()
            add_image(doc, path, image.group(1).strip())
            i += 1
            continue

        if stripped.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                if not all(re.fullmatch(r':?-{2,}:?', c) for c in cells if c):
                    rows.append(cells)
                i += 1
            add_table(doc, rows)
            continue

        if stripped in ('---', '***'):
            i += 1
            continue

        bullet = re.match(r'^[-*]\s+(.*)$', stripped)
        if bullet:
            add_body_paragraph(doc, bullet.group(1), style='List Bullet',
                               first_line=None, space_after=Pt(2))
            i += 1
            continue

        numbered = re.match(r'^(\d+)\.\s+(.*)$', stripped)
        if numbered:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.left_indent = Cm(0.75)
            p.paragraph_format.first_line_indent = Cm(-0.5)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.5
            r = p.add_run(numbered.group(1) + '. ')
            set_run_font(r, BODY_SIZE)
            add_inline(p, numbered.group(2))
            i += 1
            continue

        if stripped.startswith('> '):
            p = doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            p.paragraph_format.right_indent = Cm(1)
            p.paragraph_format.space_after = Pt(6)
            add_inline(p, stripped[2:], Pt(11), base_italic=True)
            i += 1
            continue

        add_body_paragraph(doc, stripped)
        i += 1

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))
    return out_path


def main(argv):
    paths = [Path(p) for p in argv[1:]]
    if not paths:
        paths = sorted(Path(__file__).resolve().parent.glob('*_report.md'))
    for md in paths:
        docx = md.with_suffix('.docx')
        convert(md, docx)
        print(f'{md.name} -> {docx.name}')


if __name__ == '__main__':
    main(sys.argv)
