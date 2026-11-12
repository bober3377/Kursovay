from io import BytesIO

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor


NAVY = "17365D"
TEAL = "087E8B"
LIGHT_BLUE = "EAF2F8"


def _shade_cell(cell, fill: str) -> None:
    properties = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    properties.append(shading)


def _set_repeat_table_header(row) -> None:
    properties = row._tr.get_or_add_trPr()
    repeat = OxmlElement("w:tblHeader")
    repeat.set(qn("w:val"), "true")
    properties.append(repeat)


def _style_document(document: Document) -> None:
    section = document.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    normal = document.styles["Normal"]
    normal.font.name = "Calibri"
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(48, 65, 83)
    normal.paragraph_format.space_after = Pt(6)

    for style_name, size in (("Title", 24), ("Heading 1", 17), ("Heading 2", 13)):
        style = document.styles[style_name]
        style.font.name = "Calibri"
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor.from_string(NAVY)

    header = section.header.paragraphs[0]
    header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    run = header.add_run("ЭЛЕКТРОННЫЙ УЧЕБНИК  ·  ИНФОРМАЦИОННЫЕ ТЕХНОЛОГИИ")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor.from_string(TEAL)

    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = footer.add_run("Учебный шаблон практической работы")
    run.font.size = Pt(8)
    run.font.color.rgb = RGBColor(110, 125, 142)


def _add_table(document: Document, headers: tuple[str, ...], rows: list[tuple[str, ...]]):
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = True
    for index, title in enumerate(headers):
        cell = table.rows[0].cells[index]
        cell.text = title
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        _shade_cell(cell, NAVY)
        for run in cell.paragraphs[0].runs:
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
    _set_repeat_table_header(table.rows[0])

    for row_number, values in enumerate(rows):
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = value
            cells[index].vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            if row_number % 2 == 0:
                _shade_cell(cells[index], LIGHT_BLUE)
    return table


def build_practice_template() -> bytes:
    document = Document()
    _style_document(document)

    eyebrow = document.add_paragraph()
    eyebrow.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = eyebrow.add_run("ПРАКТИЧЕСКАЯ РАБОТА  ·  ОБРАЗЕЦ")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor.from_string(TEAL)

    title = document.add_heading("Устройства компьютера", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    subtitle = document.add_paragraph("Пример оформления и выполнения задания")
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER

    document.add_heading("Цель работы", level=1)
    document.add_paragraph(
        "Научиться определять назначение устройств компьютера и относить их "
        "к устройствам ввода, вывода, хранения или обработки данных."
    )

    document.add_heading("Пример выполнения", level=1)
    document.add_paragraph(
        "Для каждой части компьютера укажем её основную функцию и коротко "
        "объясним, для чего она нужна."
    )
    _add_table(
        document,
        ("Устройство", "Функция", "Пояснение"),
        [
            ("Клавиатура", "Ввод", "С помощью клавиш пользователь вводит текст и команды."),
            ("Монитор", "Вывод", "Показывает пользователю текст, изображения и окна программ."),
        ],
    )

    document.add_heading("Как выполнять", level=1)
    for text in (
        "1. Прочитай задание на следующей странице.",
        "2. Заполни пустую таблицу своими словами.",
        "3. Сохрани документ и загрузи его в электронный учебник.",
    ):
        document.add_paragraph(text)

    document.add_page_break()

    eyebrow = document.add_paragraph()
    eyebrow.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = eyebrow.add_run("ТВОЯ РАБОТА  ·  ЗАПОЛНИ САМОСТОЯТЕЛЬНО")
    run.bold = True
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor.from_string(TEAL)

    title = document.add_heading("Определи назначение устройств", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    document.add_paragraph(
        "По образцу на первой странице заполни функцию каждого устройства "
        "и напиши, какую задачу оно выполняет. Используй функции: "
        "ввод, вывод, хранение или обработка."
    )
    _add_table(
        document,
        ("Устройство", "Функция", "Пояснение своими словами"),
        [
            ("Процессор", "\n\n", "\n\n"),
            ("Сканер", "\n\n", "\n\n"),
            ("Твердотельный накопитель (SSD)", "\n\n", "\n\n"),
            ("Принтер", "\n\n", "\n\n"),
        ],
    )

    document.add_heading("Перед отправкой", level=2)
    document.add_paragraph(
        "Проверь, что заполнены все строки, сохрани файл в формате .docx "
        "и загрузи его в раздел «Практические работы»."
    )

    output = BytesIO()
    document.save(output)
    return output.getvalue()
