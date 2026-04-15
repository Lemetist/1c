from __future__ import annotations

import re
import shutil
from pathlib import Path

import officehelper
import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.text.ControlCharacter import PARAGRAPH_BREAK


SECTION_TITLES = {
    "Содержание",
    "Цель работы",
    "Задание",
    "Ход работы",
    "Заключение",
    "Ответы на контрольные вопросы",
}

NORMAL_WEIGHT = 100.0
BOLD_WEIGHT = 150.0


def make_property(name: str, value):
    prop = PropertyValue()
    prop.Name = name
    prop.Value = value
    return prop


def to_url(path: Path) -> str:
    return uno.systemPathToFileUrl(str(path.resolve()))


def line_spacing(height: int):
    spacing = uno.createUnoStruct("com.sun.star.style.LineSpacing")
    spacing.Mode = 0
    spacing.Height = height
    return spacing


def bootstrap_desktop():
    context = officehelper.bootstrap()
    service_manager = context.ServiceManager
    return service_manager.createInstanceWithContext("com.sun.star.frame.Desktop", context)


def load_document(desktop, path: Path):
    return desktop.loadComponentFromURL(
        to_url(path),
        "_blank",
        0,
        (make_property("Hidden", True),),
    )


def iter_paragraphs(document):
    enumeration = document.Text.createEnumeration()
    paragraphs = []
    while enumeration.hasMoreElements():
        paragraph = enumeration.nextElement()
        if hasattr(paragraph, "String"):
            paragraphs.append(paragraph)
    return paragraphs


def iter_text_portions(paragraph):
    try:
        enumeration = paragraph.createEnumeration()
    except Exception:
        return
    while enumeration.hasMoreElements():
        portion = enumeration.nextElement()
        if getattr(portion, "TextPortionType", "") == "Text":
            yield portion


def set_text_format(paragraph, *, size=14.0, bold=False, italic=False):
    weight = BOLD_WEIGHT if bold else NORMAL_WEIGHT
    paragraph.CharFontName = "Times New Roman"
    paragraph.CharFontNameAsian = "Times New Roman"
    paragraph.CharFontNameComplex = "Times New Roman"
    paragraph.CharHeight = size
    paragraph.CharHeightAsian = size
    paragraph.CharHeightComplex = size
    paragraph.CharWeight = weight
    paragraph.CharWeightAsian = weight
    paragraph.CharWeightComplex = weight
    paragraph.CharPosture = "ITALIC" if italic else "NONE"
    paragraph.CharPostureAsian = "ITALIC" if italic else "NONE"
    paragraph.CharPostureComplex = "ITALIC" if italic else "NONE"
    for portion in iter_text_portions(paragraph):
        portion.CharFontName = "Times New Roman"
        portion.CharFontNameAsian = "Times New Roman"
        portion.CharFontNameComplex = "Times New Roman"
        portion.CharHeight = size
        portion.CharHeightAsian = size
        portion.CharHeightComplex = size
        portion.CharWeight = weight
        portion.CharWeightAsian = weight
        portion.CharWeightComplex = weight
        portion.CharPosture = "ITALIC" if italic else "NONE"
        portion.CharPostureAsian = "ITALIC" if italic else "NONE"
        portion.CharPostureComplex = "ITALIC" if italic else "NONE"


def format_title_paragraph(paragraph):
    paragraph.ParaAdjust = 3
    paragraph.ParaFirstLineIndent = 0
    paragraph.ParaLeftMargin = 0
    paragraph.ParaRightMargin = 0
    paragraph.ParaTopMargin = 0
    paragraph.ParaBottomMargin = 282
    paragraph.ParaLineSpacing = line_spacing(150)
    set_text_format(paragraph)


def format_heading(paragraph, *, top_margin=635):
    paragraph.ParaStyleName = "Heading 1"
    paragraph.ParaAdjust = 0
    paragraph.ParaFirstLineIndent = 0
    paragraph.ParaLeftMargin = 0
    paragraph.ParaRightMargin = 0
    paragraph.ParaTopMargin = top_margin
    paragraph.ParaBottomMargin = 141
    paragraph.ParaLineSpacing = line_spacing(116)
    set_text_format(paragraph, bold=True)


def format_body(paragraph, *, indent=True):
    paragraph.ParaAdjust = 2
    paragraph.ParaFirstLineIndent = 1270 if indent else 0
    paragraph.ParaLeftMargin = 0
    paragraph.ParaRightMargin = 0
    paragraph.ParaTopMargin = 0
    paragraph.ParaBottomMargin = 0
    paragraph.ParaLineSpacing = line_spacing(150)
    set_text_format(paragraph)


def format_caption(paragraph):
    paragraph.ParaAdjust = 3
    paragraph.ParaFirstLineIndent = 0
    paragraph.ParaLeftMargin = 0
    paragraph.ParaRightMargin = 0
    paragraph.ParaTopMargin = 150
    paragraph.ParaBottomMargin = 150
    paragraph.ParaLineSpacing = line_spacing(116)
    set_text_format(paragraph, italic=True)


def find_index(paragraphs, text: str) -> int:
    for index, paragraph in enumerate(paragraphs):
        if paragraph.String.strip() == text:
            return index
    raise ValueError(f"Не найден абзац: {text}")


def delete_old_toc(document):
    paragraphs = iter_paragraphs(document)
    toc_index = find_index(paragraphs, "Содержание")
    body_index = find_index(paragraphs, "Цель работы")
    if body_index - toc_index <= 1:
        return
    cursor = document.Text.createTextCursorByRange(paragraphs[toc_index + 1].getStart())
    cursor.gotoRange(paragraphs[body_index].getStart(), True)
    cursor.setString("")


def insert_toc(document):
    paragraphs = iter_paragraphs(document)
    toc_heading_index = find_index(paragraphs, "Содержание")
    cursor = document.Text.createTextCursorByRange(paragraphs[toc_heading_index].getEnd())
    document.Text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
    index = document.createInstance("com.sun.star.text.ContentIndex")
    index.CreateFromOutline = True
    index.Level = 10
    index.Title = ""
    document.Text.insertTextContent(cursor, index, False)
    document.Text.insertControlCharacter(cursor, PARAGRAPH_BREAK, False)
    index.update()


def apply_template_styles(document, template_path: Path):
    options = (
        make_property("OverwriteStyles", True),
        make_property("LoadPageStyles", True),
        make_property("LoadFrameStyles", True),
        make_property("LoadTextStyles", True),
        make_property("LoadNumberingStyles", True),
    )
    document.StyleFamilies.loadStylesFromURL(to_url(template_path), options)


def rename_conclusion(paragraphs):
    for paragraph in paragraphs:
        if paragraph.String.strip() == "Вывод":
            paragraph.String = "Заключение"
            return


def format_document(document):
    delete_old_toc(document)
    paragraphs = iter_paragraphs(document)
    rename_conclusion(paragraphs)
    paragraphs = iter_paragraphs(document)
    title_limit = find_index(paragraphs, "Содержание")
    current_section = ""

    for index, paragraph in enumerate(paragraphs):
        text = paragraph.String.strip()
        if index < title_limit:
            format_title_paragraph(paragraph)
            continue
        if not text:
            continue
        if text in SECTION_TITLES:
            current_section = text
            if text == "Содержание":
                paragraph.ParaStyleName = "Standard"
                paragraph.OutlineLevel = 10
                format_heading(paragraph, top_margin=0)
                paragraph.ParaStyleName = "Standard"
            else:
                format_heading(paragraph, top_margin=635)
            continue
        if text.startswith("Рисунок "):
            format_caption(paragraph)
            continue
        is_numbered = bool(re.match(r"^\d+\.\s", text))
        indent = not is_numbered and current_section not in {"Задание", "Ответы на контрольные вопросы"}
        format_body(paragraph, indent=indent)

    insert_toc(document)
    document.updateLinks()
    document.refresh()
    for idx in range(document.DocumentIndexes.getCount()):
        document.DocumentIndexes.getByIndex(idx).update()


def store_document(document, path: Path, filter_name: str):
    document.storeToURL(
        to_url(path),
        (
            make_property("FilterName", filter_name),
            make_property("Overwrite", True),
        ),
    )


def main():
    base_dir = Path("/home/lemetist/1c")
    source_path = base_dir / "Лаба 2.odt"
    template_path = Path("/tmp/Лабораторная работа №3.odt")
    if not template_path.exists():
        template_path = base_dir / "Лаба 3.odt"

    output_odt = base_dir / "Лабораторная работа №2.odt"
    output_docx = base_dir / "Лабораторная работа №2.docx"

    shutil.copy(source_path, output_odt)
    desktop = bootstrap_desktop()
    document = load_document(desktop, output_odt)
    try:
        apply_template_styles(document, template_path)
        format_document(document)
        store_document(document, output_odt, "writer8")
        store_document(document, output_docx, "Office Open XML Text")
    finally:
        document.close(True)

    print(output_odt)
    print(output_docx)


if __name__ == "__main__":
    main()
