import copy
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
ET.register_namespace("w", W_NS)


CAPTIONS = {
    126: [
        "Рисунок 41 - Включение UFW и проверка статуса брандмауэра.",
    ],
    127: [
        "Рисунок 42 - Установка пакета net-tools и запуск команды ifconfig.",
    ],
    130: [
        "Рисунок 43 - Запуск службы Nginx через systemctl.",
        "Рисунок 44 - Проверка стандартной страницы Nginx в браузере.",
    ],
    131: [
        "Рисунок 45 - Создание символической ссылки на info.txt в домашнем каталоге.",
    ],
}


def paragraph_text(paragraph: ET.Element) -> str:
    return "".join(t.text or "" for t in paragraph.findall(".//w:t", NS)).strip()


def set_paragraph_text(paragraph: ET.Element, text: str) -> None:
    text_nodes = paragraph.findall(".//w:t", NS)
    if not text_nodes:
        run = paragraph.find("w:r", NS)
        if run is None:
            run = ET.SubElement(paragraph, f"{{{W_NS}}}r")
            ET.SubElement(run, f"{{{W_NS}}}rPr")
        text_node = ET.SubElement(run, f"{{{W_NS}}}t")
        text_node.text = text
        return
    text_nodes[0].text = text
    for node in text_nodes[1:]:
        node.text = ""


def main(docx_path: str) -> int:
    path = Path(docx_path)
    with zipfile.ZipFile(path, "r") as src:
        document_xml = src.read("word/document.xml")

    root = ET.fromstring(document_xml)
    body = root.find("w:body", NS)
    if body is None:
        raise RuntimeError("word/document.xml does not contain w:body")

    paragraphs = body.findall("w:p", NS)
    template = None
    for paragraph in paragraphs:
        if paragraph_text(paragraph).startswith("Рисунок 40 -"):
            template = paragraph
            break
    if template is None:
        raise RuntimeError("Caption template paragraph not found")

    body_children = list(body)
    for para_index in sorted(CAPTIONS.keys(), reverse=True):
        image_paragraph = paragraphs[para_index]
        insert_at = body_children.index(image_paragraph) + 1
        for caption in reversed(CAPTIONS[para_index]):
            new_paragraph = copy.deepcopy(template)
            set_paragraph_text(new_paragraph, caption)
            body.insert(insert_at, new_paragraph)

    updated_xml = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx", dir=path.parent) as tmp:
        tmp_path = Path(tmp.name)

    try:
        with zipfile.ZipFile(path, "r") as src, zipfile.ZipFile(
            tmp_path, "w", compression=zipfile.ZIP_DEFLATED
        ) as dst:
            for item in src.infolist():
                data = updated_xml if item.filename == "word/document.xml" else src.read(item.filename)
                dst.writestr(item, data)
        shutil.move(str(tmp_path), str(path))
    finally:
        if tmp_path.exists():
            tmp_path.unlink()

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1]))
