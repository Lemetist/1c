from __future__ import annotations

import re
from copy import deepcopy
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


BASE_DIR = Path("/home/dola/1c")
TARGET_DOC = BASE_DIR / "Максимов М.М БСТ2501.docx"
SAMPLE_DOC = BASE_DIR / "ПР Бынеева БСТ2501.docx"
OUTPUT_DOC = BASE_DIR / "Максимов М.М БСТ2501_оформлено.docx"

NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W_NS = NS["w"]
W = f"{{{W_NS}}}"
XML_NS = "http://www.w3.org/XML/1998/namespace"
XML_SPACE = f"{{{XML_NS}}}space"


def parse_xml(raw: bytes) -> etree._Element:
    return etree.fromstring(raw)


def body_children(root: etree._Element) -> list[etree._Element]:
    body = root.find(".//w:body", namespaces=NS)
    if body is None:
        raise RuntimeError("word/document.xml does not contain w:body")
    return list(body)


def paragraph_alignment(paragraph: etree._Element) -> str:
    jc = paragraph.find("./w:pPr/w:jc", namespaces=NS)
    return jc.get(f"{W}val") if jc is not None else ""


def split_paragraph_lines(paragraph: etree._Element) -> list[str]:
    lines = [""]
    nodes = paragraph.xpath(".//w:r/*", namespaces=NS)
    if not nodes:
        return [""]
    for node in nodes:
        if node.tag == f"{W}t":
            lines[-1] += node.text or ""
        elif node.tag == f"{W}tab":
            lines[-1] += "\t"
        elif node.tag == f"{W}br":
            lines.append("")
    return [line.replace("\xa0", " ") for line in lines] or [""]


def extract_table_matrix(table: etree._Element) -> list[list[str]]:
    rows: list[list[str]] = []
    for row in table.xpath("./w:tr", namespaces=NS):
        cells: list[str] = []
        for cell in row.xpath("./w:tc", namespaces=NS):
            parts = [text.replace("\xa0", " ") for text in cell.xpath(".//w:t/text()", namespaces=NS)]
            cells.append("".join(parts).strip())
        rows.append(cells)
    return rows


def remove_children_except(node: etree._Element, allowed_tags: set[str]) -> None:
    for child in list(node):
        if child.tag not in allowed_tags:
            node.remove(child)


def first_run_template(paragraph: etree._Element) -> etree._Element | None:
    for run in paragraph.xpath("./w:r", namespaces=NS):
        return run
    return None


def ensure_ppr(paragraph: etree._Element) -> etree._Element:
    ppr = paragraph.find("./w:pPr", namespaces=NS)
    if ppr is None:
        ppr = etree.Element(f"{W}pPr")
        paragraph.insert(0, ppr)
    return ppr


def strip_numbering(ppr: etree._Element) -> None:
    for tag in ("numPr", "tabs"):
        for child in ppr.findall(f"./w:{tag}", namespaces=NS):
            ppr.remove(child)


def set_indent(ppr: etree._Element, left: int | None = None, first_line: int | None = 0) -> None:
    ind = ppr.find("./w:ind", namespaces=NS)
    if ind is None:
        if left is None and first_line is None:
            return
        ind = etree.SubElement(ppr, f"{W}ind")
    if left is None:
        ind.attrib.pop(f"{W}left", None)
    else:
        ind.set(f"{W}left", str(left))
    if first_line is None:
        ind.attrib.pop(f"{W}firstLine", None)
    else:
        ind.set(f"{W}firstLine", str(first_line))
    ind.attrib.pop(f"{W}hanging", None)
    ind.attrib.pop(f"{W}start", None)


def clone_paragraph(
    template: etree._Element,
    text: str | None,
    *,
    left: int | None = None,
    first_line: int | None = 0,
    strip_numpr: bool = True,
) -> etree._Element:
    paragraph = deepcopy(template)
    remove_children_except(paragraph, {f"{W}pPr"})
    ppr = ensure_ppr(paragraph)
    if strip_numpr:
        strip_numbering(ppr)
    if left is not None or first_line is not None:
        set_indent(ppr, left=left, first_line=first_line)
    if not text:
        return paragraph

    run_template = first_run_template(template)
    if run_template is None:
        run = etree.Element(f"{W}r")
    else:
        run = deepcopy(run_template)
        for child in list(run):
            run.remove(child)
    text_node = etree.SubElement(run, f"{W}t")
    if text[:1].isspace() or text[-1:].isspace():
        text_node.set(XML_SPACE, "preserve")
    text_node.text = text
    paragraph.append(run)
    return paragraph


def clone_table(sample_table: etree._Element, target_matrix: list[list[str]], paragraph_template: etree._Element) -> etree._Element:
    table = deepcopy(sample_table)
    sample_cells = table.xpath("./w:tr/w:tc", namespaces=NS)
    flat_target = [cell for row in target_matrix for cell in row]
    if len(sample_cells) != len(flat_target):
        raise RuntimeError("Sample and target tables have different cell counts")

    for cell, text in zip(sample_cells, flat_target):
        tc_pr = cell.find("./w:tcPr", namespaces=NS)
        for child in list(cell):
            if child is not tc_pr:
                cell.remove(child)
        inner_template = cell.find("./w:p", namespaces=NS) or paragraph_template
        cell.append(clone_paragraph(inner_template, text, strip_numpr=True))
    return table


def heading_kind(text: str) -> str | None:
    if text == "Цель работы.":
        return "goal"
    if text.startswith("Задание №1."):
        return "s1"
    if text.startswith("Задание №2."):
        return "s2"
    if text.startswith("Задание №3."):
        return "s3"
    if text.startswith("Задание №4."):
        return "s4"
    if text.startswith("Задание №5."):
        return "s5"
    if text.startswith("Доп. задание."):
        return "extra"
    if text == "Вывод":
        return "outro"
    return None


TASK_RE = re.compile(r"^Задача\s+\d+\.?$")
SUBNET_RE = re.compile(r"^(?:[А-Яа-яA-Za-z]\)\s*)?Подсеть\s+\d+:$")
IP_BINARY_RE = re.compile(r"^(?:\d{1,3}\.){3}\d{1,3}:\s+[01.]+$")
MASK_BINARY_RE = re.compile(r"^/\d{1,2}:\s+[01.]+$")
PURE_BINARY_RE = re.compile(r"^[01]{8}(?:\.[01]{8}){3}$")


def is_binaryish(text: str) -> bool:
    return bool(
        IP_BINARY_RE.match(text)
        or MASK_BINARY_RE.match(text)
        or PURE_BINARY_RE.match(text)
        or text.startswith("000000")
        or text.startswith("111111")
        or text.startswith("001")
        or text.startswith("010")
        or text.startswith("011")
        or text.startswith("100")
        or text.startswith("101")
        or text.startswith("110")
    )


def is_subnet_item(text: str) -> bool:
    plain = text.removeprefix("• ").removeprefix("- ").strip()
    return plain.startswith("Адрес подсети:") or plain.startswith("Блок адресов:")


def classify_body_line(text: str, section: str | None, block: str | None) -> tuple[str, int | None]:
    if section in {"goal", "outro"}:
        return "plain", None
    if section == "s1":
        return "plain", None
    if section == "s4":
        return "plain", None
    if section in {"s2", "s5"}:
        if block == "given":
            return "indent1", 708
        if block == "solution":
            return ("indent2", 1416) if is_binaryish(text) else ("indent1", 708)
        if block == "answer":
            return "indent1", 708
        return "plain", None
    if section == "s3":
        if block == "given":
            return "indent1", 708
        if block == "solution":
            return ("indent2", 1416) if is_binaryish(text) else ("indent1", 708)
        if block == "answer":
            if text.startswith("Адрес узла в двоичном виде:"):
                return "plain", None
            return "indent1", 708
        return "plain", None
    if section == "extra":
        if block == "given":
            return "plain", None
        if block == "solution":
            if is_subnet_item(text):
                return "indent1", 1068
            if SUBNET_RE.match(text):
                return "indent1", 708
            if is_binaryish(text):
                return "indent1", 708
            if (
                text.startswith("Представим")
                or text.startswith("Применим")
                or text.startswith("Адрес сети")
            ):
                return "indent1", 708
            return "plain", None
        if block == "answer":
            if is_subnet_item(text):
                return "indent1", 1068
            if SUBNET_RE.match(text):
                return "indent1", 708
            return "plain", None
        return "plain", None
    return "plain", None


def main() -> None:
    with ZipFile(SAMPLE_DOC) as sample_zip, ZipFile(TARGET_DOC) as target_zip:
        sample_doc = parse_xml(sample_zip.read("word/document.xml"))
        target_doc = parse_xml(target_zip.read("word/document.xml"))

        sample_paragraphs = sample_doc.xpath("//w:body/w:p", namespaces=NS)
        sample_tables = sample_doc.xpath("//w:body/w:tbl", namespaces=NS)

        templates = {
            "title_center": sample_paragraphs[0],
            "title_blank": sample_paragraphs[5],
            "author_right": sample_paragraphs[18],
            "author_right_plain": sample_paragraphs[20],
            "date_center": sample_paragraphs[29],
            "heading": sample_paragraphs[44],
            "plain": sample_paragraphs[45],
            "body_blank": sample_paragraphs[46],
            "task": sample_paragraphs[50],
            "label": sample_paragraphs[74],
            "indent1": sample_paragraphs[79],
            "indent2": sample_paragraphs[80],
            "answer_label": sample_paragraphs[86],
        }

        sample_body = sample_doc.find(".//w:body", namespaces=NS)
        if sample_body is None:
            raise RuntimeError("Sample document body not found")
        sect_pr = sample_body.find("./w:sectPr", namespaces=NS)
        if sect_pr is None:
            raise RuntimeError("Sample document section properties not found")
        sect_pr_copy = deepcopy(sect_pr)

        for child in list(sample_body):
            sample_body.remove(child)

        target_children = body_children(target_doc)
        title_phase = True
        title_line_index = 0
        current_section: str | None = None
        current_block: str | None = None
        table_index = 0

        for child in target_children:
            local_name = etree.QName(child).localname
            if local_name == "sectPr":
                continue

            if local_name == "tbl":
                target_matrix = extract_table_matrix(child)
                if table_index >= len(sample_tables):
                    raise RuntimeError("Target has more tables than sample")
                sample_body.append(clone_table(sample_tables[table_index], target_matrix, templates["plain"]))
                table_index += 1
                continue

            if local_name != "p":
                continue

            lines = split_paragraph_lines(child)
            for raw_line in lines:
                text = raw_line.strip()
                if not text:
                    blank_template = templates["title_blank"] if title_phase else templates["body_blank"]
                    sample_body.append(clone_paragraph(blank_template, None))
                    continue

                if title_phase and text != "Цель работы.":
                    title_line_index += 1
                    if title_line_index <= 8:
                        template = templates["title_center"]
                    elif title_line_index <= 10:
                        template = templates["author_right"]
                    elif title_line_index <= 12:
                        template = templates["author_right_plain"]
                    else:
                        template = templates["date_center"]
                    sample_body.append(clone_paragraph(template, text))
                    continue

                if title_phase and text == "Цель работы.":
                    title_phase = False

                heading = heading_kind(text)
                if heading is not None:
                    current_section = heading
                    current_block = None
                    sample_body.append(clone_paragraph(templates["heading"], text))
                    continue

                if TASK_RE.match(text):
                    current_block = None
                    sample_body.append(clone_paragraph(templates["task"], text))
                    continue

                if text in {"Дано:", "Решение:", "Ответ:"}:
                    current_block = {"Дано:": "given", "Решение:": "solution", "Ответ:": "answer"}[text]
                    template = templates["answer_label"] if text == "Ответ:" else templates["label"]
                    sample_body.append(clone_paragraph(template, text))
                    continue

                template_name, left = classify_body_line(text, current_section, current_block)
                template = templates[template_name]
                sample_body.append(clone_paragraph(template, text, left=left))

        sample_body.append(sect_pr_copy)
        new_document_xml = etree.tostring(
            sample_doc,
            xml_declaration=True,
            encoding="UTF-8",
            standalone="yes",
        )

        with ZipFile(SAMPLE_DOC) as sample_zip, ZipFile(TARGET_DOC) as target_zip, ZipFile(OUTPUT_DOC, "w", compression=ZIP_DEFLATED) as output_zip:
            target_overrides = {
                "docProps/core.xml",
                "docProps/app.xml",
            }
            for info in sample_zip.infolist():
                if info.filename == "word/document.xml":
                    output_zip.writestr(info, new_document_xml)
                elif info.filename in target_overrides and info.filename in target_zip.namelist():
                    output_zip.writestr(info, target_zip.read(info.filename))
                else:
                    output_zip.writestr(info, sample_zip.read(info.filename))

    print(f"Saved: {OUTPUT_DOC}")


if __name__ == "__main__":
    main()
