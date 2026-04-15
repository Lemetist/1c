from __future__ import annotations

import re
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape


ROOT = Path("/home/lemetist/1c")
TEMPLATE_PATH = ROOT / "Максимов М.М БСТ2501.docx"
OUTPUT_PATH = ROOT / "Максимов М.М БСТ2501 вариант 12.docx"

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_HEADER = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:o="urn:schemas-microsoft-com:office:office" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:w10="urn:schemas-microsoft-com:office:word" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
    'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" '
    'mc:Ignorable="w14 wp14 w15"><w:body>'
)
XML_FOOTER = "</w:body></w:document>"


VARIANT_CASES = [
    {
        "task": 1,
        "ip": "62.187.216.242",
        "prefix": 26,
        "class": "A",
        "first_octet_bits": "00111110",
        "mask_bits": "11111111.11111111.11111111.11000000",
        "net_bits": "00111110.10111011.11011000.11000000",
        "net": "62.187.216.192",
        "inv_bits": "00000000.00000000.00000000.00111111",
        "host_bits": "00000000.00000000.00000000.00110010",
        "host": "0.0.0.50",
        "host_num": 50,
        "broadcast_bits": "00111110.10111011.11011000.11111111",
        "broadcast": "62.187.216.255",
        "subnets": [
            ("62.187.216.192/29", "62.187.216.192 - 62.187.216.199"),
            ("62.187.216.200/29", "62.187.216.200 - 62.187.216.207"),
            ("62.187.216.208/29", "62.187.216.208 - 62.187.216.215"),
            ("62.187.216.216/29", "62.187.216.216 - 62.187.216.223"),
            ("62.187.216.224/29", "62.187.216.224 - 62.187.216.231"),
            ("62.187.216.232/29", "62.187.216.232 - 62.187.216.239"),
            ("62.187.216.240/29", "62.187.216.240 - 62.187.216.247"),
            ("62.187.216.248/29", "62.187.216.248 - 62.187.216.255"),
        ],
    },
    {
        "task": 2,
        "ip": "144.17.221.246",
        "prefix": 17,
        "class": "B",
        "first_octet_bits": "10010000",
        "mask_bits": "11111111.11111111.10000000.00000000",
        "net_bits": "10010000.00010001.10000000.00000000",
        "net": "144.17.128.0",
        "inv_bits": "00000000.00000000.01111111.11111111",
        "host_bits": "00000000.00000000.01011101.11110110",
        "host": "0.0.93.246",
        "host_num": 24054,
        "broadcast_bits": "10010000.00010001.11111111.11111111",
        "broadcast": "144.17.255.255",
        "subnets": [
            ("144.17.128.0/20", "144.17.128.0 - 144.17.143.255"),
            ("144.17.144.0/20", "144.17.144.0 - 144.17.159.255"),
            ("144.17.160.0/20", "144.17.160.0 - 144.17.175.255"),
            ("144.17.176.0/20", "144.17.176.0 - 144.17.191.255"),
            ("144.17.192.0/20", "144.17.192.0 - 144.17.207.255"),
            ("144.17.208.0/20", "144.17.208.0 - 144.17.223.255"),
            ("144.17.224.0/20", "144.17.224.0 - 144.17.239.255"),
            ("144.17.240.0/20", "144.17.240.0 - 144.17.255.255"),
        ],
    },
    {
        "task": 3,
        "ip": "196.147.69.227",
        "prefix": 24,
        "class": "C",
        "first_octet_bits": "11000100",
        "mask_bits": "11111111.11111111.11111111.00000000",
        "net_bits": "11000100.10010011.01000101.00000000",
        "net": "196.147.69.0",
        "inv_bits": "00000000.00000000.00000000.11111111",
        "host_bits": "00000000.00000000.00000000.11100011",
        "host": "0.0.0.227",
        "host_num": 227,
        "broadcast_bits": "11000100.10010011.01000101.11111111",
        "broadcast": "196.147.69.255",
        "subnets": [
            ("196.147.69.0/27", "196.147.69.0 - 196.147.69.31"),
            ("196.147.69.32/27", "196.147.69.32 - 196.147.69.63"),
            ("196.147.69.64/27", "196.147.69.64 - 196.147.69.95"),
            ("196.147.69.96/27", "196.147.69.96 - 196.147.69.127"),
            ("196.147.69.128/27", "196.147.69.128 - 196.147.69.159"),
            ("196.147.69.160/27", "196.147.69.160 - 196.147.69.191"),
            ("196.147.69.192/27", "196.147.69.192 - 196.147.69.223"),
            ("196.147.69.224/27", "196.147.69.224 - 196.147.69.255"),
        ],
    },
    {
        "task": 4,
        "ip": "227.147.196.69",
        "prefix": 21,
        "class": "D",
        "first_octet_bits": "11100011",
        "mask_bits": "11111111.11111111.11111000.00000000",
        "net_bits": "11100011.10010011.11000000.00000000",
        "net": "227.147.192.0",
        "inv_bits": "00000000.00000000.00000111.11111111",
        "host_bits": "00000000.00000000.00000100.01000101",
        "host": "0.0.4.69",
        "host_num": 1093,
        "broadcast_bits": "11100011.10010011.11000111.11111111",
        "broadcast": "227.147.199.255",
        "subnets": [
            ("227.147.192.0/24", "227.147.192.0 - 227.147.192.255"),
            ("227.147.193.0/24", "227.147.193.0 - 227.147.193.255"),
            ("227.147.194.0/24", "227.147.194.0 - 227.147.194.255"),
            ("227.147.195.0/24", "227.147.195.0 - 227.147.195.255"),
            ("227.147.196.0/24", "227.147.196.0 - 227.147.196.255"),
            ("227.147.197.0/24", "227.147.197.0 - 227.147.197.255"),
            ("227.147.198.0/24", "227.147.198.0 - 227.147.198.255"),
            ("227.147.199.0/24", "227.147.199.0 - 227.147.199.255"),
        ],
    },
]


def extract_sect_pr(template_doc_xml: str) -> str:
    match = re.search(r"(<w:sectPr[\s\S]*</w:sectPr>)\s*</w:body></w:document>", template_doc_xml)
    if not match:
        raise RuntimeError("Не удалось извлечь w:sectPr из шаблона")
    return match.group(1)


def run_xml(text: str, *, bold: bool = False, italic: bool = False, size: int = 28) -> str:
    if text == "":
        return f"<w:r><w:rPr><w:sz w:val=\"{size}\"/><w:szCs w:val=\"{size}\"/></w:rPr></w:r>"

    parts = [f"<w:sz w:val=\"{size}\"/>", f"<w:szCs w:val=\"{size}\"/>"]
    if bold:
        parts.append("<w:b/><w:bCs/>")
    if italic:
        parts.append("<w:i/><w:iCs/>")
    return (
        "<w:r>"
        f"<w:rPr>{''.join(parts)}</w:rPr>"
        f"<w:t xml:space=\"preserve\">{escape(text)}</w:t>"
        "</w:r>"
    )


def paragraph_xml(
    text: str = "",
    *,
    style: str = "Normal",
    align: str | None = None,
    bold: bool = False,
    italic: bool = False,
    size: int = 28,
    before: int | None = None,
    after: int | None = None,
    line: int | None = None,
    page_break_before: bool = False,
) -> str:
    ppr = [f"<w:pStyle w:val=\"{style}\"/>"]
    if page_break_before:
        ppr.append("<w:pageBreakBefore/>")
    if before is not None or after is not None or line is not None:
        attrs = []
        if before is not None:
            attrs.append(f'w:before="{before}"')
        if after is not None:
            attrs.append(f'w:after="{after}"')
        if line is not None:
            attrs.append(f'w:line="{line}" w:lineRule="auto"')
        ppr.append(f"<w:spacing {' '.join(attrs)}/>")
    if align is not None:
        ppr.append(f"<w:jc w:val=\"{align}\"/>")
    return f"<w:p><w:pPr>{''.join(ppr)}</w:pPr>{run_xml(text, bold=bold, italic=italic, size=size)}</w:p>"


def page_break_xml() -> str:
    return "<w:p><w:r><w:br w:type=\"page\"/></w:r></w:p>"


def build_title_page() -> list[str]:
    items = [
        paragraph_xml("Ордена Трудового Красного Знамени", align="center"),
        paragraph_xml("Федеральное Государственное бюджетное образовательное учреждение", align="center"),
        paragraph_xml("высшего образования", align="center"),
        paragraph_xml("«Московский Технический Университет Связи и Информатики»", align="center"),
        paragraph_xml("Кафедра «Сетевые информационные технологии и сервисы»", align="center"),
        paragraph_xml("", align="center"),
        paragraph_xml("", align="center"),
        paragraph_xml("", align="center"),
        paragraph_xml("Практическая работа", align="center"),
        paragraph_xml("по дисциплине «Введение в ИТ» на тему:", align="center"),
        paragraph_xml("«IPv4»", align="center"),
        paragraph_xml("", align="center"),
        paragraph_xml("", align="center"),
        paragraph_xml("", align="center"),
        paragraph_xml("", align="center"),
        paragraph_xml("Выполнил: студент группы БСТ2501", style="NormalWeb", align="end"),
        paragraph_xml("Максимов М.М", style="NormalWeb", align="end"),
        paragraph_xml("Проверила: старший преподаватель", align="end"),
        paragraph_xml("Комкова М.Г.", align="end"),
        paragraph_xml("", align="center"),
        paragraph_xml("", align="center"),
        paragraph_xml("", align="center"),
        paragraph_xml("Москва, 2026", align="center"),
    ]
    return items


def build_task_1() -> list[str]:
    items = [
        paragraph_xml(
            "Задание №1. Определение принадлежности IP-адреса конкретному классу сети.",
            style="Heading2",
        )
    ]
    for case in VARIANT_CASES:
        items.extend(
            [
                paragraph_xml(f"Задача {case['task']}", bold=True),
                paragraph_xml(f"Дано: IP-адрес: {case['ip']}."),
                paragraph_xml(
                    "Решение: первый октет адреса в двоичном виде равен "
                    f"{case['first_octet_bits']}. Следовательно, адрес относится к классу {case['class']}."
                ),
                paragraph_xml(f"Ответ: {case['ip']} – класс {case['class']}."),
            ]
        )
    return items


def build_task_2() -> list[str]:
    items = [
        paragraph_xml(
            "Задание №2. Вычисление адреса сети по заданному IP-адресу узла и маске подсети.",
            style="Heading2",
        )
    ]
    for case in VARIANT_CASES:
        items.extend(
            [
                paragraph_xml(f"Задача {case['task']}", bold=True),
                paragraph_xml(f"Дано: IP-адрес: {case['ip']}."),
                paragraph_xml(f"Маска подсети: /{case['prefix']}."),
                paragraph_xml(f"Запись адреса: {case['ip']}/{case['prefix']}."),
                paragraph_xml("Решение:", bold=True),
                paragraph_xml(f"{case['ip']}: {ip_to_bits(case['ip'])}"),
                paragraph_xml(f"/{case['prefix']}: {case['mask_bits']}"),
                paragraph_xml("Применим операцию побитового «И»:"),
                paragraph_xml(ip_to_bits(case["ip"])),
                paragraph_xml(case["mask_bits"]),
                paragraph_xml(case["net_bits"]),
                paragraph_xml(f"Ответ: адрес сети в десятичном виде: {case['net']}."),
            ]
        )
    return items


def build_task_3() -> list[str]:
    items = [
        paragraph_xml(
            "Задание №3. Вычисление адреса узла в сети по заданному IP-адресу и маске подсети.",
            style="Heading2",
        )
    ]
    for case in VARIANT_CASES:
        items.extend(
            [
                paragraph_xml(f"Задача {case['task']}", bold=True),
                paragraph_xml(f"Дано: IP-адрес: {case['ip']}."),
                paragraph_xml(f"Маска подсети: /{case['prefix']}."),
                paragraph_xml(f"Запись адреса: {case['ip']}/{case['prefix']}."),
                paragraph_xml("Решение:", bold=True),
                paragraph_xml(f"Обратная маска: {case['inv_bits']}"),
                paragraph_xml(f"{case['ip']}: {ip_to_bits(case['ip'])}"),
                paragraph_xml(case["inv_bits"]),
                paragraph_xml(case["host_bits"]),
                paragraph_xml(f"Ответ: адрес узла в десятичном виде: {case['host']}."),
            ]
        )
    return items


def build_task_4() -> list[str]:
    items = [
        paragraph_xml(
            "Задание №4. Вычисление порядкового номера узла в сети.",
            style="Heading2",
        ),
        paragraph_xml(
            "Порядковый номер узла определяется переводом адреса узла из двоичного вида в десятичное число."
        ),
    ]
    for case in VARIANT_CASES:
        items.extend(
            [
                paragraph_xml(f"Задача {case['task']}", bold=True),
                paragraph_xml(f"Адрес узла в двоичном виде: {case['host_bits']}"),
                paragraph_xml(f"Адрес узла в десятичном виде: {case['host_num']}"),
            ]
        )
    return items


def build_task_5() -> list[str]:
    items = [
        paragraph_xml(
            "Задание №5. Определение широковещательного IP-адреса сети.",
            style="Heading2",
        )
    ]
    for case in VARIANT_CASES:
        items.extend(
            [
                paragraph_xml(f"Задача {case['task']}", bold=True),
                paragraph_xml(f"Дано: IP-адрес: {case['ip']}."),
                paragraph_xml(f"Маска подсети: /{case['prefix']}."),
                paragraph_xml(f"Запись адреса: {case['ip']}/{case['prefix']}."),
                paragraph_xml("Решение:", bold=True),
                paragraph_xml(f"{case['ip']}: {ip_to_bits(case['ip'])}"),
                paragraph_xml(f"Обратная маска: {case['inv_bits']}"),
                paragraph_xml(case["broadcast_bits"]),
                paragraph_xml(
                    f"Ответ: широковещательный IP-адрес в десятичном виде: {case['broadcast']}."
                ),
            ]
        )
    return items


def build_extra_task() -> list[str]:
    items = [
        paragraph_xml(
            "Доп. задание. Распределение адресного пространства на основе маски подсети. Деление блоков IP-адресов на части.",
            style="Heading2",
        ),
        paragraph_xml(
            "Каждая сеть из задания №2 разделена на 8 равных подсетей путем увеличения префикса на 3 бита."
        ),
    ]
    for case in VARIANT_CASES:
        items.extend(
            [
                paragraph_xml(f"Задача {case['task']}", bold=True),
                paragraph_xml(f"Исходная сеть: {case['net']}/{case['prefix']}."),
            ]
        )
        for index, (subnet, block) in enumerate(case["subnets"], start=1):
            items.append(paragraph_xml(f"Подсеть {index}:", bold=True))
            items.append(paragraph_xml(f"Адрес подсети: {subnet};"))
            items.append(paragraph_xml(f"Блок адресов: {block}"))
    return items


def build_document_xml(sect_pr: str) -> str:
    body = []
    body.extend(build_title_page())
    body.append(page_break_xml())

    body.append(paragraph_xml("Оглавление", style="Heading1"))
    body.append(paragraph_xml("Цель работы"))
    body.append(paragraph_xml("Задание"))
    body.append(paragraph_xml("Задание №1. Определение принадлежности IP-адреса конкретному классу сети"))
    body.append(paragraph_xml("Задание №2. Вычисление адреса сети по заданному IP-адресу узла и маске подсети"))
    body.append(paragraph_xml("Задание №3. Вычисление адреса узла в сети по заданному IP-адресу и маске подсети"))
    body.append(paragraph_xml("Задание №4. Вычисление порядкового номера узла в сети"))
    body.append(paragraph_xml("Задание №5. Определение широковещательного IP-адреса сети"))
    body.append(paragraph_xml("Доп. задание. Деление адресного пространства на 8 подсетей"))
    body.append(paragraph_xml("Вывод"))

    body.append(paragraph_xml("Цель работы", style="Heading1"))
    body.append(
        paragraph_xml(
            "Закрепление теоретических знаний и получение практических навыков по распределению "
            "адресного пространства протокола IPv4, определению классов IP-адресов, вычислению "
            "адресов сетей и узлов, нахождению широковещательных адресов, порядковых номеров узлов "
            "и разбиению сетей на подсети."
        )
    )

    body.append(paragraph_xml("Задание", style="Heading1"))
    body.append(
        paragraph_xml(
            "В работе используется вариант 12. Для него необходимо решить задачи 1–5 по данным из "
            "таблиц методических указаний и выполнить дополнительное разбиение каждой полученной сети "
            "на 8 равных подсетей."
        )
    )

    body.extend(build_task_1())
    body.extend(build_task_2())
    body.extend(build_task_3())
    body.extend(build_task_4())
    body.extend(build_task_5())
    body.extend(build_extra_task())

    body.append(paragraph_xml("Вывод", style="Heading2"))
    body.append(
        paragraph_xml(
            "В ходе выполнения практической работы были определены классы IP-адресов для варианта 12, "
            "вычислены адреса сетей, адреса узлов, порядковые номера узлов и широковещательные адреса. "
            "Дополнительно каждая исходная сеть была разделена на 8 равных подсетей, что позволило "
            "закрепить навыки работы с масками подсети и принципами распределения адресного пространства IPv4."
        )
    )

    return XML_HEADER + "".join(body) + sect_pr + XML_FOOTER


def ip_to_bits(ip: str) -> str:
    return ".".join(f"{int(part):08b}" for part in ip.split("."))


def build_core_xml(template_core_xml: str) -> str:
    result = re.sub(
        r"<dc:title>.*?</dc:title>",
        "<dc:title>Практическая работа IPv4, вариант 12</dc:title>",
        template_core_xml,
    )
    result = re.sub(
        r"<dc:subject>.*?</dc:subject>",
        "<dc:subject>IPv4 variant 12</dc:subject>",
        result,
    )
    return result


def main() -> None:
    with zipfile.ZipFile(TEMPLATE_PATH, "r") as src:
        template_document_xml = src.read("word/document.xml").decode("utf-8")
        template_core_xml = src.read("docProps/core.xml").decode("utf-8")
        sect_pr = extract_sect_pr(template_document_xml)
        document_xml = build_document_xml(sect_pr)
        core_xml = build_core_xml(template_core_xml)

        with zipfile.ZipFile(OUTPUT_PATH, "w", compression=zipfile.ZIP_DEFLATED) as dst:
            for info in src.infolist():
                data = src.read(info.filename)
                if info.filename == "word/document.xml":
                    data = document_xml.encode("utf-8")
                elif info.filename == "docProps/core.xml":
                    data = core_xml.encode("utf-8")
                dst.writestr(info, data)


if __name__ == "__main__":
    main()
