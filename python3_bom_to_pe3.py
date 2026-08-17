#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор Перечня элементов (ПЭ3) из BOM Cadence Allegro или PADS
с использованием готового ODT-шаблона с рамкой и штампом ГОСТ.
Процедурный стиль, без классов.

Запуск:
    python3 bom_to_pe3.py input_bom.txt -t "Перечень элементов.odt" -p params.json
"""

import argparse
import json
import os
import re
import subprocess
import sys
import zipfile
from xml.sax.saxutils import escape as xml_escape

# --------------------------------------------------------------------------
# Настройки по умолчанию
# --------------------------------------------------------------------------

CONFIG = {
    "document_designation": "АБВГ.XXXXXX.XXX",
    "document_type": "ПЭ3",
    "version": "01",
    "product_name": "",
    "literature": "",
    "developer": "",
    "checker": "",
    "standard_control": "",
    "approver": "",
    "inventory_number": "",
    "first_application": "",
    "min_compression": 0.9,
    "rows_on_first_page": 29,
}

# Соответствие текстовых маркеров в шаблоне ключам параметров.
# Маркеры должны быть заранее вписаны как обычный текст в соответствующие
# ячейки styles.xml или content.xml. Децимальный номер обрабатывается
# отдельно, см. replace_decimal_number_blocks, поэтому в этом словаре его нет.
PLACEHOLDER_MAP = {
    "{Наименование}": "product_name",
    "{Л}": "literature",
    "{Разработал}": "developer",
    "{Проверил}": "checker",
    "{НКонтроль}": "standard_control",
    "{Утвердил}": "approver",
    "{Инвномер}": "inventory_number",
    "{Первпримен}": "first_application",
}

EXCLUDE_EMPTY_TYPE = True
REPLACE_DNP_LABEL = True

DNP_PATTERN = re.compile(r"\bDNP\b", re.IGNORECASE)
DNP_LABEL = "не устанавливается"

CATEGORY_HEADERS = {
    "C": "Конденсаторы",
    "D": "Микросхемы",
    "F": "Предохранители",
    "G": "Генераторы",
    "L": "Фильтр ферритовый",
    "R": "Резисторы",
    "S": "Переключатели",
    "V": "Полупроводниковые приборы",
    "X": "Соединители",
}

# Группы алиасов префиксов, которые нужно считать одной категорией.
PREFIX_ALIAS_GROUPS = {
    "R": {"R", "RN", "RZ"},
    "X": {"X", "XP", "XS", "XW"},
}

TEMPLATE_TABLE_NAME = "Перечень"

COLUMN_CHAR_LIMITS_AT_100 = {
    "A": 7,
    "B": 58,
    "D": 24,
}

BASE_STYLE_FOR_COLUMN = {
    "A": "P3",
    "B": "P5",
    "D": "P1",
}

ROW_STYLE_NAME = "СтрокаПеречня"

ROW_STYLE_XML = (
    '<style:style style:name="%s" style:family="table-row">'
    '<style:table-row-properties style:min-row-height="0.801cm"/>'
    "</style:style>"
) % ROW_STYLE_NAME

# Стиль абзаца для строк "Удостоверен ..." и "Имя файла ..." перед штампом,
# выравнивание вправо вместо стандартного центрирования заголовков категорий,
# без курсива.
TRAILER_TEXT_STYLE_NAME = "СтрокаТрейлер"

TRAILER_TEXT_STYLE_XML = (
    '<style:style style:name="%s" style:family="paragraph" style:parent-style-name="Table_20_Contents">'
    '<style:paragraph-properties fo:text-align="start" style:justify-single-word="false"/>'
    '<style:text-properties style:font-name="GOST type A1" fo:font-size="14pt"/>'
    "</style:style>"
) % TRAILER_TEXT_STYLE_NAME

ITALIC_ATTR_PATTERN = re.compile(
    r'(fo:font-style|style:font-style-asian|style:font-style-complex)="italic"'
)


def strip_italic(text):
    return ITALIC_ATTR_PATTERN.sub(lambda m: '%s="normal"' % m.group(1), text)


# Децимальный номер в шаблоне устроен через закладку ODF, а не как обычный
# текстовый маркер, и при копировании рамки между страницами LibreOffice
# может переименовывать закладки и дробить текст на несколько span в
# произвольных местах. Поэтому вместо жёстко заданной разметки используется
# общий механизм: находится любой блок от { до }, с него снимаются все теги,
# и если очищенный текст совпадает с "Децимальный номер", блок целиком
# заменяется на значение.
CURLY_BLOCK_PATTERN = re.compile(r"\{.*?\}", re.DOTALL)
TAG_PATTERN = re.compile(r"<[^>]+>")


def strip_tags(fragment):
    return TAG_PATTERN.sub("", fragment)


def replace_decimal_number_blocks(text, value):
    escaped_value = xml_escape(value)

    def repl(match):
        fragment = match.group(0)
        visible = strip_tags(fragment[1:-1]).strip()
        if visible == "Децимальный номер":
            return escaped_value
        return fragment

    return CURLY_BLOCK_PATTERN.sub(repl, text)


def apply_placeholders(text, display_values):
    decimal_value = display_values.get("document_designation", "")
    text = replace_decimal_number_blocks(text, decimal_value)

    for placeholder, key in PLACEHOLDER_MAP.items():
        value = display_values.get(key, "")
        text = text.replace(placeholder, xml_escape(value))

    return text


# --------------------------------------------------------------------------
# Описание форматов BOM: расположение полей и порядок их объединения
# --------------------------------------------------------------------------

FORMATS = {
    "cadence": {
        "header_map": {
            "reference": 0,
            "type": 1,
            "partnumber": 2,
            "tke": 3,
            "voltage": 4,
            "value_gost": 5,
            "tolerance": 6,
            "tu": 7,
            "manufacturer1": 8,
            "manufacturer": 9,
        },
        "merge_order": ["type", "partnumber", "tke", "voltage", "value_gost", "tolerance", "tu"],
        "field_count": 10,
        "skip_lines": 11,
    },
    "pads": {
        "header_map": {
            "reference": 0,
            "type": 1,
            "partnumber": 2,
            "tke": 3,
            "voltage": 4,
            "value_gost": 5,
            "tolerance": 6,
            "package": 7,
            "manufacturer1": 8,
        },
        "merge_order": ["partnumber", "tke", "voltage", "value_gost", "tolerance", "package"],
        "field_count": 9,
    },
}


# --------------------------------------------------------------------------
# Чтение файла BOM и определение формата
# --------------------------------------------------------------------------

def decode_any_text(raw_bytes):
    for enc in ("utf-8-sig", "utf-8", "cp1251"):
        try:
            return raw_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("cp1251", errors="replace")


def load_raw_lines(path):
    with open(path, "rb") as f:
        raw = f.read()
    text = decode_any_text(raw)
    return text.splitlines()


def detect_bom_format(lines):
    for line in lines[:5]:
        if "Bill Of Materials" in line:
            return "pads"
    return "cadence"


# --------------------------------------------------------------------------
# Разбор формата Cadence Allegro (табуляция)
# --------------------------------------------------------------------------

def is_separator_line_cadence(line):
    stripped = line.strip()
    if not stripped:
        return False
    return bool(re.fullmatch(r"[-=_]{5,}", stripped))


def strip_service_lines_cadence(lines, skip_count):
    working = lines[skip_count:]
    return [ln for ln in working if not is_separator_line_cadence(ln)]


def split_row_cadence(line, field_count):
    fields = line.rstrip("\r\n").split("\t")
    fields = [f.strip() for f in fields]
    if len(fields) < field_count:
        fields.extend([""] * (field_count - len(fields)))
    return fields[:field_count]


def parse_rows_cadence(lines, fmt):
    data_lines = strip_service_lines_cadence(lines, fmt["skip_lines"])
    return [split_row_cadence(ln, fmt["field_count"]) for ln in data_lines if ln.strip()]


# --------------------------------------------------------------------------
# Разбор формата PADS (таблица с "|")
# --------------------------------------------------------------------------

def is_pads_border_line(line):
    stripped = line.strip()
    if not stripped:
        return False
    return bool(re.fullmatch(r"[-+|]+", stripped))


def strip_service_lines_pads(lines):
    working = []
    for ln in lines:
        stripped = ln.strip()
        if not stripped:
            continue
        if is_pads_border_line(stripped):
            continue
        if "Bill Of Materials" in stripped:
            continue
        if stripped.startswith("|Reference|"):
            continue
        working.append(ln)
    return working


def split_row_pads(line, field_count):
    stripped = line.strip()
    if stripped.startswith("|"):
        stripped = stripped[1:]
    if stripped.endswith("|"):
        stripped = stripped[:-1]
    fields = [f.strip() for f in stripped.split("|")]
    if len(fields) < field_count:
        fields.extend([""] * (field_count - len(fields)))
    return fields[:field_count]


def parse_rows_pads(lines, fmt):
    data_lines = strip_service_lines_pads(lines)
    return [split_row_pads(ln, fmt["field_count"]) for ln in data_lines if ln.strip()]


PARSERS = {
    "cadence": parse_rows_cadence,
    "pads": parse_rows_pads,
}


# --------------------------------------------------------------------------
# Доступ к полям с учётом карты конкретного формата
# --------------------------------------------------------------------------

def get_field(fields, key, header_map):
    idx = header_map.get(key)
    if idx is None:
        return ""
    return fields[idx]


# --------------------------------------------------------------------------
# Объединение колонок в единое наименование
# --------------------------------------------------------------------------

def merge_type_fields(fields, header_map, merge_order):
    parts = []
    partnumber_value = get_field(fields, "partnumber", header_map)
    is_generic_partnumber = partnumber_value.upper().startswith("SMD")
    for key in merge_order:
        value = get_field(fields, key, header_map)
        if not value:
            continue
        if key == "package":
            if not is_generic_partnumber:
                continue
            if value.lower() in partnumber_value.lower():
                continue
        parts.append(value)
    return " ".join(parts)


def apply_dnp_label(merged_type):
    if REPLACE_DNP_LABEL and DNP_PATTERN.search(merged_type):
        return DNP_LABEL
    return merged_type


def get_manufacturer_note(fields, header_map):
    manufacturer1 = get_field(fields, "manufacturer1", header_map)
    if manufacturer1:
        return manufacturer1
    manufacturer = get_field(fields, "manufacturer", header_map)
    if manufacturer:
        return manufacturer
    return ""


# --------------------------------------------------------------------------
# Работа с позиционными обозначениями
# --------------------------------------------------------------------------

def split_refdes(field):
    parts = re.split(r"[,\s;]+", field.strip())
    return [p for p in parts if p]


def get_letter_prefix(refdes):
    match = re.match(r"^([A-ZА-Я]+)", refdes.upper())
    return match.group(1) if match else ""


def get_number_suffix(refdes):
    match = re.search(r"(\d+)$", refdes)
    return int(match.group(1)) if match else None


# --------------------------------------------------------------------------
# Построение плоского списка позиций и его сортировка
# --------------------------------------------------------------------------

def build_items(rows, header_map, merge_order):
    items = []
    for fields in rows:
        refdes_field = get_field(fields, "reference", header_map)
        if not refdes_field:
            continue
        refs = split_refdes(refdes_field)
        merged_type_raw = merge_type_fields(fields, header_map, merge_order)
        merged_type = apply_dnp_label(merged_type_raw)
        if EXCLUDE_EMPTY_TYPE and not merged_type:
            continue
        note = get_manufacturer_note(fields, header_map)
        for ref in refs:
            items.append({
                "ref": ref,
                "prefix": get_letter_prefix(ref),
                "number": get_number_suffix(ref),
                "type": merged_type,
                "note": note,
            })
    return items


def sort_items(items):
    return sorted(
        items,
        key=lambda it: (it["prefix"], it["number"] if it["number"] is not None else 0),
    )


# --------------------------------------------------------------------------
# Объединение только реально соседних позиций
# --------------------------------------------------------------------------

def merge_consecutive_items(items):
    groups = []
    current = None
    for item in items:
        same_group = (
            current is not None
            and current["prefix"] == item["prefix"]
            and current["type"] == item["type"]
            and current["note"] == item["note"]
        )
        if same_group:
            current["refs"].append(item["ref"])
        else:
            current = {
                "prefix": item["prefix"],
                "type": item["type"],
                "note": item["note"],
                "refs": [item["ref"]],
            }
            groups.append(current)
    return groups


# --------------------------------------------------------------------------
# Разрыв групп по числовым пропускам в обозначениях
# --------------------------------------------------------------------------

def split_group_by_contiguity(group):
    numbers = []
    text_only_refs = []
    for ref in group["refs"]:
        number = get_number_suffix(ref)
        if number is None:
            text_only_refs.append(ref)
        else:
            numbers.append(number)

    numbers = sorted(set(numbers))
    prefix = group["prefix"]
    runs = []
    if numbers:
        start = prev = numbers[0]
        for num in numbers[1:]:
            if num == prev + 1:
                prev = num
                continue
            runs.append(list(range(start, prev + 1)))
            start = prev = num
        runs.append(list(range(start, prev + 1)))

    sub_groups = []
    for run in runs:
        refs = ["%s%d" % (prefix, n) for n in run]
        sub_groups.append({
            "prefix": prefix,
            "type": group["type"],
            "note": group["note"],
            "refs": refs,
        })

    if text_only_refs:
        sub_groups.append({
            "prefix": prefix,
            "type": group["type"],
            "note": group["note"],
            "refs": text_only_refs,
        })

    return sub_groups


def split_groups_by_contiguity(groups):
    result = []
    for group in groups:
        result.extend(split_group_by_contiguity(group))
    return result


def find_duplicate_refs(groups):
    seen = {}
    duplicates = set()
    for index, group in enumerate(groups):
        for ref in group["refs"]:
            if ref in seen and seen[ref] != index:
                duplicates.add(ref)
            seen[ref] = index
    return sorted(duplicates)


# --------------------------------------------------------------------------
# Заголовки категорий и пустые строки между разными обозначениями
# --------------------------------------------------------------------------

def get_category_key(prefix):
    for canonical, aliases in PREFIX_ALIAS_GROUPS.items():
        if prefix in aliases:
            return canonical
    return prefix


def insert_category_headers(groups):
    entries = []
    last_category = None
    for group in groups:
        category = get_category_key(group["prefix"])
        if category != last_category:
            if last_category is not None:
                entries.append({"is_blank": True})
                entries.append({"is_blank": True})
            if category in CATEGORY_HEADERS:
                entries.append({"is_header": True, "text": CATEGORY_HEADERS[category]})
                entries.append({"is_blank": True})
        entries.append(group)
        last_category = category
    return entries


# --------------------------------------------------------------------------
# Формирование диапазонов вида R1-R10 и пар вида R11,R12
# --------------------------------------------------------------------------

def format_refs_for_prefix(prefix, numbers):
    numbers = sorted(set(numbers))
    ranges = []
    start = prev = numbers[0]
    for num in numbers[1:]:
        if num == prev + 1:
            prev = num
            continue
        ranges.append((start, prev))
        start = prev = num
    ranges.append((start, prev))

    parts = []
    for lo, hi in ranges:
        run_length = hi - lo + 1
        if run_length == 1:
            parts.append("%s%d" % (prefix, lo))
        elif run_length == 2:
            parts.append("%s%d,%s%d" % (prefix, lo, prefix, hi))
        else:
            parts.append("%s%d-%s%d" % (prefix, lo, prefix, hi))
    return parts


def format_group_refdes(refs):
    prefix = get_letter_prefix(refs[0]) if refs else ""
    numbers = []
    text_only = []
    for ref in refs:
        number = get_number_suffix(ref)
        if number is None:
            text_only.append(ref)
        else:
            numbers.append(number)
    parts = format_refs_for_prefix(prefix, numbers) if numbers else []
    parts.extend(text_only)
    return ", ".join(parts)


def build_result_entries(entries):
    result = []
    for entry in entries:
        if entry.get("is_blank") or entry.get("is_header"):
            result.append(entry)
            continue
        refs = entry["refs"]
        if not refs:
            continue
        result.append({
            "refdes": format_group_refdes(refs),
            "name": entry["type"],
            "qty": len(refs),
            "note": entry["note"],
        })
    return result


# --------------------------------------------------------------------------
# Параметры штампа: версия, обозначение, отображаемые значения, трейлер
# --------------------------------------------------------------------------

def load_params(path, base_config):
    config = dict(base_config)
    if not path:
        return config
    with open(path, "r", encoding="utf-8") as f:
        loaded = json.load(f)
    for key, value in loaded.items():
        if key not in config:
            print("Внимание: неизвестный параметр в файле параметров: %s" % key)
            continue
        config[key] = value
    return config


def format_version(value):
    if isinstance(value, int):
        return "%02d" % value
    return str(value)


def build_file_stem(config):
    designation = config.get("document_designation", "")
    doc_type = config.get("document_type", "ПЭ3")
    version = format_version(config.get("version", "01"))
    return "%s%s-В%s" % (designation, doc_type, version)


def build_full_designation(config):
    designation = config.get("document_designation", "")
    doc_type = config.get("document_type", "ПЭ3")
    return ("%s %s" % (designation, doc_type)).strip()


def build_trailer_entries(config):
    designation = config.get("document_designation", "")
    file_stem = build_file_stem(config)
    certified_text = "Удостоверен %s-УЛ" % designation
    filename_text = "Имя файла %s.odt" % file_stem
    return [
        {"is_trailer": True, "text": certified_text},
        {"is_trailer": True, "text": filename_text},
    ]

def build_display_values(config):
    document_designation = build_full_designation(config)
    first_application = config.get("first_application", "") or document_designation
    return {
        "document_designation": document_designation,
        "product_name": config.get("product_name", ""),
        "literature": config.get("literature", ""),
        "developer": config.get("developer", ""),
        "checker": config.get("checker", ""),
        "standard_control": config.get("standard_control", ""),
        "approver": config.get("approver", ""),
        "inventory_number": config.get("inventory_number", ""),
        "first_application": first_application,
    }

# --------------------------------------------------------------------------
# Первая страница: строки "Удостоверен" и "Имя файла" перед штампом
# --------------------------------------------------------------------------

def apply_first_page_trailer(result_entries, rows_on_first_page, trailer_entries):
    slots_for_content = max(rows_on_first_page - len(trailer_entries), 0)

    if len(result_entries) <= slots_for_content:
        page1_content = list(result_entries)
        rest = []
        while len(page1_content) < slots_for_content:
            page1_content.append({"is_blank": True})
    else:
        page1_content = result_entries[:slots_for_content]
        rest = result_entries[slots_for_content:]

    combined = []
    combined.extend(page1_content)
    combined.extend(trailer_entries)
    combined.append({"force_page_break": True})
    combined.extend(rest)
    return combined


# --------------------------------------------------------------------------
# Единое сжатие текста по столбцу
# --------------------------------------------------------------------------

def compute_uniform_percent(entries, field_name, limit_at_100, min_compression):
    max_length = 0
    for entry in entries:
        if entry.get("is_blank") or entry.get("is_header") or entry.get("is_trailer") or entry.get("force_page_break"):
            continue
        length = len(str(entry.get(field_name, "")))
        if length > max_length:
            max_length = length
    if limit_at_100 <= 0 or max_length <= limit_at_100:
        return 100
    floor_percent = int(round(min_compression * 100))
    needed_percent = int(limit_at_100 * 100 / max_length)
    return max(floor_percent, min(needed_percent, 100))


def style_name_for(column, percent):
    base = BASE_STYLE_FOR_COLUMN[column]
    if percent == 100:
        return base
    return "%s_%d" % (base, percent)


def build_scaled_style_xml(name, base, percent):
    scale_attr = 'style:text-scale="%d%%"' % percent
    if base == "P1":
        return (
            '<style:style style:name="%s" style:family="paragraph" style:parent-style-name="Table_20_Contents">'
            '<style:paragraph-properties fo:text-align="center" style:justify-single-word="false"/>'
            '<style:text-properties style:font-name="GOST type A1" fo:font-size="14pt" %s/>'
            "</style:style>"
        ) % (name, scale_attr)
    if base == "P3":
        return (
            '<style:style style:name="%s" style:family="paragraph" style:parent-style-name="Table_20_Contents">'
            '<style:paragraph-properties fo:text-align="center" style:justify-single-word="false"/>'
            '<style:text-properties style:font-name="GOST type A1" fo:font-size="14pt" fo:language="en" fo:country="US" %s/>'
            "</style:style>"
        ) % (name, scale_attr)
    if base == "P5":
        return (
            '<style:style style:name="%s" style:family="paragraph" style:parent-style-name="Table_20_Contents">'
            '<style:paragraph-properties fo:margin-left="0.199cm" fo:margin-right="0.199cm" fo:text-indent="0cm" '
            'style:auto-text-indent="false"><style:tab-stops><style:tab-stop style:position="0.199cm"/></style:tab-stops>'
            "</style:paragraph-properties>"
            '<style:text-properties style:font-name="GOST type A1" fo:font-size="14pt" %s/>'
            "</style:style>"
        ) % (name, scale_attr)
    raise ValueError(base)


def build_all_scaled_styles_xml(min_compression):
    floor_percent = int(round(min_compression * 100))
    parts = []
    for base in ("P1", "P3", "P5"):
        for percent in range(floor_percent, 100):
            parts.append(build_scaled_style_xml("%s_%d" % (base, percent), base, percent))
    return "".join(parts)


# --------------------------------------------------------------------------
# Сборка строк таблицы
# --------------------------------------------------------------------------

def build_data_row_xml(refdes, name, qty, note, percent_a, percent_b, percent_d):
    pos_style = style_name_for("A", percent_a)
    name_style = style_name_for("B", percent_b)
    note_style = style_name_for("D", percent_d)

    return (
        '<table:table-row table:style-name="%s">'
        '<table:table-cell table:style-name="Перечень.A2" office:value-type="string">'
        '<text:p text:style-name="%s">%s</text:p></table:table-cell>'
        '<table:table-cell table:style-name="Перечень.B2" office:value-type="string">'
        '<text:p text:style-name="%s">%s</text:p></table:table-cell>'
        '<table:table-cell table:style-name="Перечень.C2">'
        '<text:p text:style-name="P3">%s</text:p></table:table-cell>'
        '<table:table-cell table:style-name="Перечень.D2" office:value-type="string">'
        '<text:p text:style-name="%s">%s</text:p></table:table-cell>'
        "</table:table-row>"
    ) % (ROW_STYLE_NAME, pos_style, xml_escape(refdes), name_style, xml_escape(name),
         xml_escape(str(qty)), note_style, xml_escape(note))


def build_blank_row_xml(percent_a, percent_b, percent_d):
    return build_data_row_xml("", "", "", "", percent_a, percent_b, percent_d)


def build_header_row_xml(text):
    return (
        '<table:table-row table:style-name="%s">'
        '<table:table-cell table:style-name="Перечень.A2" office:value-type="string">'
        '<text:p text:style-name="P3"/></table:table-cell>'
        '<table:table-cell table:style-name="Перечень.B2" office:value-type="string">'
        '<text:p text:style-name="P1">%s</text:p></table:table-cell>'
        '<table:table-cell table:style-name="Перечень.C2">'
        '<text:p text:style-name="P3"/></table:table-cell>'
        '<table:table-cell table:style-name="Перечень.D2" office:value-type="string">'
        '<text:p text:style-name="P1"/></table:table-cell>'
        "</table:table-row>"
    ) % (ROW_STYLE_NAME, xml_escape(text))


def build_trailer_row_xml(text):
    return (
        '<table:table-row table:style-name="%s">'
        '<table:table-cell table:style-name="Перечень.A2" office:value-type="string">'
        '<text:p text:style-name="P3"/></table:table-cell>'
        '<table:table-cell table:style-name="Перечень.B2" office:value-type="string">'
        '<text:p text:style-name="%s">%s</text:p></table:table-cell>'
        '<table:table-cell table:style-name="Перечень.C2">'
        '<text:p text:style-name="P3"/></table:table-cell>'
        '<table:table-cell table:style-name="Перечень.D2" office:value-type="string">'
        '<text:p text:style-name="P1"/></table:table-cell>'
        "</table:table-row>"
    ) % (ROW_STYLE_NAME, TRAILER_TEXT_STYLE_NAME, xml_escape(text))


def build_rows_xml(result_entries, percent_a, percent_b, percent_d):
    rows = []
    for entry in result_entries:
        if entry.get("force_page_break"):
            rows.append("<text:soft-page-break/>")
        elif entry.get("is_blank"):
            rows.append(build_blank_row_xml(percent_a, percent_b, percent_d))
        elif entry.get("is_trailer"):
            rows.append(build_trailer_row_xml(entry["text"]))
        elif entry.get("is_header"):
            rows.append(build_header_row_xml(entry["text"]))
        else:
            rows.append(build_data_row_xml(
                entry["refdes"], entry["name"], str(entry["qty"]), entry.get("note", ""),
                percent_a, percent_b, percent_d,
            ))
    return "".join(rows)


# --------------------------------------------------------------------------
# Заполнение ODT-шаблона
# --------------------------------------------------------------------------

def extract_column_defs(table_xml):
    parts = re.findall(r"<table:table-column[^/]*/>", table_xml)
    return "".join(parts)


def replace_table_in_content(content_bytes, rows_xml, display_values):
    text = content_bytes.decode("utf-8")
    text = strip_italic(text)
    text = apply_placeholders(text, display_values)

    styles_marker = "</office:automatic-styles>"
    styles_idx = text.index(styles_marker)
    injected_styles = ROW_STYLE_XML + TRAILER_TEXT_STYLE_XML + build_all_scaled_styles_xml(CONFIG["min_compression"])
    text = text[:styles_idx] + injected_styles + text[styles_idx:]

    start_marker = '<table:table table:name="%s"' % TEMPLATE_TABLE_NAME
    start_idx = text.index(start_marker)
    end_idx = text.index("</table:table>", start_idx) + len("</table:table>")
    old_table_xml = text[start_idx:end_idx]

    column_defs_xml = extract_column_defs(old_table_xml)
    new_table_xml = (
        '<table:table table:name="%s" table:style-name="%s">'
        % (TEMPLATE_TABLE_NAME, TEMPLATE_TABLE_NAME)
        + column_defs_xml
        + rows_xml
        + "</table:table>"
    )

    new_text = text[:start_idx] + new_table_xml + text[end_idx:]
    return new_text.encode("utf-8")


def patch_styles_xml(styles_bytes, display_values):
    text = styles_bytes.decode("utf-8")
    text = strip_italic(text)
    text = apply_placeholders(text, display_values)
    text = text.replace('fo:margin-bottom="4.001cm"', 'fo:margin-bottom="2.8cm"')
    return text.encode("utf-8")


def fill_template(template_path, output_path, result_entries, display_values, percent_a, percent_b, percent_d):
    rows_xml = build_rows_xml(result_entries, percent_a, percent_b, percent_d)

    with zipfile.ZipFile(template_path, "r") as zin:
        content_bytes = zin.read("content.xml")
        new_content_bytes = replace_table_in_content(content_bytes, rows_xml, display_values)

        with zipfile.ZipFile(output_path, "w") as zout:
            if "mimetype" in zin.namelist():
                zout.writestr("mimetype", zin.read("mimetype"), compress_type=zipfile.ZIP_STORED)

            for item in zin.infolist():
                if item.filename == "mimetype":
                    continue
                data = zin.read(item.filename)
                if item.filename == "content.xml":
                    data = new_content_bytes
                elif item.filename == "styles.xml":
                    data = patch_styles_xml(data, display_values)
                zout.writestr(item, data)


# --------------------------------------------------------------------------
# Конвертация в PDF в фоне
# --------------------------------------------------------------------------

def convert_to_pdf(odt_path):
    output_dir = os.path.dirname(os.path.abspath(odt_path)) or "."
    subprocess.Popen(
        ["libreoffice", "--headless", "--convert-to", "pdf", "--outdir", output_dir, odt_path],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )


# --------------------------------------------------------------------------
# Точка входа
# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Генерация Перечня элементов ПЭ3 из BOM Cadence Allegro или PADS")
    parser.add_argument("bom_file", help="путь к исходному BOM файлу")
    parser.add_argument("-t", "--template", default="Перечень элементов.odt", help="путь к ODT-шаблону с рамкой")
    parser.add_argument("-o", "--output", help="имя выходного ODT файла", default=None)
    parser.add_argument("-f", "--format", choices=["cadence", "pads"], default=None,
                         help="принудительно задать формат BOM, по умолчанию определяется автоматически")
    parser.add_argument("-p", "--params", default=None, help="путь к JSON файлу с параметрами штампа")
    args = parser.parse_args()

    if not os.path.isfile(args.template):
        print("Не найден файл шаблона: %s" % args.template)
        sys.exit(1)

    config = load_params(args.params, CONFIG)
    file_stem = build_file_stem(config)
    output_path = args.output or ("%s.odt" % file_stem)

    raw_lines = load_raw_lines(args.bom_file)
    format_name = args.format or detect_bom_format(raw_lines)
    print("Определён формат BOM: %s" % format_name)

    fmt = FORMATS[format_name]
    rows = PARSERS[format_name](raw_lines, fmt)

    items = build_items(rows, fmt["header_map"], fmt["merge_order"])
    items = sort_items(items)
    groups = merge_consecutive_items(items)
    groups = split_groups_by_contiguity(groups)

    duplicates = find_duplicate_refs(groups)
    if duplicates:
        print("Внимание: обозначения встречаются в разных позициях перечня:")
        print(", ".join(duplicates))

    entries = insert_category_headers(groups)
    result_entries = build_result_entries(entries)

    percent_a = compute_uniform_percent(result_entries, "refdes", COLUMN_CHAR_LIMITS_AT_100["A"], config["min_compression"])
    percent_b = compute_uniform_percent(result_entries, "name", COLUMN_CHAR_LIMITS_AT_100["B"], config["min_compression"])
    percent_d = compute_uniform_percent(result_entries, "note", COLUMN_CHAR_LIMITS_AT_100["D"], config["min_compression"])

    trailer_entries = build_trailer_entries(config)
    result_entries = apply_first_page_trailer(result_entries, config["rows_on_first_page"], trailer_entries)

    display_values = build_display_values(config)
    fill_template(args.template, output_path, result_entries, display_values, percent_a, percent_b, percent_d)

    print("ODT файл создан: %s" % output_path)
    convert_to_pdf(output_path)
    print("Конвертация в PDF запущена в фоне.")


if __name__ == "__main__":
    main()
