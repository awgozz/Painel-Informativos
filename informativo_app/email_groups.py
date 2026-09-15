# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import re
import unicodedata
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET


APP_DATA_DIR = Path(os.environ.get("APPDATA", Path.home())) / "PainelInformativosGarbuio"
EMAIL_GROUPS_FILE = APP_DATA_DIR / "email_groups.json"
SHEET_NAME = "Grupo de e-mails"

EMAIL_RE = re.compile(r"[\w.\-+%]+@[\w.\-]+\.[A-Za-z]{2,}")
NS = {
    "m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


class EmailGroupError(Exception):
    """Erro ao importar ou carregar grupos de e-mail."""


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKD", value or "")
    value = "".join(char for char in value if not unicodedata.combining(char))
    value = re.sub(r"\s+", " ", value).strip().upper()
    return value


def normalize_emails(value: str) -> list[str]:
    seen: set[str] = set()
    emails: list[str] = []
    for email in EMAIL_RE.findall(value or ""):
        normalized = email.strip().lower()
        if normalized and normalized not in seen:
            seen.add(normalized)
            emails.append(normalized)
    return emails


def merge_emails(*groups: list[str]) -> list[str]:
    seen: set[str] = set()
    merged: list[str] = []
    for group in groups:
        for email in group:
            if email not in seen:
                seen.add(email)
                merged.append(email)
    return merged


def groups_file_path() -> Path:
    return EMAIL_GROUPS_FILE


def load_email_groups() -> list[dict[str, object]]:
    if not EMAIL_GROUPS_FILE.exists():
        return []
    try:
        data = json.loads(EMAIL_GROUPS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    groups = data.get("groups", [])
    return groups if isinstance(groups, list) else []


def save_email_groups(groups: list[dict[str, object]], source_path: str | Path) -> None:
    try:
        APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
        payload = {
            "source_path": str(source_path),
            "imported_at": datetime.now().isoformat(timespec="seconds"),
            "groups": groups,
        }
        EMAIL_GROUPS_FILE.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError as error:
        raise EmailGroupError(
            "Não consegui salvar os grupos de e-mail no perfil do Windows."
        ) from error


def find_email_group(
    groups: list[dict[str, object]],
    operation: str,
) -> dict[str, object] | None:
    target = normalize_text(operation)
    for group in groups:
        if normalize_text(str(group.get("operacao", ""))) == target:
            return group
    return None


def import_groups_from_xlsx(file_path: str | Path) -> list[dict[str, object]]:
    file_path = Path(file_path)
    if not file_path.exists():
        raise EmailGroupError("A planilha selecionada não foi encontrada.")

    try:
        with zipfile.ZipFile(file_path) as workbook:
            shared_strings = _read_shared_strings(workbook)
            sheet_path = _find_sheet_path(workbook, SHEET_NAME)
            rows = _read_sheet_rows(workbook, sheet_path, shared_strings)
    except zipfile.BadZipFile as error:
        raise EmailGroupError("O arquivo selecionado não parece ser uma planilha .xlsx válida.") from error
    except KeyError as error:
        raise EmailGroupError("Não consegui ler a estrutura interna da planilha.") from error

    header_index, columns = _find_header(rows)
    groups: list[dict[str, object]] = []

    for row in rows[header_index + 1 :]:
        operation = row.get(columns["operacao"], "").strip()
        if not operation or operation == "-":
            continue

        groups.append(
            {
                "operacao": operation,
                "programador": normalize_emails(row.get(columns["programador"], "")),
                "coordenador_gerente": normalize_emails(
                    row.get(columns["coordenador_gerente"], "")
                ),
                "ssmaq": normalize_emails(row.get(columns["ssmaq"], "")),
                "adicionais_cc": normalize_emails(row.get(columns["adicionais_cc"], "")),
                "relacionados": normalize_emails(row.get(columns["relacionados"], "")),
            }
        )

    if not groups:
        raise EmailGroupError("Nenhuma operação com e-mails foi encontrada na planilha.")
    return groups


def _read_shared_strings(workbook: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in workbook.namelist():
        return []

    root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall("m:si", NS):
        strings.append("".join(text.text or "" for text in item.findall(".//m:t", NS)))
    return strings


def _find_sheet_path(workbook: zipfile.ZipFile, sheet_name: str) -> str:
    workbook_xml = ET.fromstring(workbook.read("xl/workbook.xml"))
    relationships_xml = ET.fromstring(workbook.read("xl/_rels/workbook.xml.rels"))
    relationships = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in relationships_xml
    }

    fallback_path = ""
    for sheet in workbook_xml.findall("m:sheets/m:sheet", NS):
        relationship_id = sheet.attrib.get(
            "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
        )
        target = relationships.get(relationship_id or "", "")
        if target:
            path = "xl/" + target.lstrip("/")
            fallback_path = fallback_path or path
            if normalize_text(sheet.attrib.get("name", "")) == normalize_text(sheet_name):
                return path

    if fallback_path:
        return fallback_path
    raise EmailGroupError("Não encontrei nenhuma aba na planilha.")


def _read_sheet_rows(
    workbook: zipfile.ZipFile,
    sheet_path: str,
    shared_strings: list[str],
) -> list[dict[int, str]]:
    sheet_xml = ET.fromstring(workbook.read(sheet_path))
    rows_by_number: dict[int, dict[int, str]] = {}

    for row in sheet_xml.findall("m:sheetData/m:row", NS):
        row_number = int(row.attrib.get("r", "0") or "0") - 1
        values: dict[int, str] = {}
        for cell in row.findall("m:c", NS):
            column_index = _column_index(cell.attrib.get("r", ""))
            values[column_index] = _cell_value(cell, shared_strings)
        rows_by_number[row_number] = values

    _fill_merged_cells(sheet_xml, rows_by_number)

    rows: list[dict[int, str]] = []
    for row_number in sorted(rows_by_number):
        values = rows_by_number[row_number]
        if any(value.strip() for value in values.values()):
            rows.append(values)
    return rows


def _fill_merged_cells(
    sheet_xml: ET.Element,
    rows_by_number: dict[int, dict[int, str]],
) -> None:
    for merge in sheet_xml.findall("m:mergeCells/m:mergeCell", NS):
        reference = merge.attrib.get("ref", "")
        start, end = reference.split(":") if ":" in reference else (reference, reference)
        start_row, start_column = _cell_position(start)
        end_row, end_column = _cell_position(end)
        merged_value = rows_by_number.get(start_row, {}).get(start_column, "")
        if not merged_value:
            continue

        for row_number in range(start_row, end_row + 1):
            row = rows_by_number.setdefault(row_number, {})
            for column_index in range(start_column, end_column + 1):
                if not row.get(column_index, "").strip():
                    row[column_index] = merged_value


def _cell_position(cell_reference: str) -> tuple[int, int]:
    row_match = re.search(r"\d+", cell_reference)
    row_number = int(row_match.group(0)) - 1 if row_match else 0
    return row_number, _column_index(cell_reference)


def _cell_value(cell: ET.Element, shared_strings: list[str]) -> str:
    if cell.attrib.get("t") == "inlineStr":
        return "".join(text.text or "" for text in cell.findall(".//m:t", NS))

    value = cell.find("m:v", NS)
    if value is None:
        return ""

    text = value.text or ""
    if cell.attrib.get("t") == "s" and text.isdigit():
        index = int(text)
        return shared_strings[index] if index < len(shared_strings) else ""
    return text


def _column_index(cell_reference: str) -> int:
    letters = re.match(r"[A-Z]+", cell_reference.upper())
    if not letters:
        return 0

    index = 0
    for letter in letters.group(0):
        index = index * 26 + (ord(letter) - ord("A") + 1)
    return index - 1


def _find_header(rows: list[dict[int, str]]) -> tuple[int, dict[str, int]]:
    for row_index, row in enumerate(rows):
        headers = {column: normalize_text(value) for column, value in row.items()}
        columns = {
            "operacao": _find_column(headers, "OPERACAO"),
            "programador": _find_column(headers, "PROGRAMADOR"),
            "coordenador_gerente": _find_column(headers, "COORDENADOR", "GERENTE"),
            "ssmaq": _find_column(headers, "SSMAQ"),
            "adicionais_cc": _find_column(headers, "ADICIONAIS"),
            "relacionados": _find_column(headers, "SINISTROS", "INFORMATIVOS", "RELACIONADOS"),
        }
        if columns["operacao"] is not None and columns["programador"] is not None:
            return row_index, {key: value if value is not None else -1 for key, value in columns.items()}

    raise EmailGroupError("Não encontrei a linha de cabeçalho da planilha.")


def _find_column(headers: dict[int, str], *needles: str) -> int | None:
    for column, header in headers.items():
        if all(needle in header for needle in needles):
            return column
    return None
