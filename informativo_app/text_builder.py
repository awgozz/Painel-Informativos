# -*- coding: utf-8 -*-
from __future__ import annotations

from .config import REQUEST_LINE
from .utils import extract_header_date


def get_classification(data: dict[str, str]) -> str:
    return data.get("classificacao", "").strip() or "INFORMATIVO"


def build_text(data: dict[str, str], include_request_line: bool) -> str:
    lines = [
        f"➡️ Operação: {data.get('operacao', '').strip()}",
        f"➡️ Motorista: {data.get('motorista', '').strip()}",
        f"➡️ Placa do cavalo: {data.get('placa_cavalo', '').strip()}",
        f"➡️ Data e hora: {data.get('data_hora', '').strip()}",
        f"➡️ Local: {data.get('local', '').strip()}",
        "DESCRIÇÃO:",
        data.get("descricao", "").strip(),
    ]

    if include_request_line:
        lines.extend(["", REQUEST_LINE])

    return "\n".join(lines)


def build_word_header(data: dict[str, str]) -> str:
    return (
        f"{get_classification(data)} – {extract_header_date(data.get('data_hora', ''))} – "
        f"{data.get('motorista', '').strip()} - "
        f"{data.get('placa_cavalo', '').strip()} – "
        f"{data.get('operacao', '').strip()}"
    )


def build_word_intro_lines(data: dict[str, str]) -> list[str]:
    return [
        "Prezados, bom dia.",
        (
            f"Segue a abertura do {get_classification(data)} referente ao veículo "
            f"de placa {data.get('placa_cavalo', '').strip()}, ocorrido em "
            f"{extract_header_date(data.get('data_hora', ''))}:"
        ),
        "",
    ]


def build_attachment_lines(data: dict[str, str]) -> list[str]:
    lines = [
        "Documentos Anexos:",
        "• Segue anexo a CNH e CRLV;",
        "• Segue imagens do ocorrido;",
        f"• Contato condutor: {data.get('contato_condutor', '').strip()};",
    ]

    contato_terceiro = data.get("contato_terceiro", "").strip()
    if contato_terceiro:
        lines.append(f"• Contato terceiro envolvido: {contato_terceiro};")

    return lines
