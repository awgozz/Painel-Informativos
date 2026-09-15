# -*- coding: utf-8 -*-
from __future__ import annotations

import html

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


def build_email_subject(data: dict[str, str]) -> str:
    return build_word_header(data)


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


def build_email_html(data: dict[str, str]) -> str:
    intro_lines = build_word_intro_lines(data)
    attachment_lines = build_attachment_lines(data)
    description = html.escape(data.get("descricao", "").strip()).replace("\n", "<br>")

    field_rows = [
        ("Operação", data.get("operacao", "")),
        ("Motorista", data.get("motorista", "")),
        ("Placa do cavalo", data.get("placa_cavalo", "")),
        ("Data e hora", data.get("data_hora", "")),
        ("Local", data.get("local", "")),
    ]
    fields_html = "".join(
        "<p style='margin:4px 0;'>"
        f"<strong>{html.escape(label)}:</strong> {html.escape(value.strip())}"
        "</p>"
        for label, value in field_rows
    )
    attachments_html = "".join(
        f"<li>{html.escape(line.removeprefix('• ').strip())}</li>"
        for line in attachment_lines
        if line != "Documentos Anexos:"
    )

    return f"""\
<html>
  <body style="font-family: Calibri, Arial, sans-serif; font-size: 14pt; color: #1f1f1f;">
    <p style="margin:0 0 10px 0;">{html.escape(intro_lines[0])}<br>{html.escape(intro_lines[1])}</p>
    {fields_html}
    <p style="margin:12px 0 4px 0;"><strong>DESCRIÇÃO:</strong></p>
    <p style="margin:0 0 12px 0;">{description}</p>
    <p style="margin:12px 0 4px 0;"><strong>Documentos Anexos:</strong></p>
    <ul style="margin-top:0;">{attachments_html}</ul>
  </body>
</html>
"""
