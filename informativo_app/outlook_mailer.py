# -*- coding: utf-8 -*-
from __future__ import annotations

import base64
import json
import os
import subprocess
import tempfile

from .email_groups import merge_emails
from .text_builder import build_email_html, build_email_subject, get_classification


RELATED_CLASSIFICATIONS = {
    "INFORMATIVO",
    "ATIVO LEVE (A1)",
    "ATIVO MEDIO (A2)",
    "ATIVO GRAVE (A3)",
    "PESSOA LEVE (P1)",
    "PESSOA MEDIO(P2)",
    "PESSOA GRAVE(P3)",
}


class OutlookDraftError(Exception):
    """Erro ao criar rascunho no Outlook."""


def create_outlook_draft(data: dict[str, str], group: dict[str, object]) -> None:
    to_emails, cc_emails = build_recipients(data, group)
    if not to_emails:
        raise OutlookDraftError("Não encontrei destinatários no grupo de e-mail desta operação.")

    subject = build_email_subject(data)
    html_body = build_email_html(data)

    try:
        import win32com.client  # type: ignore[import-not-found]
    except ImportError:
        _create_outlook_draft_with_powershell(to_emails, cc_emails, subject, html_body)
        return

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.CreateItem(0)
        message.To = "; ".join(to_emails)
        message.CC = "; ".join(cc_emails)
        message.Subject = subject
        message.HTMLBody = html_body
        message.Display()
    except Exception as error:
        raise OutlookDraftError(
            "Não foi possível criar o rascunho no Outlook. "
            "Confira se o Outlook desktop está instalado e com a conta configurada."
        ) from error


def build_recipients(
    data: dict[str, str],
    group: dict[str, object],
) -> tuple[list[str], list[str]]:
    to_emails = merge_emails(
        _email_list(group.get("programador")),
        _email_list(group.get("coordenador_gerente")),
        _email_list(group.get("ssmaq")),
    )

    if get_classification(data) in RELATED_CLASSIFICATIONS:
        to_emails = merge_emails(to_emails, _email_list(group.get("relacionados")))

    cc_emails = merge_emails(_email_list(group.get("adicionais_cc")))
    return to_emails, cc_emails


def _email_list(value: object) -> list[str]:
    if isinstance(value, list):
        return [str(item) for item in value if str(item).strip()]
    return []


def _create_outlook_draft_with_powershell(
    to_emails: list[str],
    cc_emails: list[str],
    subject: str,
    html_body: str,
) -> None:
    payload = {
        "To": "; ".join(to_emails),
        "CC": "; ".join(cc_emails),
        "Subject": subject,
        "HTMLBody": html_body,
    }

    payload_path = ""
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".json",
            prefix="painel_informativo_email_",
            encoding="utf-8",
            delete=False,
        ) as payload_file:
            payload_path = payload_file.name
            json.dump(payload, payload_file, ensure_ascii=False)

        path_token = base64.b64encode(payload_path.encode("utf-8")).decode("ascii")
        script = f"""
$payloadPath = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('{path_token}'))
$payload = Get-Content -LiteralPath $payloadPath -Raw -Encoding UTF8 | ConvertFrom-Json
$outlook = New-Object -ComObject Outlook.Application
$message = $outlook.CreateItem(0)
$message.To = $payload.To
$message.CC = $payload.CC
$message.Subject = $payload.Subject
$message.HTMLBody = $payload.HTMLBody
$message.Display()
"""
        encoded_script = base64.b64encode(script.encode("utf-16le")).decode("ascii")
        startup_info = None
        creation_flags = 0
        if os.name == "nt":
            startup_info = subprocess.STARTUPINFO()
            startup_info.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            creation_flags = subprocess.CREATE_NO_WINDOW

        result = subprocess.run(
            [
                "powershell.exe",
                "-NoProfile",
                "-ExecutionPolicy",
                "Bypass",
                "-EncodedCommand",
                encoded_script,
            ],
            capture_output=True,
            text=True,
            timeout=60,
            startupinfo=startup_info,
            creationflags=creation_flags,
            check=False,
        )
    except Exception as error:
        raise OutlookDraftError(
            "Não foi possível criar o rascunho no Outlook. "
            "Confira se o Outlook desktop está instalado e com a conta configurada."
        ) from error
    finally:
        if payload_path:
            try:
                os.unlink(payload_path)
            except OSError:
                pass

    if result.returncode != 0:
        raise OutlookDraftError(
            "Não foi possível criar o rascunho no Outlook. "
            "Confira se o Outlook desktop está instalado e com a conta configurada."
        )
