# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from datetime import datetime


def current_date_time() -> str:
    return datetime.now().strftime("%d/%m/%Y %H:%M")


def clean_filename(value: str) -> str:
    value = re.sub(r"[^\w\s-]", "", value, flags=re.UNICODE).strip()
    value = re.sub(r"[\s_-]+", "_", value)
    return value[:50] or "informativo"


def extract_header_date(value: str) -> str:
    value = value.strip()
    match = re.search(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", value)
    return match.group(0) if match else value
