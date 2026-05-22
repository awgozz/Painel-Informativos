# -*- coding: utf-8 -*-
from __future__ import annotations


REQUEST_LINE = (
    "👥 Time SSMAQ, solicitamos apoio para indicação da classificação do ocorrido "
    "com base nas informações acima (Ativo A1 / A2 / A3 ou Informativo)."
)

FIELD_DEFINITIONS = [
    ("Operação", "operacao"),
    ("Motorista", "motorista"),
    ("Placa do cavalo", "placa_cavalo"),
    ("Data e hora", "data_hora"),
    ("Local", "local"),
]

CONTACT_FIELD_DEFINITIONS = [
    ("Contato condutor", "contato_condutor"),
    ("Contato terceiro envolvido", "contato_terceiro"),
]

REQUIRED_FIELD_DEFINITIONS = FIELD_DEFINITIONS + [
    ("Contato condutor", "contato_condutor"),
]

CLASSIFICATION_OPTIONS = [
    "INFORMATIVO",
    "SINISTRO",
    "VAZAMENTO",
    "CONTAMINAÇÃO",
    "ROUBO/FURTO",
    "OCLUSÃO DE CAMERA",
    "TOMBAMENTO",
]
