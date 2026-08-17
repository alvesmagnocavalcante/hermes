from __future__ import annotations

import re
import unicodedata
import warnings
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from automations.excel_reader import load_workbook_compatible as load_workbook


# Normaliza chaves textuais usadas para relacionar registros entre relatórios.
def normalize_key(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    plain = "".join(char for char in text if not unicodedata.combining(char))
    return re.sub(r"[^A-Z0-9]", "", plain.upper())


def identifier(value: Any) -> str:
    text = str(value or "").strip()
    return (
        str(int(float(text)))
        if re.fullmatch(r"\d+(?:\.0+)?", text)
        else normalize_key(text)
    )


# Converte números de diferentes formatos e padroniza a apresentação brasileira.
def decimal_value(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal()
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return Decimal(str(value).replace(".", "").replace(",", "."))


def nullable_decimal(value: Any) -> Decimal:
    if value in (None, "", "NULL"):
        return Decimal()
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return Decimal(str(value).replace(".", "").replace(",", "."))


def report_decimal(value: Any) -> Decimal:
    if value in (None, "", "NULL", "-"):
        return Decimal()
    text = str(value).strip().replace("R$", "").replace(" ", "")
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal()


def money(value: Decimal) -> str:
    return f"R$ {value:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def optional_money(value: Decimal | None) -> str:
    return "—" if value is None else money(value)


def integer(value: int) -> str:
    return f"{value:,}".replace(",", ".")


# Abre a primeira aba e devolve cabeçalho e linhas em um formato comum.
def active_sheet_rows(path: Path) -> tuple[tuple[Any, ...], list[tuple[Any, ...]]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    sheet.reset_dimensions()
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    return rows[0], rows[1:]
