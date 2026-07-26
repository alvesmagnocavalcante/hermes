from __future__ import annotations

import re
import unicodedata
import warnings
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from automations.excel_reader import load_workbook_compatible as load_workbook

TOLERANCE = Decimal("0.01")
ENTRY_ACCOUNTS = {
    "ALIMENTOS": "Alimentos",
    "VINHOECHAMPANHE": "Vinhos & Champanhe",
    "BEBIDASALCOOLICAS": "Alcoolicos",
    "BEBIDASNAOALCOOLICAS": "Não Alcoolicos",
    "FRIGOBAR": "Frigobar",
}
INVENTORY_CODES = {
    "Alimentos": ("01",),
    "Vinhos & Champanhe": ("02",),
    "Alcoolicos": ("03",),
    "Não Alcoolicos": ("04",),
    "Frigobar": ("05",),
    "Mimos Hospedes": ("06", "0706"),
    "Amenitees": ("0701",),
    "Material de Higiene e Limpeza": ("0702",),
    "Material de Escritório/Informatica": ("0704", "0711"),
    "Decoracao": ("0708",),
    "Eletroeletronicos": ("0709",),
    "Suprimentos de uso do Hospedes": ("0713", "0806"),
    "Material de Copa e Cozinha": ("0802", "0804", "2001"),
    "Uniforme": ("0805",),
    "Material de Manutenção de Edifícios e Instalações": (
        "0901",
        "0902",
        "0904",
        "0909",
    ),
    "Material de Manutenção de Maquinas e Equipamentos": ("0908", "0910"),
    "Material de Manutenção da Piscina": ("0903",),
    "Materal de Reposicao": ("0705",),
    "Equipamento de Protecao": ("0907",),
    "SPA": ("10",),
}


@dataclass(frozen=True)
class Row:
    analysis: str
    account: str
    source: Decimal
    accounting: Decimal

    @property
    def difference(self) -> Decimal:
        return self.source - self.accounting

    @property
    def status(self) -> str:
        return "Conciliado" if abs(self.difference) <= TOLERANCE else "Divergente"


def normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return re.sub(
        r"[^A-Z0-9]",
        "",
        "".join(c for c in text if not unicodedata.combining(c)).upper(),
    )


def number(value: Any) -> Decimal:
    if value in (None, "", "NULL"):
        return Decimal()
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return Decimal(str(value).replace(".", "").replace(",", "."))


def money(value: Decimal) -> str:
    return f"R$ {value:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def read(path: Path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    sheet.reset_dimensions()
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    return rows[0], rows[1:]


def identify(paths: list[Path]):
    files = {}
    for path in paths:
        name = normalize(path.name)
        if "DOCUMENTOSLANCADOS" in name:
            key = "documents"
        elif "RAZAOANALITICOESTOQUEAB" in name:
            key = "entry_ledger"
        elif "INVENTARIOFISICO" in name:
            key = "inventory"
        elif "RAZAOANALITICOESTOQUES" in name:
            key = "stock_ledger"
        else:
            raise ValueError(f"Arquivo não reconhecido: {path.name}")
        files[key] = read(path)
    if set(files) != {"documents", "entry_ledger", "inventory", "stock_ledger"}:
        raise ValueError("Selecione os quatro arquivos da Atividade 10.")
    return files


def grouped(data, key: str, value: str):
    header, rows = data
    ki, vi = header.index(key), header.index(value)
    result = defaultdict(Decimal)
    for row in rows:
        if len(row) > max(ki, vi) and row[ki] not in (None, "", "NULL"):
            result[str(row[ki]).strip()] += number(row[vi])
    return result


def analyze(paths: list[Path]) -> list[Row]:
    if len(paths) != 4:
        raise ValueError("Selecione exatamente os quatro arquivos da Atividade 10.")
    files = identify(paths)
    documents_raw = grouped(files["documents"], "DESCRICAOTDESEMB", "VALORLANÇADO")
    documents = {normalize(key): value for key, value in documents_raw.items()}
    entry_ledger = grouped(files["entry_ledger"], "DescricaoConta", "Debito")
    entry_normalized = {normalize(key): value for key, value in entry_ledger.items()}
    result = [
        Row(
            "Entradas",
            account,
            documents.get(key, Decimal()),
            entry_normalized.get(normalize(account), Decimal()),
        )
        for key, account in ENTRY_ACCOUNTS.items()
    ]

    inventory = grouped(files["inventory"], "GrupoCodigo", "SaldoValor")
    header, ledger_rows = files["stock_ledger"]
    name_i, balance_i = header.index("DescricaoConta"), header.index("SaldoAtual")
    final_balances = {}
    for row in ledger_rows:
        if len(row) > max(name_i, balance_i) and row[name_i] not in (None, "", "NULL"):
            final_balances[str(row[name_i]).strip()] = number(row[balance_i])
    for account, codes in INVENTORY_CODES.items():
        inventory_value = sum(
            (
                value
                for code, value in inventory.items()
                if any(str(code).startswith(prefix) for prefix in codes)
            ),
            Decimal(),
        )
        ledger_value = next(
            (
                value
                for name, value in final_balances.items()
                if normalize(name) == normalize(account)
            ),
            Decimal(),
        )
        result.append(Row("Saldo final", account, inventory_value, ledger_value))
    return result


def export_excel(rows: list[Row], path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Conferência de Custos"
    sheet.append(
        ["Análise", "Conta", "CAP / Inventário", "Contabilidade", "Diferença", "Status"]
    )
    for row in rows:
        sheet.append(
            [
                row.analysis,
                row.account,
                float(row.source),
                float(row.accounting),
                float(row.difference),
                row.status,
            ]
        )
    for cell in sheet[1]:
        cell.style = "Headline 4"
    for column in "CDE":
        for cell in sheet[column][1:]:
            cell.number_format = "R$ #,##0.00"
    for column, width in {"A": 18, "B": 52, "C": 22, "D": 22, "E": 20, "F": 16}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(path)


def export_pdf(rows: list[Row], path: Path) -> None:
    styles = getSampleStyleSheet()
    data = [
        ["Análise", "Conta", "CAP / Inventário", "Contabilidade", "Diferença", "Status"]
    ]
    data.extend(
        [
            [
                r.analysis,
                r.account,
                money(r.source),
                money(r.accounting),
                money(r.difference),
                r.status,
            ]
            for r in rows
        ]
    )
    table = Table(
        data,
        repeatRows=1,
        colWidths=[27 * mm, 64 * mm, 38 * mm, 38 * mm, 36 * mm, 28 * mm],
    )
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#24588A")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 7),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
            ]
        )
    )
    doc = SimpleDocTemplate(
        str(path),
        pagesize=landscape(A4),
        title="Conferência dos Custos da Mercadoria Vendida",
    )
    doc.build(
        [
            Paragraph("Conferência dos Custos da Mercadoria Vendida", styles["Title"]),
            Spacer(1, 5 * mm),
            table,
        ]
    )
