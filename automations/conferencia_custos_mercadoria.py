from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from automations.common import active_sheet_rows as read
from automations.common import money, normalize_key as normalize
from automations.common import nullable_decimal as number

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


# Modelo de uma linha consolidada de custo por conta contábil.
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


# Identifica CAP, Inventário e Contabilidade e agrupa seus valores.
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


def entry_postings(data) -> dict[str, Decimal]:
    header, rows = data
    name_i = header.index("DescricaoConta")
    debit_i = header.index("Debito")
    history_i = header.index("Historico")
    result = defaultdict(Decimal)
    required_terms = ("NOTAFISCAL", "MERCADORIA", "TERCEIROS")
    for row in rows:
        if len(row) <= max(name_i, debit_i, history_i):
            continue
        account = row[name_i]
        history = normalize(row[history_i])
        if account not in (None, "", "NULL") and all(
            term in history for term in required_terms
        ):
            result[str(account).strip()] += number(row[debit_i])
    return result


# Mantém a conta corrente para capturar saldos em linhas sem descrição.
def final_balances(data) -> dict[str, Decimal]:
    header, rows = data
    name_i, balance_i = header.index("DescricaoConta"), header.index("SaldoAtual")
    result: dict[str, Decimal] = {}
    current_account = ""
    for row in rows:
        if len(row) > name_i and row[name_i] not in (None, "", "NULL"):
            current_account = str(row[name_i]).strip()
        if (
            current_account
            and len(row) > balance_i
            and row[balance_i] not in (None, "", "NULL")
        ):
            result[current_account] = number(row[balance_i])
    return result


# Compara os custos das fontes e classifica diferenças por conta.
def analyze(paths: list[Path]) -> list[Row]:
    if len(paths) != 4:
        raise ValueError("Selecione exatamente os quatro arquivos da Atividade 10.")
    files = identify(paths)
    documents_raw = grouped(files["documents"], "DESCRICAOTDESEMB", "VALORLANÇADO")
    documents = {normalize(key): value for key, value in documents_raw.items()}
    entry_ledger = entry_postings(files["entry_ledger"])
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
    accounting_balances = final_balances(files["stock_ledger"])
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
                for name, value in accounting_balances.items()
                if normalize(name) == normalize(account)
            ),
            Decimal(),
        )
        result.append(Row("Saldo final", account, inventory_value, ledger_value))
    return result


# Exporta a conferência de custos em Excel e PDF.
def export_excel(rows: list[Row], path: Path) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for analysis, source_name in (
        ("Entradas", "CAP"),
        ("Saldo final", "Inventário"),
    ):
        sheet = workbook.create_sheet(analysis)
        sheet.append(
            ["Conta", source_name, "Contabilidade", "Diferença", "Status"]
        )
        for row in rows:
            if row.analysis == analysis:
                sheet.append(
                    [
                        row.account,
                        float(row.source),
                        float(row.accounting),
                        float(row.difference),
                        row.status,
                    ]
                )
        for cell in sheet[1]:
            cell.style = "Headline 4"
        for column in "BCD":
            for cell in sheet[column][1:]:
                cell.number_format = "R$ #,##0.00"
        for column, width in {
            "A": 52,
            "B": 22,
            "C": 22,
            "D": 20,
            "E": 16,
        }.items():
            sheet.column_dimensions[column].width = width
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
    workbook.save(path)


def export_pdf(rows: list[Row], path: Path) -> None:
    styles = getSampleStyleSheet()
    story = [
        Paragraph("Conferência dos Custos da Mercadoria Vendida", styles["Title"]),
        Spacer(1, 5 * mm),
    ]
    for analysis, source_name in (
        ("Entradas", "CAP"),
        ("Saldo final", "Inventário"),
    ):
        data = [["Conta", source_name, "Contabilidade", "Diferença", "Status"]]
        data.extend(
            [
                [
                    row.account,
                    money(row.source),
                    money(row.accounting),
                    money(row.difference),
                    row.status,
                ]
                for row in rows
                if row.analysis == analysis
            ]
        )
        table = Table(
            data,
            repeatRows=1,
            colWidths=[75 * mm, 40 * mm, 40 * mm, 38 * mm, 28 * mm],
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
        story.extend(
            [Paragraph(analysis, styles["Heading2"]), Spacer(1, 2 * mm), table]
        )
    doc = SimpleDocTemplate(
        str(path),
        pagesize=landscape(A4),
        title="Conferência dos Custos da Mercadoria Vendida",
    )
    doc.build(story)
