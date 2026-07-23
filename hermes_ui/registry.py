from __future__ import annotations

import importlib
from dataclasses import dataclass, fields, is_dataclass
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any


@dataclass(frozen=True)
class Column:
    key: str
    label: str
    numeric: bool = False


@dataclass(frozen=True)
class AutomationSpec:
    key: str
    name: str
    description: str
    module: str
    analyzer: str
    rows_attribute: str | None
    columns: tuple[Column, ...]
    extensions: tuple[str, ...] = ("xlsx", "xlsm", "xls", "csv", "xml")
    formats: tuple[str, ...] = ("Excel", "PDF")
    hotel_option: bool = False

    def analyze(self, paths: list[Path], hotel: str) -> Any:
        function = getattr(importlib.import_module(self.module), self.analyzer)
        return function(paths, hotel) if self.hotel_option else function(paths)

    def export(self, result: Any, output: Path, output_format: str) -> None:
        module = importlib.import_module(self.module)
        names = {
            "Excel": "save_excel",
            "PDF": "save_pdf",
            "CSV": "save_csv",
        }
        function_name = names[output_format]
        if self.key == "receita":
            function_name += "_result"
        elif self.key == "custos":
            function_name = f"export_{output_format.lower()}"
        function = getattr(module, function_name)
        function(result, output)

    def rows(self, result: Any) -> list[Any]:
        if self.key == "receber":
            rows = [
                {
                    "category": "Clientes",
                    "name": row.client,
                    "accounting": row.accounting,
                    "financial": row.financial,
                    "difference": row.difference,
                    "status": row.status,
                }
                for row in result.clients
            ]
            for category, check in (("Notas a faturar", result.billing), ("Comissões", result.commissions)):
                rows.append({
                    "category": category, "name": check.name,
                    "accounting": check.accounting_value, "financial": check.source_value,
                    "difference": check.difference, "status": check.status,
                })
            return rows
        if self.rows_attribute:
            return list(getattr(result, self.rows_attribute))
        return list(result)


def record_value(record: Any, key: str) -> Any:
    if isinstance(record, dict):
        return record.get(key)
    return getattr(record, key, None)


def record_status(record: Any) -> str:
    value = record_value(record, "status")
    if value:
        return str(value)
    reconciled = record_value(record, "reconciled")
    if isinstance(reconciled, bool):
        return "Conciliado" if reconciled else "Divergente"
    difference = record_value(record, "difference")
    if isinstance(difference, Decimal):
        return "Conciliado" if abs(difference) <= Decimal("0.01") else "Divergente"
    return "Processado"


def searchable(record: Any, columns: tuple[Column, ...]) -> str:
    return " ".join(str(record_value(record, column.key) or "") for column in columns).casefold()


def format_value(value: Any) -> str:
    if value is None or value == "":
        return "—"
    if isinstance(value, bool):
        return "Sim" if value else "Não"
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, Decimal):
        sign = "-" if value < 0 else ""
        formatted = f"{abs(value):,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")
        return f"{sign}R$ {formatted}"
    return str(value)


def result_details(result: Any) -> dict[str, str]:
    if not is_dataclass(result):
        return {}
    ignored = {"rows", "coupons", "entities", "clients"}
    labels = {
        "company": "Empresa", "hotel": "Hotel", "journal_rows": "Linhas do Journal",
        "cmflex_total": "Total Contabilidade", "opera_total": "Total Opera",
        "simphony_total": "Total Simphony", "fiscal_total": "Total Fiscal",
        "sefaz_total": "Total SEFAZ", "cancelled": "Cancelados",
        "external_count": "Notas externas", "cap_count": "Notas CAP",
        "matched_count": "Localizadas no CAP", "approved_count": "BPM aprovadas",
        "retained_count": "ISS Prefeitura", "cap_retained_count": "ISS CAP",
        "expected_hotel": "Hotel esperado", "fiscal_start": "Início do Fiscal",
        "fiscal_end": "Fim do Fiscal", "service_profile": "Hotel identificado", "earnings": "Proventos",
        "deductions": "Descontos", "net_payable": "Líquido a pagar",
        "ignored_rows": "Totalizadores removidos",
        "excluded_rows": "Eventos desconsiderados",
    }
    details = {}
    for field in fields(result):
        key = field.name
        value = getattr(result, key)
        if key in ignored or is_dataclass(value) or isinstance(value, (list, dict, tuple)):
            continue
        details[labels.get(key, key.replace("_", " ").title())] = format_value(value)
    return details


C = Column
SPECS = (
    AutomationSpec("receita", "Conciliação de Receita", "Compara os movimentos da Contabilidade com os lançamentos do Opera.",
                   "automations.conciliacao_receita", "reconcile", "rows",
                   (C("document", "Documento"), C("cmflex", "Contabilidade", True), C("opera", "Opera", True), C("difference", "Diferença", True), C("status", "Resultado"))),
    AutomationSpec("diarias", "Conciliação da Receita de Diárias", "Valida as receitas de diárias do Journal pelos códigos de transação configurados.",
                   "automations.conciliacao_receita_diarias", "analyze", "rows",
                   (C("trx_code", "TRX Code"), C("description", "Descrição"), C("daily", "Diária"), C("average_daily", "Diária média"), C("transactions", "Transações", True), C("value", "Valor", True), C("status", "Resultado")), hotel_option=True),
    AutomationSpec("folha", "Lançamento da Folha de Pagamento", "Gera os lançamentos contábeis por centro de custo sem totalizadores duplicados.",
                   "automations.lancamento_folha_pagamento", "analyze", "rows",
                   (C("source", "Saída"), C("organogram", "Organograma"), C("cost_center", "Centro de custo"), C("description", "Descrição"), C("debit", "Débito"), C("credit", "Crédito"), C("value", "Valor", True), C("accounting_origin", "Origem contábil"), C("status", "Resultado")), formats=("Excel", "CSV", "PDF")),
    AutomationSpec("cupons_hospede", "Cupons Emitidos x Conta do Hóspede", "Confere se os cupons emitidos constam e foram cobrados na conta do hóspede.",
                   "automations.conciliacao_cupons_hospedes", "analyze", "coupons",
                   (C("company", "Hotel"), C("document", "Cupom"), C("issue_date", "Emissão"), C("posting_date", "Lançamento"), C("room", "Quarto"), C("guest", "Hóspede"), C("pdv_value", "PDV", True), C("journal_value", "Conta", True), C("difference", "Diferença", True), C("status", "Resultado"), C("detail", "Explicação"))),
    AutomationSpec("rps", "RPS de Serviços Prestados", "Confere os RPS encerrados no Opera, integrados no Fiscal e emitidos na Prefeitura.",
                   "automations.conferencia_rps_servicos_prestados", "analyze", "rows",
                   (C("rps", "RPS"), C("opera_date", "Data Opera"), C("customer", "Hóspede/Tomador"), C("opera_value", "Opera", True), C("fiscal_value", "Fiscal", True), C("city_value", "Prefeitura", True), C("difference", "Diferença", True), C("city_nfse", "NFS-e"), C("status", "Resultado"), C("detail", "Explicação")), extensions=("xml", "xlsx", "xlsm", "xls")),
    AutomationSpec("debito", "Relatório de Notas de Débito", "Consolida as notas de débito por hotel, comprador, emissão, item e valor.",
                   "automations.relatorio_notas_debito", "extract", None,
                   (C("hotel", "Hotel"), C("comprador", "Comprador"), C("nota", "Nota"), C("emissao", "Emissão"), C("item", "Item"), C("valor", "Valor", True)), extensions=("xlsx", "xlsm")),
    AutomationSpec("entrada", "Notas Fiscais de Entrada em Atraso", "Classifica as notas do Manifesto como em dia, em alerta ou em atraso, mesmo quando ainda não foram lançadas.",
                   "automations.notas_entrada_atrasadas", "analyze", "rows",
                   (C("key", "Chave"), C("company", "Empresa"), C("supplier", "Fornecedor"), C("state", "UF"), C("emission_date", "Emissão"), C("entry_date", "Entrada"), C("days", "Dias", True), C("limit", "Limite", True), C("launch_status", "Lançamento"), C("status", "Situação"))),
    AutomationSpec("cupons", "Conferência dos Cupons", "Compara Simphony, Fiscal e SEFAZ por chave, data e valor.",
                   "automations.conferencia_cupons", "reconcile", "rows",
                   (C("document_type", "Tipo"), C("key", "Chave fiscal"), C("simphony_date", "Data Simphony"), C("fiscal_date", "Data Fiscal"), C("sefaz_date", "Data SEFAZ"), C("simphony", "Simphony", True), C("fiscal", "Fiscal", True), C("sefaz", "SEFAZ", True), C("difference", "Diferença", True), C("status", "Resultado"))),
    AutomationSpec("servicos", "Notas de Serviços Tomados", "Compara documentos externos, CAP, BPM, hotel e ISS retido da Prefeitura.",
                   "automations.conferencia_notas_servicos_tomados", "analyze", "rows",
                   (C("source", "Fonte"), C("provider", "Prestador"), C("cnpj", "CNPJ"), C("number", "Nota"), C("emission_date", "Data"), C("gross", "Valor externo", True), C("cap_gross", "Valor CAP", True), C("iss", "ISS Prefeitura", True), C("cap_iss", "ISS CAP", True), C("bpm", "BPM"), C("cap_hotel", "Hotel CAP"), C("status", "Resultado"))),
    AutomationSpec("receber", "Conferência do Contas a Receber", "Confere clientes, notas a faturar e comissões contra o financeiro.",
                   "automations.conferencia_contas_receber", "analyze", None,
                   (C("category", "Conferência"), C("name", "Cliente/Conta"), C("accounting", "Contabilidade", True), C("financial", "Financeiro", True), C("difference", "Diferença", True), C("status", "Resultado"))),
    AutomationSpec("pagar", "Conferência do Contas a Pagar", "Confere fornecedores, adiantamentos e impostos contra o financeiro.",
                   "automations.conferencia_contas_pagar", "analyze", "entities",
                   (C("category", "Conferência"), C("name", "Fornecedor/Subconta"), C("accounting", "Contabilidade", True), C("financial", "Financeiro", True), C("difference", "Diferença", True), C("status", "Resultado"))),
    AutomationSpec("custos", "Custos da Mercadoria Vendida", "Compara entradas do CAP e saldos do inventário com a Contabilidade.",
                   "automations.conferencia_custos_mercadoria", "analyze", None,
                   (C("analysis", "Conferência"), C("account", "Conta/Grupo"), C("source", "Origem", True), C("accounting", "Contabilidade", True), C("difference", "Diferença", True), C("status", "Resultado"))),
)
