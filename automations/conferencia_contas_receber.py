from __future__ import annotations

import re
import tkinter as tk
import unicodedata
import warnings
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any

import customtkinter as ctk
from openpyxl import Workbook, load_workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from automations.base import Automation
from automations.ui import TableColumn, clear_table, create_result_table, result_tag


TOLERANCE = Decimal("0.01")


@dataclass(frozen=True)
class ClientRow:
    client: str
    accounting: Decimal
    financial: Decimal

    @property
    def difference(self) -> Decimal:
        return self.accounting - self.financial

    @property
    def status(self) -> str:
        return "Conciliado" if abs(self.difference) <= TOLERANCE else "Divergente"


@dataclass(frozen=True)
class TotalCheck:
    name: str
    source_value: Decimal
    accounting_value: Decimal

    @property
    def difference(self) -> Decimal:
        return self.source_value - self.accounting_value

    @property
    def status(self) -> str:
        return "Conciliado" if abs(self.difference) <= TOLERANCE else "Divergente"


@dataclass(frozen=True)
class ReceivablesResult:
    clients: list[ClientRow]
    client_accounting_total: Decimal
    client_financial_total: Decimal
    billing: TotalCheck
    commissions: TotalCheck


def normalize_name(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    text = "".join(char for char in text if not unicodedata.combining(char)).upper()
    return re.sub(r"[^A-Z0-9]", "", text)


def decimal_value(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal()
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return Decimal(str(value).replace(".", "").replace(",", "."))


def currency(value: Decimal) -> str:
    text = f"R$ {value:,.2f}"
    return text.replace(",", "_").replace(".", ",").replace("_", ".")


def integer(value: int) -> str:
    return f"{value:,}".replace(",", ".")


def load_rows(path: Path) -> tuple[tuple[Any, ...], list[tuple[Any, ...]]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook.active
        sheet.reset_dimensions()
        rows = list(sheet.iter_rows(values_only=True))
        if not rows:
            raise ValueError(f"{path.name}: planilha vazia.")
        return rows[0], rows[1:]
    finally:
        workbook.close()


def identify_files(paths: list[Path]) -> dict[str, tuple[tuple[Any, ...], list[tuple[Any, ...]]]]:
    identified: dict[str, tuple[tuple[Any, ...], list[tuple[Any, ...]]]] = {}
    for path in paths:
        header, rows = load_rows(path)
        headers = set(header)
        if {"DescricaoSubconta", "Saldo"}.issubset(headers):
            kind = "balancete"
        elif {"Cliente", "Saldo", "ContaContabilRateio"}.issubset(headers):
            kind = "posicao"
        elif {"Valor", "NumeroDaTransacao", "Status"}.issubset(headers):
            kind = "bordero"
        elif {"DescricaoConta", "Debito", "Historico"}.issubset(headers) and "Notas a Faturar" in path.name:
            kind = "razao_faturar"
        elif {"NumeroDocumento", "Valor", "IdAgregado"}.issubset(headers):
            kind = "agregados"
        elif {"DescricaoConta", "Movimento", "Historico"}.issubset(headers):
            kind = "razao_comissao"
        else:
            raise ValueError(f"{path.name}: arquivo não reconhecido para esta conferência.")
        if kind in identified:
            raise ValueError(f"Dois arquivos foram identificados como {kind}.")
        identified[kind] = (header, rows)
    expected = {"balancete", "posicao", "bordero", "razao_faturar", "agregados", "razao_comissao"}
    if set(identified) != expected:
        missing = ", ".join(sorted(expected - set(identified)))
        raise ValueError(f"Selecione os seis arquivos da Atividade 8. Ausentes: {missing}.")
    return identified


def grouped(header: tuple[Any, ...], rows: list[tuple[Any, ...]], name: str, value: str) -> dict[str, tuple[str, Decimal]]:
    name_i, value_i = header.index(name), header.index(value)
    result: dict[str, tuple[str, Decimal]] = {}
    sums: defaultdict[str, Decimal] = defaultdict(Decimal)
    labels: dict[str, str] = {}
    for row in rows:
        if len(row) <= max(name_i, value_i) or row[name_i] in (None, "", "NULL"):
            continue
        key = normalize_name(row[name_i])
        labels.setdefault(key, str(row[name_i]).strip())
        sums[key] += decimal_value(row[value_i])
    for key, value_sum in sums.items():
        result[key] = (labels[key], value_sum)
    return result


def column_total(header: tuple[Any, ...], rows: list[tuple[Any, ...]], column: str, absolute: bool = False) -> Decimal:
    index = header.index(column)
    total = Decimal()
    for row in rows:
        if len(row) <= index:
            continue
        value = decimal_value(row[index])
        total += abs(value) if absolute else value
    return total


def analyze(paths: list[Path]) -> ReceivablesResult:
    if len(paths) != 6:
        raise ValueError("Selecione exatamente os seis arquivos da Atividade 8.")
    files = identify_files(paths)

    accounting = grouped(*files["balancete"], "DescricaoSubconta", "Saldo")
    financial = grouped(*files["posicao"], "Cliente", "Saldo")
    clients = [
        ClientRow(
            accounting.get(key, financial.get(key))[0],
            accounting.get(key, ("", Decimal()))[1],
            financial.get(key, ("", Decimal()))[1],
        )
        for key in set(accounting) | set(financial)
    ]
    clients.sort(key=lambda row: (row.status == "Conciliado", -abs(row.difference), row.client))

    bordero_total = column_total(*files["bordero"], "Valor", absolute=True)
    billing_debit = column_total(*files["razao_faturar"], "Debito")
    aggregate_total = column_total(*files["agregados"], "Valor", absolute=True)
    commission_movement = column_total(*files["razao_comissao"], "Movimento", absolute=True)
    return ReceivablesResult(
        clients,
        sum((value for _, value in accounting.values()), Decimal()),
        sum((value for _, value in financial.values()), Decimal()),
        TotalCheck("Notas a faturar", bordero_total, billing_debit),
        TotalCheck("Comissões de cartão", aggregate_total, commission_movement),
    )


def save_excel(result: ReceivablesResult, path: Path) -> None:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumo"
    summary.append(["Conferência", "Origem", "Contabilidade", "Diferença", "Status"])
    summary.append(["Clientes", float(result.client_financial_total), float(result.client_accounting_total),
                    float(result.client_financial_total - result.client_accounting_total),
                    "Conciliado" if abs(result.client_financial_total - result.client_accounting_total) <= TOLERANCE else "Divergente"])
    for check in (result.billing, result.commissions):
        summary.append([check.name, float(check.source_value), float(check.accounting_value), float(check.difference), check.status])
    for cell in summary[1]:
        cell.style = "Headline 4"
    for column in ("B", "C", "D"):
        for cell in summary[column][1:]:
            cell.number_format = 'R$ #,##0.00'
    for column, width in {"A": 28, "B": 20, "C": 20, "D": 18, "E": 16}.items():
        summary.column_dimensions[column].width = width

    details = workbook.create_sheet("Clientes")
    details.append(["Cliente / Subconta", "Saldo Contabilidade", "Saldo Financeiro", "Diferença", "Status"])
    for row in result.clients:
        details.append([row.client, float(row.accounting), float(row.financial), float(row.difference), row.status])
    for cell in details[1]:
        cell.style = "Headline 4"
    for column in ("B", "C", "D"):
        for cell in details[column][1:]:
            cell.number_format = 'R$ #,##0.00'
    for column, width in {"A": 48, "B": 22, "C": 22, "D": 18, "E": 16}.items():
        details.column_dimensions[column].width = width
    details.freeze_panes = "A2"
    details.auto_filter.ref = details.dimensions
    workbook.save(path)


def save_pdf(result: ReceivablesResult, path: Path) -> None:
    styles = getSampleStyleSheet()
    document = SimpleDocTemplate(str(path), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                                 topMargin=12 * mm, bottomMargin=12 * mm, title="Conferência do Contas a Receber")
    client_diff = result.client_financial_total - result.client_accounting_total
    data = [["Conferência", "Origem", "Contabilidade", "Diferença", "Status"],
            ["Clientes", currency(result.client_financial_total), currency(result.client_accounting_total),
             currency(client_diff), "Conciliado" if abs(client_diff) <= TOLERANCE else "Divergente"]]
    for check in (result.billing, result.commissions):
        data.append([check.name, currency(check.source_value), currency(check.accounting_value),
                     currency(check.difference), check.status])
    table = Table(data, colWidths=[52 * mm, 45 * mm, 45 * mm, 42 * mm, 35 * mm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#24588A")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("ALIGN", (1, 1), (-2, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), .5, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    document.build([Paragraph("Conferência do Contas a Receber", styles["Title"]), Spacer(1, 6 * mm), table,])


class ReceivablesAutomation(Automation):
    name = "Conferência do Contas a Receber"

    def __init__(self, app, container: ctk.CTkFrame) -> None:
        super().__init__(app, container)
        self.paths: list[Path] = []
        self.result: ReceivablesResult | None = None
        self.output_format = ctk.StringVar(value="Excel")
        self.filter_status = ctk.StringVar(value="Divergentes")
        self.search_text = ctk.StringVar()
        self.page = 0
        self.page_size = 100

    def render(self) -> None:
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(6, weight=1)
        ctk.CTkLabel(self.container, text=self.name, font=ctk.CTkFont(size=26, weight="bold")).grid(
            row=0, column=0, padx=30, pady=(22, 3), sticky="w")
        ctk.CTkLabel(self.container, text="Confere clientes, notas a faturar e comissões integradas ao financeiro.",
                     text_color="gray70").grid(row=1, column=0, padx=30, pady=(0, 10), sticky="w")
        controls = ctk.CTkFrame(self.container, fg_color="transparent")
        controls.grid(row=2, column=0, padx=30, sticky="ew")
        self.select_button = ctk.CTkButton(controls, text="Selecionar os seis arquivos", command=self._select)
        self.select_button.pack(side="left", padx=(0, 10))
        self.format_selector = ctk.CTkSegmentedButton(controls, values=["Excel", "PDF"], variable=self.output_format)
        self.format_selector.pack(side="left", padx=10)
        self.export_button = ctk.CTkButton(controls, text="Exportar resultado", state="disabled", command=self._export)
        self.export_button.pack(side="left", padx=10)
        self.clear_button = ctk.CTkButton(controls, text="Limpar", fg_color="gray35", command=self._clear)
        self.clear_button.pack(side="left", padx=10)
        self.files_label = ctk.CTkLabel(self.container, text="Selecione os seis arquivos.", text_color="gray70", anchor="w")
        self.files_label.grid(row=3, column=0, padx=30, pady=(10, 6), sticky="ew")

        dashboard = ctk.CTkFrame(self.container, fg_color="transparent")
        dashboard.grid(row=4, column=0, padx=30, pady=(0, 8), sticky="ew")
        dashboard.grid_columnconfigure(0, weight=3)
        dashboard.grid_columnconfigure(1, weight=2)
        summary = ctk.CTkFrame(dashboard, fg_color="transparent")
        summary.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        self.check_labels: dict[str, tuple[ctk.CTkLabel, ctk.CTkLabel]] = {}
        for column, (title, key) in enumerate((("Clientes", "clients"), ("Notas a faturar", "billing"), ("Comissões", "commissions"))):
            summary.grid_columnconfigure(column, weight=1)
            card = ctk.CTkFrame(summary)
            card.grid(row=0, column=column, padx=(0 if column == 0 else 7, 0), sticky="nsew")
            ctk.CTkLabel(card, text=title, font=ctk.CTkFont(size=14, weight="bold")).pack(pady=(7, 0))
            values = ctk.CTkLabel(card, text="—", text_color="gray70")
            values.pack()
            status = ctk.CTkLabel(card, text="Aguardando", font=ctk.CTkFont(weight="bold"))
            status.pack(pady=(0, 7))
            self.check_labels[key] = (values, status)

        chart_frame = ctk.CTkFrame(dashboard)
        chart_frame.grid(row=0, column=1, sticky="nsew")
        chart_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            chart_frame,
            text="Conciliação por cliente",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="center",
        ).grid(row=0, column=0, pady=(6, 0), sticky="ew")
        self.chart = tk.Canvas(chart_frame, height=125, background="#2b2b2b", highlightthickness=0)
        self.chart.grid(row=1, column=0, padx=8, sticky="ew")
        self.chart.bind("<Configure>", lambda _: self._draw_chart())

        filters = ctk.CTkFrame(self.container, fg_color="transparent")
        filters.grid(row=5, column=0, padx=30, pady=(0, 8), sticky="ew")
        filters.grid_columnconfigure(2, weight=1)
        self.status_selector = ctk.CTkSegmentedButton(filters, values=["Divergentes", "Conciliados", "Todos"],
                                                       variable=self.filter_status, command=lambda _: self._reset_filter())
        self.status_selector.grid(row=0, column=0, padx=(0, 12))
        ctk.CTkLabel(filters, text="Buscar cliente:").grid(row=0, column=1, padx=(0, 8))
        self.search_entry = ctk.CTkEntry(filters, textvariable=self.search_text, placeholder_text="Digite cliente ou subconta")
        self.search_entry.grid(row=0, column=2, sticky="ew")
        self.search_entry.bind("<KeyRelease>", lambda _: self._reset_filter())
        self.preview = create_result_table(self.container, (
            TableColumn("client", "Cliente / Subconta", 390),
            TableColumn("accounting", "Contabilidade", 145, "e"),
            TableColumn("financial", "Financeiro", 145, "e"),
            TableColumn("difference", "Diferença", 145, "e"),
            TableColumn("status", "Situação", 130),
        ), row=6)
        pagination = ctk.CTkFrame(self.container, fg_color="transparent")
        pagination.grid(row=7, column=0, padx=30, pady=(7, 14), sticky="ew")
        pagination.grid_columnconfigure(1, weight=1)
        self.previous_button = ctk.CTkButton(pagination, text="Anterior", width=100, command=self._previous)
        self.previous_button.grid(row=0, column=0)
        self.page_label = ctk.CTkLabel(pagination, text="Página 0 de 0")
        self.page_label.grid(row=0, column=1)
        self.next_button = ctk.CTkButton(pagination, text="Próxima", width=100, command=self._next)
        self.next_button.grid(row=0, column=2)
        self._render_preview()

    def _select(self) -> None:
        names = filedialog.askopenfilenames(title="Selecionar os seis arquivos da Atividade 8",
                                            filetypes=[("Planilhas Excel", "*.xlsx")])
        if not names:
            return
        self.paths = [Path(name) for name in names]
        self._busy(True)
        self.app.set_status("Conferindo contas a receber...", .1)
        self.app.run_background(lambda: analyze(self.paths), self._completed, self._failed)

    def _completed(self, result: ReceivablesResult) -> None:
        self.result = result
        self.page = 0
        self._busy(False)
        self._render_preview()
        self.app.set_status("Conferência concluída", 1)

    def _failed(self) -> None:
        self._busy(False)

    def _filtered(self) -> list[ClientRow]:
        if not self.result:
            return []
        selected, search = self.filter_status.get(), normalize_name(self.search_text.get())
        return [row for row in self.result.clients if
                (selected == "Todos" or (selected == "Divergentes" and row.status == "Divergente")
                 or (selected == "Conciliados" and row.status == "Conciliado"))
                and (not search or search in normalize_name(row.client))]

    def _render_preview(self) -> None:
        clear_table(self.preview)
        if not self.result:
            self.files_label.configure(text="Selecione os seis arquivos.")
            for values, status in self.check_labels.values():
                values.configure(text="—")
                status.configure(text="Aguardando", text_color="gray70")
            self.page_label.configure(text="Página 0 de 0")
            self.previous_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
        else:
            result = self.result
            client_diff = result.client_financial_total - result.client_accounting_total
            client_status = "Conciliado" if abs(client_diff) <= TOLERANCE else "Divergente"
            checks = {
                "clients": (f"Financeiro {currency(result.client_financial_total)} | Contábil {currency(result.client_accounting_total)}", client_status),
                "billing": (f"Borderô {currency(result.billing.source_value)} | Razão {currency(result.billing.accounting_value)}", result.billing.status),
                "commissions": (f"Agregados {currency(result.commissions.source_value)} | Razão {currency(result.commissions.accounting_value)}", result.commissions.status),
            }
            for key, (text, status_text) in checks.items():
                values, status = self.check_labels[key]
                values.configure(text=text)
                status.configure(text=f"{status_text} • Dif. {currency({'clients': client_diff, 'billing': result.billing.difference, 'commissions': result.commissions.difference}[key])}",
                                 text_color="#21a67a" if status_text == "Conciliado" else "#dc5a5a")
            self.files_label.configure(text=f"Arquivos carregados: {len(self.paths)} • Clientes: {len(result.clients)}")
            filtered = self._filtered()
            total_pages = max(1, (len(filtered) + self.page_size - 1) // self.page_size)
            self.page = min(self.page, total_pages - 1)
            page_rows = filtered[self.page * self.page_size:(self.page + 1) * self.page_size]
            for row in page_rows:
                self.preview.insert("", "end", values=(row.client, currency(row.accounting), currency(row.financial),
                    currency(row.difference), row.status), tags=(result_tag(row.status),))
            self.page_label.configure(text=f"Página {self.page + 1} de {total_pages} • {integer(len(filtered))} clientes")
            self.previous_button.configure(state="normal" if self.page else "disabled")
            self.next_button.configure(state="normal" if self.page + 1 < total_pages else "disabled")
        self.export_button.configure(state="normal" if self.result else "disabled")
        self.app.after_idle(self._draw_chart)

    def _draw_chart(self) -> None:
        self.chart.delete("all")
        width = max(self.chart.winfo_width(), 300)
        center_x, center_y, radius = width * 0.34, 53, 43
        bounds = (center_x - radius, center_y - radius, center_x + radius, center_y + radius)
        if not self.result or not self.result.clients:
            self.chart.create_oval(*bounds, outline="#4b5563", width=13)
            self.chart.create_text(center_x, center_y, text="—", fill="#e5e7eb")
            self.chart.create_text(width * 0.58, center_y, text="Aguardando conferência", fill="#cbd5e1", anchor="w")
            return

        total = len(self.result.clients)
        reconciled = sum(row.status == "Conciliado" for row in self.result.clients)
        divergent = total - reconciled
        reconciled_rate = reconciled / total
        divergent_rate = divergent / total
        split = -360 * reconciled_rate
        self.chart.create_arc(*bounds, start=90, extent=split, style="arc", outline="#21a67a", width=13)
        self.chart.create_arc(
            *bounds, start=90 + split, extent=-360 * divergent_rate,
            style="arc", outline="#dc5a5a", width=13,
        )
        self.chart.create_text(
            center_x, center_y, text=f"{reconciled_rate:.1%}", fill="#e5e7eb", font=("Segoe UI", 10, "bold")
        )
        self.chart.create_text(
            width * 0.58, 40, text=f"● Conciliados: {integer(reconciled)}", fill="#21a67a", anchor="w"
        )
        self.chart.create_text(
            width * 0.58, 67, text=f"● Divergentes: {integer(divergent)}", fill="#dc5a5a", anchor="w"
        )

    def _reset_filter(self) -> None:
        self.page = 0
        self._render_preview()

    def _previous(self) -> None:
        if self.page:
            self.page -= 1
            self._render_preview()

    def _next(self) -> None:
        if (self.page + 1) * self.page_size < len(self._filtered()):
            self.page += 1
            self._render_preview()

    def _export(self) -> None:
        if not self.result:
            return
        output_format = self.output_format.get()
        extension = ".pdf" if output_format == "PDF" else ".xlsx"
        name = filedialog.asksaveasfilename(title="Salvar conferência", defaultextension=extension,
                                            initialfile=f"conferencia_contas_receber{extension}",
                                            filetypes=[(f"Arquivo {output_format}", f"*{extension}")])
        if not name:
            return
        self._busy(True)
        task = lambda: save_pdf(self.result, Path(name)) if output_format == "PDF" else save_excel(self.result, Path(name))
        self.app.run_background(task, lambda _: self._saved(Path(name)), self._failed)

    def _saved(self, path: Path) -> None:
        self._busy(False)
        self.app.set_status("Resultado exportado", 1)
        messagebox.showinfo("Exportação concluída", f"Arquivo salvo em:\n{path}")

    def _clear(self) -> None:
        self.paths.clear()
        self.result = None
        self.search_text.set("")
        self.page = 0
        self._render_preview()
        self.app.set_status("Seleção limpa", 0)

    def _busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.select_button.configure(state=state)
        self.clear_button.configure(state=state)
        self.format_selector.configure(state=state)
        self.status_selector.configure(state=state)
        self.search_entry.configure(state=state)
        self.export_button.configure(state="disabled" if busy or not self.result else "normal")


AUTOMATION_CLASS = ReceivablesAutomation
