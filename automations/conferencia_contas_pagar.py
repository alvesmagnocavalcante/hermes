from __future__ import annotations

import re
import unicodedata
import warnings
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from automations.legacy_ui import ctk, filedialog, messagebox, tk
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
class EntityRow:
    category: str
    name: str
    accounting: Decimal
    financial: Decimal

    @property
    def difference(self) -> Decimal:
        return self.financial - self.accounting

    @property
    def status(self) -> str:
        return "Conciliado" if abs(self.difference) <= TOLERANCE else "Divergente"


@dataclass(frozen=True)
class Check:
    name: str
    financial: Decimal
    accounting: Decimal

    @property
    def difference(self) -> Decimal:
        return self.financial - self.accounting

    @property
    def status(self) -> str:
        return "Conciliado" if abs(self.difference) <= TOLERANCE else "Divergente"


@dataclass(frozen=True)
class PayablesResult:
    entities: list[EntityRow]
    suppliers: Check
    advances: Check
    taxes: tuple[Check, Check, Check]


def normalize(value: Any) -> str:
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


def read(path: Path) -> tuple[tuple[Any, ...], list[tuple[Any, ...]]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook.active
        sheet.reset_dimensions()
        rows = list(sheet.iter_rows(values_only=True))
        return rows[0], rows[1:]
    finally:
        workbook.close()


def identify(paths: list[Path]) -> dict[str, tuple[tuple[Any, ...], list[tuple[Any, ...]]]]:
    files: dict[str, tuple[tuple[Any, ...], list[tuple[Any, ...]]]] = {}
    for path in paths:
        header, rows = read(path)
        name = normalize(path.name)
        if "BALANCETECOMSUBCONTA" in name and "ADTO" in name:
            key = "balancete_adto"
        elif "BALANCETECOMSUBCONTA" in name:
            key = "balancete_fornecedor"
        elif "POSICAOPORFORNECEDOR" in name:
            key = "posicao_fornecedor"
        elif "ADIANTAMENTOSEMABERTO" in name:
            key = "adiantamentos"
        elif "AGREGADOSLANCADOS" in name and "IRRF" in name:
            key = "agregado_irrf"
        elif "AGREGADOSLANCADOS" in name and "CSRF" in name:
            key = "agregado_csrf"
        elif "AGREGADOSLANCADOS" in name and "ISS" in name:
            key = "agregado_iss"
        elif "RAZAOANALITICO" in name:
            key = "razao_impostos"
        else:
            raise ValueError(f"{path.name}: arquivo não reconhecido para esta conferência.")
        files[key] = (header, rows)
    expected = {"balancete_adto", "balancete_fornecedor", "posicao_fornecedor", "adiantamentos",
                "agregado_irrf", "agregado_csrf", "agregado_iss", "razao_impostos"}
    if set(files) != expected:
        raise ValueError("Selecione os oito arquivos da Atividade 9.")
    return files


def grouped(data, name_column: str, value_column: str) -> dict[str, tuple[str, Decimal]]:
    header, rows = data
    name_i, value_i = header.index(name_column), header.index(value_column)
    sums: defaultdict[str, Decimal] = defaultdict(Decimal)
    labels: dict[str, str] = {}
    for row in rows:
        if len(row) <= max(name_i, value_i) or row[name_i] in (None, "", "NULL"):
            continue
        key = normalize(row[name_i])
        labels.setdefault(key, str(row[name_i]).strip())
        sums[key] += decimal_value(row[value_i])
    return {key: (labels[key], abs(value)) for key, value in sums.items()}


def entities(category: str, accounting, financial) -> list[EntityRow]:
    result: list[EntityRow] = []
    accounting_left = dict(accounting)
    financial_left = dict(financial)

    for key in set(accounting_left) & set(financial_left):
        accounting_name, accounting_value = accounting_left.pop(key)
        _, financial_value = financial_left.pop(key)
        result.append(EntityRow(category, accounting_name, accounting_value, financial_value))

    candidates = sorted(
        [
            (
            SequenceMatcher(None, accounting_key, financial_key).ratio(),
            accounting_key,
            financial_key,
            )
            for accounting_key, (_, accounting_value) in accounting_left.items()
            for financial_key, (_, financial_value) in financial_left.items()
            if accounting_value != 0 and abs(accounting_value - financial_value) <= TOLERANCE
        ],
        reverse=True,
    )
    for _, accounting_key, financial_key in candidates:
        if accounting_key not in accounting_left or financial_key not in financial_left:
            continue
        accounting_name, accounting_value = accounting_left.pop(accounting_key)
        financial_name, financial_value = financial_left.pop(financial_key)
        label = accounting_name if accounting_name == financial_name else f"{accounting_name} / {financial_name}"
        result.append(EntityRow(category, label, accounting_value, financial_value))

    result.extend(EntityRow(category, name, value, Decimal()) for name, value in accounting_left.values())
    result.extend(EntityRow(category, name, Decimal(), value) for name, value in financial_left.values())
    return result


def absolute_total(data, column: str) -> Decimal:
    header, rows = data
    index = header.index(column)
    return sum((abs(decimal_value(row[index])) for row in rows if len(row) > index), Decimal())


def analyze(paths: list[Path]) -> PayablesResult:
    if len(paths) != 8:
        raise ValueError("Selecione exatamente os oito arquivos da Atividade 9.")
    files = identify(paths)
    supplier_accounting = grouped(files["balancete_fornecedor"], "DescricaoSubconta", "Saldo")
    supplier_financial = grouped(files["posicao_fornecedor"], "Fornecedor", "Saldo")
    advance_accounting = grouped(files["balancete_adto"], "DescricaoSubconta", "Saldo")
    advance_financial = grouped(files["adiantamentos"], "NomeFornecedor", "Saldo")
    rows = entities("Fornecedores", supplier_accounting, supplier_financial)
    rows.extend(entities("Adiantamentos", advance_accounting, advance_financial))
    rows.sort(key=lambda row: (row.status == "Conciliado", -abs(row.difference), row.name))

    tax_financial = {
        "IRRF": absolute_total(files["agregado_irrf"], "Valor"),
        "CSRF": absolute_total(files["agregado_csrf"], "Valor"),
        "ISS": absolute_total(files["agregado_iss"], "Valor"),
    }
    header, ledger_rows = files["razao_impostos"]
    name_i, movement_i = header.index("DescricaoConta"), header.index("Movimento")
    tax_accounting = {"IRRF": Decimal(), "CSRF": Decimal(), "ISS": Decimal()}
    for row in ledger_rows:
        if len(row) <= max(name_i, movement_i):
            continue
        movement = decimal_value(row[movement_i])
        if movement >= 0:
            continue
        account = normalize(row[name_i])
        tax = "IRRF" if "IRRF" in account else "CSRF" if "PISCOFINSCSLL" in account else "ISS" if "ISSRETIDO" in account else None
        if tax:
            tax_accounting[tax] += abs(movement)

    supplier_accounting_total = sum((value for _, value in supplier_accounting.values()), Decimal())
    supplier_financial_total = sum((value for _, value in supplier_financial.values()), Decimal())
    advance_accounting_total = sum((value for _, value in advance_accounting.values()), Decimal())
    advance_financial_total = sum((value for _, value in advance_financial.values()), Decimal())
    return PayablesResult(
        rows,
        Check("Fornecedores", supplier_financial_total, supplier_accounting_total),
        Check("Adiantamentos", advance_financial_total, advance_accounting_total),
        tuple(Check(tax, tax_financial[tax], tax_accounting[tax]) for tax in ("IRRF", "CSRF", "ISS")),
    )


def save_excel(result: PayablesResult, path: Path) -> None:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumo"
    summary.append(["Conferência", "Financeiro", "Contabilidade", "Diferença", "Status"])
    for check in (result.suppliers, result.advances, *result.taxes):
        summary.append([check.name, float(check.financial), float(check.accounting), float(check.difference), check.status])
    for cell in summary[1]:
        cell.style = "Headline 4"
    for column in ("B", "C", "D"):
        for cell in summary[column][1:]:
            cell.number_format = 'R$ #,##0.00'
    for column, width in {"A": 25, "B": 20, "C": 20, "D": 18, "E": 16}.items():
        summary.column_dimensions[column].width = width
    details = workbook.create_sheet("Fornecedores e Adiantamentos")
    details.append(["Tipo", "Fornecedor / Subconta", "Contabilidade", "Financeiro", "Diferença", "Status"])
    for row in result.entities:
        details.append([row.category, row.name, float(row.accounting), float(row.financial), float(row.difference), row.status])
    for cell in details[1]:
        cell.style = "Headline 4"
    for column in ("C", "D", "E"):
        for cell in details[column][1:]:
            cell.number_format = 'R$ #,##0.00'
    for column, width in {"A": 18, "B": 50, "C": 20, "D": 20, "E": 18, "F": 16}.items():
        details.column_dimensions[column].width = width
    details.freeze_panes = "A2"
    details.auto_filter.ref = details.dimensions
    workbook.save(path)


def save_pdf(result: PayablesResult, path: Path) -> None:
    styles = getSampleStyleSheet()
    document = SimpleDocTemplate(str(path), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
                                 topMargin=12 * mm, bottomMargin=12 * mm, title="Conferência do Contas a Pagar")
    data = [["Conferência", "Financeiro", "Contabilidade", "Diferença", "Status"]]
    for check in (result.suppliers, result.advances, *result.taxes):
        data.append([check.name, currency(check.financial), currency(check.accounting), currency(check.difference), check.status])
    table = Table(data, colWidths=[48 * mm, 44 * mm, 44 * mm, 42 * mm, 35 * mm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#24588A")), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("ALIGN", (1, 1), (-2, -1), "RIGHT"),
        ("GRID", (0, 0), (-1, -1), .5, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    document.build([Paragraph("Conferência do Contas a Pagar", styles["Title"]), Spacer(1, 6 * mm), table])


class PayablesAutomation(Automation):
    name = "Conferência do Contas a Pagar"

    def __init__(self, app, container: ctk.CTkFrame) -> None:
        super().__init__(app, container)
        self.paths: list[Path] = []
        self.result: PayablesResult | None = None
        self.output_format = ctk.StringVar(value="Excel")
        self.category = ctk.StringVar(value="Fornecedores")
        self.search_text = ctk.StringVar()
        self.page = 0
        self.page_size = 100

    def render(self) -> None:
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(6, weight=1)
        ctk.CTkLabel(self.container, text=self.name, font=ctk.CTkFont(size=26, weight="bold")).grid(
            row=0, column=0, padx=30, pady=(22, 3), sticky="w")
        ctk.CTkLabel(self.container, text="Confere fornecedores, adiantamentos e impostos integrados ao financeiro.",
                     text_color="gray70").grid(row=1, column=0, padx=30, pady=(0, 10), sticky="w")
        controls = ctk.CTkFrame(self.container, fg_color="transparent")
        controls.grid(row=2, column=0, padx=30, sticky="ew")
        self.select_button = ctk.CTkButton(controls, text="Selecionar os oito arquivos", command=self._select)
        self.select_button.pack(side="left", padx=(0, 10))
        self.format_selector = ctk.CTkSegmentedButton(controls, values=["Excel", "PDF"], variable=self.output_format)
        self.format_selector.pack(side="left", padx=10)
        self.export_button = ctk.CTkButton(controls, text="Exportar resultado", state="disabled", command=self._export)
        self.export_button.pack(side="left", padx=10)
        self.clear_button = ctk.CTkButton(controls, text="Limpar", fg_color="gray35", command=self._clear)
        self.clear_button.pack(side="left", padx=10)
        self.files_label = ctk.CTkLabel(self.container, text="Selecione os oito arquivos.", text_color="gray70", anchor="w")
        self.files_label.grid(row=3, column=0, padx=30, pady=(10, 6), sticky="ew")

        dashboard = ctk.CTkFrame(self.container, fg_color="transparent")
        dashboard.grid(row=4, column=0, padx=30, pady=(0, 8), sticky="ew")
        for column in range(6):
            dashboard.grid_columnconfigure(column, weight=1)
        self.check_labels: dict[str, tuple[ctk.CTkLabel, ctk.CTkLabel]] = {}
        for column, key in enumerate(("Fornecedores", "Adiantamentos", "IRRF", "CSRF", "ISS")):
            card = ctk.CTkFrame(dashboard)
            card.grid(row=0, column=column, padx=(0 if column == 0 else 5, 0), sticky="nsew")
            ctk.CTkLabel(card, text=key, font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(5, 0))
            values = ctk.CTkLabel(card, text="—", text_color="gray70")
            values.pack()
            status = ctk.CTkLabel(card, text="Aguardando", font=ctk.CTkFont(weight="bold"))
            status.pack(pady=(0, 5))
            self.check_labels[key] = (values, status)
        chart = ctk.CTkFrame(dashboard)
        chart.grid(row=0, column=5, padx=(5, 0), sticky="nsew")
        ctk.CTkLabel(chart, text="Distribuição", font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(5, 0))
        self.chart = tk.Canvas(chart, width=170, height=120, bg="#2b2b2b", highlightthickness=0)
        self.chart.pack(expand=True)

        filters = ctk.CTkFrame(self.container, fg_color="transparent")
        filters.grid(row=5, column=0, padx=30, pady=(0, 7), sticky="ew")
        filters.grid_columnconfigure(2, weight=1)
        self.category_selector = ctk.CTkSegmentedButton(filters, values=["Fornecedores", "Adiantamentos"],
                                                         variable=self.category, command=lambda _: self._reset())
        self.category_selector.grid(row=0, column=0, padx=(0, 12))
        ctk.CTkLabel(filters, text="Buscar:").grid(row=0, column=1, padx=(0, 8))
        self.search_entry = ctk.CTkEntry(filters, textvariable=self.search_text, placeholder_text="Digite fornecedor ou subconta")
        self.search_entry.grid(row=0, column=2, sticky="ew")
        self.search_entry.bind("<KeyRelease>", lambda _: self._reset())
        self.preview = create_result_table(self.container, (
            TableColumn("name", "Fornecedor / Subconta", 410),
            TableColumn("accounting", "Contabilidade", 145, "e"),
            TableColumn("financial", "Financeiro", 145, "e"),
            TableColumn("difference", "Financeiro - Contabilidade", 190, "e"),
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
        names = filedialog.askopenfilenames(title="Selecionar os arquivos da Atividade 9", filetypes=[("Planilhas Excel", "*.xlsx")])
        if not names:
            return
        self.paths = [Path(name) for name in names]
        self._busy(True)
        self.app.set_status("Conferindo contas a pagar...", .1)
        self.app.run_background(lambda: analyze(self.paths), self._completed, self._failed)

    def _completed(self, result: PayablesResult) -> None:
        self.result = result
        self.page = 0
        self._busy(False)
        self._render_preview()
        self.app.set_status("Conferência concluída", 1)

    def _failed(self) -> None:
        self._busy(False)

    def _filtered(self) -> list[EntityRow]:
        if not self.result:
            return []
        search = normalize(self.search_text.get())
        return [row for row in self.result.entities if row.category == self.category.get()
                and (not search or search in normalize(row.name))]

    def _render_preview(self) -> None:
        clear_table(self.preview)
        if not self.result:
            for values, status in self.check_labels.values():
                values.configure(text="—")
                status.configure(text="Aguardando", text_color="gray70")
            self.page_label.configure(text="Página 0 de 0")
            self.previous_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
            self._draw_chart(0, 0)
        else:
            checks = (self.result.suppliers, self.result.advances, *self.result.taxes)
            for check in checks:
                values, status = self.check_labels[check.name]
                values.configure(text=f"Financeiro: {currency(check.financial)}\nContabilidade: {currency(check.accounting)}")
                status.configure(text=f"{check.status} • diferença {currency(check.difference)}",
                                 text_color="#21a67a" if check.status == "Conciliado" else "#dc5a5a")
            self.files_label.configure(text=f"Arquivos carregados: {len(self.paths)} • Registros: {len(self.result.entities)}")
            filtered = self._filtered()
            total_pages = max(1, (len(filtered) + self.page_size - 1) // self.page_size)
            self.page = min(self.page, total_pages - 1)
            page_rows = filtered[self.page * self.page_size:(self.page + 1) * self.page_size]
            for row in page_rows:
                self.preview.insert("", "end", values=(row.name, currency(row.accounting), currency(row.financial),
                    currency(row.difference), row.status), tags=(result_tag(row.status),))
            self.page_label.configure(text=f"Página {self.page + 1} de {total_pages} • {integer(len(filtered))} registros")
            self.previous_button.configure(state="normal" if self.page else "disabled")
            self.next_button.configure(state="normal" if self.page + 1 < total_pages else "disabled")
            reconciled = sum(row.status == "Conciliado" for row in self.result.entities)
            reconciled += sum(check.status == "Conciliado" for check in self.result.taxes)
            total = len(self.result.entities) + len(self.result.taxes)
            self._draw_chart(reconciled, total - reconciled)
        self.export_button.configure(state="normal" if self.result else "disabled")

    def _draw_chart(self, reconciled: int, divergent: int) -> None:
        self.chart.delete("all")
        total = reconciled + divergent
        if not total:
            self.chart.create_oval(51, 5, 119, 73, outline="#555555", width=12)
            self.chart.create_text(85, 39, text="—", fill="white", font=("Arial", 11, "bold"))
            return
        angle = reconciled / total * 360
        self.chart.create_arc(51, 5, 119, 73, start=90, extent=-angle, fill="#21a67a", outline="")
        self.chart.create_arc(51, 5, 119, 73, start=90 - angle, extent=-(360 - angle), fill="#dc5a5a", outline="")
        self.chart.create_oval(64, 18, 106, 60, fill="#2b2b2b", outline="")
        self.chart.create_text(85, 39, text=f"{reconciled / total:.1%}", fill="white", font=("Arial", 10, "bold"))
        self.chart.create_text(85, 88, text=f"{reconciled} conciliados", fill="#21a67a", font=("Arial", 9))
        self.chart.create_text(85, 104, text=f"{divergent} divergentes", fill="#dc5a5a", font=("Arial", 9))

    def _reset(self) -> None:
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
                                            initialfile=f"conferencia_contas_pagar{extension}",
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
        self.category_selector.configure(state=state)
        self.search_entry.configure(state=state)
        self.export_button.configure(state="disabled" if busy or not self.result else "normal")


AUTOMATION_CLASS = PayablesAutomation
