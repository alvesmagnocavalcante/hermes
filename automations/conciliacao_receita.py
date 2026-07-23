from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any
import warnings

from automations.legacy_ui import ctk, filedialog, messagebox, tk
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from automations.base import Automation
from automations.excel_reader import load_workbook_compatible as load_workbook
from automations.ui import TableColumn, clear_table, create_result_table, result_tag


TOLERANCE = Decimal("0.01")


@dataclass(frozen=True)
class ReconciliationRow:
    document: str
    cmflex: Decimal
    opera: Decimal

    @property
    def difference(self) -> Decimal:
        return self.cmflex - self.opera

    @property
    def status(self) -> str:
        return "Conciliado" if abs(self.difference) <= TOLERANCE else "Divergente"


@dataclass(frozen=True)
class ReconciliationResult:
    rows: list[ReconciliationRow]
    cmflex_total: Decimal
    opera_total: Decimal

    @property
    def difference(self) -> Decimal:
        return self.cmflex_total - self.opera_total

    @property
    def reconciled(self) -> int:
        return sum(row.status == "Conciliado" for row in self.rows)


def currency(value: Decimal) -> str:
    text = f"R$ {value:,.2f}"
    return text.replace(",", "_").replace(".", ",").replace("_", ".")


def integer(value: int) -> str:
    return f"{value:,}".replace(",", ".")


def decimal_value(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    try:
        return Decimal(str(value))
    except InvalidOperation:
        text = str(value).replace(".", "").replace(",", ".")
        return Decimal(text)


def read_values(path: Path) -> tuple[str, dict[str, Decimal]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook.active
        # Alguns relatórios exportados pelos sistemas não gravam a dimensão da
        # planilha. No modo somente leitura, calculate_dimension() lança
        # "Worksheet is unsized" antes que seja possível percorrer as linhas.
        sheet.reset_dimensions()
        rows = sheet.iter_rows(values_only=True)
        headers = tuple(next(rows))

        if "Movimento" in headers and "Documento" in headers:
            source, key_name, value_name, multiplier = "Contabilidade", "Documento", "Movimento", Decimal("-1")
        elif "CASHIER_DEBIT" in headers and "TRX_NO" in headers:
            source, key_name, value_name, multiplier = "Opera", "TRX_NO", "CASHIER_DEBIT", Decimal("1")
        else:
            raise ValueError(
                f"{path.name}: não foram encontradas as colunas esperadas "
                "(Movimento/Documento ou CASHIER_DEBIT/TRX_NO)."
            )

        key_index, value_index = headers.index(key_name), headers.index(value_name)
        values: defaultdict[str, Decimal] = defaultdict(Decimal)
        for row in rows:
            if key_index >= len(row) or value_index >= len(row):
                continue
            document = row[key_index]
            if document in (None, "", "NULL"):
                continue
            values[str(document).strip()] += decimal_value(row[value_index]) * multiplier
        return source, dict(values)
    finally:
        workbook.close()


def reconcile(paths: list[Path]) -> ReconciliationResult:
    if len(paths) != 2:
        raise ValueError("Selecione exatamente o arquivo da Contabilidade e o arquivo do Opera.")

    sources: dict[str, dict[str, Decimal]] = {}
    for path in paths:
        source, values = read_values(path)
        if source in sources:
            raise ValueError(f"Foram selecionados dois arquivos do tipo {source}.")
        sources[source] = values
    if set(sources) != {"Contabilidade", "Opera"}:
        raise ValueError("É necessário selecionar um arquivo da Contabilidade e um do Opera.")

    cmflex, opera = sources["Contabilidade"], sources["Opera"]
    rows = [
        ReconciliationRow(document, cmflex.get(document, Decimal()), opera.get(document, Decimal()))
        for document in sorted(set(cmflex) | set(opera))
    ]
    rows.sort(key=lambda row: (row.status == "Conciliado", -abs(row.difference), row.document))
    return ReconciliationResult(rows, sum(cmflex.values(), Decimal()), sum(opera.values(), Decimal()))


def save_excel_result(result: ReconciliationResult, path: Path) -> None:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumo"
    summary.append(["Indicador", "Valor"])
    summary.append(["Total Contabilidade (Movimento)", float(result.cmflex_total)])
    summary.append(["Total Opera (CASHIER_DEBIT)", float(result.opera_total)])
    summary.append(["Diferença", float(result.difference)])
    summary.append(["Transações conciliadas", result.reconciled])
    summary.append(["Transações divergentes", len(result.rows) - result.reconciled])
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 20
    for cell in summary[1]:
        cell.style = "Headline 4"
    for cell in summary["B"][1:4]:
        cell.number_format = 'R$ #,##0.00'

    details = workbook.create_sheet("Conciliação")
    details.append(["Documento / TRX_NO", "Movimento Contabilidade", "CASHIER_DEBIT Opera", "Diferença", "Status"])
    for row in result.rows:
        details.append([row.document, float(row.cmflex), float(row.opera), float(row.difference), row.status])
    for cell in details[1]:
        cell.style = "Headline 4"
    for column in ("B", "C", "D"):
        for cell in details[column][1:]:
            cell.number_format = 'R$ #,##0.00'
    for column, width in {"A": 24, "B": 22, "C": 24, "D": 18, "E": 16}.items():
        details.column_dimensions[column].width = width
    details.auto_filter.ref = details.dimensions
    details.freeze_panes = "A2"
    workbook.save(path)


def save_pdf_result(result: ReconciliationResult, path: Path) -> None:
    styles = getSampleStyleSheet()
    document = SimpleDocTemplate(
        str(path),
        pagesize=landscape(A4),
        leftMargin=10 * mm,
        rightMargin=10 * mm,
        topMargin=10 * mm,
        bottomMargin=10 * mm,
        title="Conciliação de Receita",
    )
    divergent = len(result.rows) - result.reconciled
    summary_data = [
        ["Total Contabilidade", "Total Opera", "Diferença", "Conciliadas", "Divergentes"],
        [
            currency(result.cmflex_total),
            currency(result.opera_total),
            currency(result.difference),
            integer(result.reconciled),
            integer(divergent),
        ],
    ]
    summary = Table(summary_data, colWidths=[48 * mm, 48 * mm, 48 * mm, 35 * mm, 35 * mm])
    summary.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#24588A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
    ]))

    document.build([
        Paragraph("Conciliação de Receita — Contabilidade x Opera", styles["Title"]),
        Spacer(1, 5 * mm),
        summary,
        Spacer(1, 5 * mm),
    ])


class RevenueReconciliationAutomation(Automation):
    name = "Conciliação de Receita"

    def __init__(self, app, container: ctk.CTkFrame) -> None:
        super().__init__(app, container)
        self.paths: list[Path] = []
        self.result: ReconciliationResult | None = None
        self.output_format = ctk.StringVar(value="Excel")
        self.search_text = ctk.StringVar()
        self.page = 0
        self.page_size = 100

    def render(self) -> None:
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(6, weight=1)
        ctk.CTkLabel(
            self.container,
            text=self.name,
            font=ctk.CTkFont(size=26, weight="bold"),
        ).grid(row=0, column=0, padx=30, pady=(22, 3), sticky="w")
        ctk.CTkLabel(
            self.container,
            text="Compara Movimento (Contabilidade) com CASHIER_DEBIT (Opera) por Documento/TRX_NO.",
            text_color="gray70",
        ).grid(row=1, column=0, padx=30, pady=(0, 10), sticky="w")

        controls = ctk.CTkFrame(self.container, fg_color="transparent")
        controls.grid(row=2, column=0, padx=30, sticky="ew")
        self.select_button = ctk.CTkButton(controls, text="Selecionar Contabilidade e Opera", command=self._select)
        self.select_button.pack(side="left", padx=(0, 10))
        self.format_selector = ctk.CTkSegmentedButton(
            controls,
            values=["Excel", "PDF"],
            variable=self.output_format,
        )
        self.format_selector.pack(side="left", padx=10)
        self.export_button = ctk.CTkButton(controls, text="Exportar resultado", state="disabled", command=self._export)
        self.export_button.pack(side="left", padx=10)
        self.clear_button = ctk.CTkButton(controls, text="Limpar", fg_color="gray35", command=self._clear)
        self.clear_button.pack(side="left", padx=10)

        self.files_label = ctk.CTkLabel(
            self.container,
            text="Selecione os dois arquivos.",
            anchor="w",
            text_color="gray70",
        )
        self.files_label.grid(row=3, column=0, padx=30, pady=(10, 6), sticky="ew")

        dashboard_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        dashboard_frame.grid(row=4, column=0, padx=30, pady=(0, 8), sticky="ew")
        dashboard_frame.grid_columnconfigure(0, weight=3)
        dashboard_frame.grid_columnconfigure(1, weight=2)

        summary_frame = ctk.CTkFrame(dashboard_frame, fg_color="transparent")
        summary_frame.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        self.summary_values: dict[str, ctk.CTkLabel] = {}
        cards = (
            ("Contabilidade", "cmflex"),
            ("Opera", "opera"),
            ("Diferença", "difference"),
            ("Conciliadas", "reconciled"),
            ("Divergentes", "divergent"),
            ("Índice", "percentage"),
        )
        for column, (title, key) in enumerate(cards):
            grid_column = column % 3
            grid_row = column // 3
            summary_frame.grid_columnconfigure(grid_column, weight=1)
            card = ctk.CTkFrame(summary_frame)
            card.grid(
                row=grid_row,
                column=grid_column,
                padx=(0 if grid_column == 0 else 5, 0),
                pady=(0 if grid_row == 0 else 5, 0),
                sticky="nsew",
            )
            ctk.CTkLabel(card, text=title, text_color="gray70").pack(padx=10, pady=(8, 0))
            value_label = ctk.CTkLabel(card, text="—", font=ctk.CTkFont(size=15, weight="bold"))
            value_label.pack(padx=10, pady=(0, 8))
            self.summary_values[key] = value_label

        chart_frame = ctk.CTkFrame(dashboard_frame)
        chart_frame.grid(row=0, column=1, sticky="nsew")
        chart_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            chart_frame,
            text="Distribuição da conciliação",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="center",
        ).grid(row=0, column=0, padx=14, pady=(8, 0), sticky="ew")
        self.chart = tk.Canvas(chart_frame, height=125, background="#2b2b2b", highlightthickness=0)
        self.chart.grid(row=1, column=0, padx=10, pady=(0, 8), sticky="ew")
        self.chart.bind("<Configure>", lambda _: self._draw_chart())

        search_frame = ctk.CTkFrame(self.container, fg_color="transparent")
        search_frame.grid(row=5, column=0, padx=30, pady=(0, 8), sticky="ew")
        search_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            search_frame,
            text="Buscar divergência:",
            anchor="w",
        ).grid(row=0, column=0, padx=(0, 10), sticky="w")
        self.search_entry = ctk.CTkEntry(
            search_frame,
            textvariable=self.search_text,
            placeholder_text="Digite o Documento ou TRX_NO",
        )
        self.search_entry.grid(row=0, column=1, sticky="ew")
        self.search_entry.bind("<KeyRelease>", lambda _: self._search())

        self.preview = create_result_table(self.container, (
            TableColumn("document", "Documento", 190), TableColumn("accounting", "Contabilidade", 155, "e"),
            TableColumn("opera", "Opera", 155, "e"), TableColumn("difference", "Diferença", 155, "e"),
            TableColumn("status", "Situação", 130),
        ), row=6)

        pagination = ctk.CTkFrame(self.container, fg_color="transparent")
        pagination.grid(row=7, column=0, padx=30, pady=(7, 14), sticky="ew")
        pagination.grid_columnconfigure(1, weight=1)
        self.previous_button = ctk.CTkButton(pagination, text="Anterior", width=100, command=self._previous_page)
        self.previous_button.grid(row=0, column=0)
        self.page_label = ctk.CTkLabel(pagination, text="Página 0 de 0")
        self.page_label.grid(row=0, column=1)
        self.next_button = ctk.CTkButton(pagination, text="Próxima", width=100, command=self._next_page)
        self.next_button.grid(row=0, column=2)
        self._render_preview()

    def _select(self) -> None:
        names = filedialog.askopenfilenames(
            title="Selecionar os arquivos Contabilidade e Opera",
            filetypes=[("Planilhas Excel", "*.xlsx *.xlsm *.xls *.xltx *.xltm")],
        )
        if not names:
            return
        self.paths = [Path(name) for name in names]
        self._busy(True)
        self.app.set_status("Conciliando arquivos...", 0.1)
        self.app.run_background(lambda: reconcile(self.paths), self._completed, self._failed)

    def _completed(self, result: ReconciliationResult) -> None:
        self.result = result
        self.page = 0
        self._busy(False)
        self._render_preview()
        self.app.set_status("Conciliação concluída", 1)

    def _failed(self) -> None:
        self._busy(False)

    def _render_preview(self) -> None:
        clear_table(self.preview)
        if not self.result:
            self.files_label.configure(text="Selecione os dois arquivos.")
            for label in self.summary_values.values():
                label.configure(text="—")
            self.page_label.configure(text="Página 0 de 0")
            self.previous_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
        else:
            result = self.result
            divergent = len(result.rows) - result.reconciled
            percentage = Decimal(result.reconciled * 100) / Decimal(len(result.rows)) if result.rows else Decimal(100)
            self.files_label.configure(
                text=f"Arquivos carregados: 2 • Contabilidade e Opera • {integer(len(result.rows))} transações"
            )
            values = {
                "cmflex": currency(result.cmflex_total),
                "opera": currency(result.opera_total),
                "difference": currency(result.difference),
                "reconciled": integer(result.reconciled),
                "divergent": integer(divergent),
                "percentage": f"{percentage:.2f}%".replace(".", ","),
            }
            for key, value in values.items():
                self.summary_values[key].configure(text=value)
            filtered = self._filtered_rows()
            total_pages = max(1, (len(filtered) + self.page_size - 1) // self.page_size)
            self.page = min(self.page, total_pages - 1)
            start = self.page * self.page_size
            page_rows = filtered[start : start + self.page_size]
            for row in page_rows:
                self.preview.insert("", "end", values=(row.document, currency(row.cmflex), currency(row.opera),
                    currency(row.difference), row.status), tags=(result_tag(row.status),))
            self.page_label.configure(
                text=f"Página {self.page + 1} de {total_pages} • {integer(len(filtered))} divergências"
            )
            self.previous_button.configure(state="normal" if self.page > 0 else "disabled")
            self.next_button.configure(state="normal" if self.page + 1 < total_pages else "disabled")
        self.export_button.configure(state="normal" if self.result else "disabled")
        self.app.after_idle(self._draw_chart)

    def _draw_chart(self) -> None:
        self.chart.delete("all")
        width = max(self.chart.winfo_width(), 300)
        center_x, center_y, radius = width / 2, 62, 50
        bounds = (center_x - radius, center_y - radius, center_x + radius, center_y + radius)
        if not self.result or not self.result.rows:
            self.chart.create_oval(*bounds, outline="#4b5563", width=14)
            self.chart.create_text(center_x, center_y, text="—", fill="#e5e7eb", font=("Segoe UI", 15, "bold"))
            self.chart.create_text(center_x, 130, text="Aguardando conciliação", fill="#cbd5e1")
            return

        reconciled_rate = self.result.reconciled / len(self.result.rows)
        divergent_rate = 1 - reconciled_rate
        self.chart.create_arc(
            *bounds, start=90, extent=-360 * reconciled_rate, style="arc", outline="#21a67a", width=14
        )
        self.chart.create_arc(
            *bounds,
            start=90 - 360 * reconciled_rate,
            extent=-360 * divergent_rate,
            style="arc",
            outline="#dc5a5a",
            width=14,
        )
        self.chart.create_text(
            center_x, center_y, text=f"{reconciled_rate:.1%}", fill="#e5e7eb", font=("Segoe UI", 11, "bold")
        )
        self.chart.create_text(
            center_x - 8, 130, text=f"● Conciliadas: {reconciled_rate:.2%}", fill="#21a67a", anchor="e"
        )
        self.chart.create_text(
            center_x + 8, 130, text=f"● Divergentes: {divergent_rate:.2%}", fill="#dc5a5a", anchor="w"
        )

    def _filtered_rows(self) -> list[ReconciliationRow]:
        if not self.result:
            return []
        search = self.search_text.get().strip().lower()
        return [
            row
            for row in self.result.rows
            if row.status == "Divergente" and (not search or search in row.document.lower())
        ]

    def _search(self) -> None:
        self.page = 0
        self._render_preview()

    def _previous_page(self) -> None:
        if self.page > 0:
            self.page -= 1
            self._render_preview()

    def _next_page(self) -> None:
        total = len(self._filtered_rows())
        if (self.page + 1) * self.page_size < total:
            self.page += 1
            self._render_preview()

    def _export(self) -> None:
        if not self.result:
            return
        output_format = self.output_format.get()
        extension = ".pdf" if output_format == "PDF" else ".xlsx"
        name = filedialog.asksaveasfilename(
            title="Salvar conciliação",
            defaultextension=extension,
            initialfile=f"conciliacao_receita_contabilidade_opera{extension}",
            filetypes=[(f"Arquivo {output_format}", f"*{extension}")],
        )
        if not name:
            return
        self._busy(True)
        if output_format == "PDF":
            task = lambda: save_pdf_result(self.result, Path(name))
        else:
            task = lambda: save_excel_result(self.result, Path(name))
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
        self.search_entry.configure(state=state)
        self.export_button.configure(state="disabled" if busy or not self.result else "normal")


AUTOMATION_CLASS = RevenueReconciliationAutomation
