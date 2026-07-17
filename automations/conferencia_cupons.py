from __future__ import annotations

import tkinter as tk
import warnings
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
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
class SourceEntry:
    value: Decimal
    emission_date: str


@dataclass(frozen=True)
class CouponRow:
    key: str
    document_type: str
    simphony_date: str
    fiscal_date: str
    sefaz_date: str
    simphony: Decimal | None
    fiscal: Decimal | None
    sefaz: Decimal | None
    status: str

    @property
    def difference(self) -> Decimal:
        values = [value for value in (self.simphony, self.fiscal, self.sefaz) if value is not None]
        return max(values) - min(values) if values else Decimal()

    @property
    def comparable(self) -> bool:
        return None not in (self.simphony, self.fiscal, self.sefaz)

    @property
    def reference_date(self) -> str:
        return next(
            (value for value in (self.simphony_date, self.fiscal_date, self.sefaz_date) if value != "—"),
            "—",
        )

@dataclass(frozen=True)
class CouponResult:
    rows: list[CouponRow]
    simphony_total: Decimal
    fiscal_total: Decimal
    sefaz_total: Decimal
    cancelled: int

    def count(self, status: str) -> int:
        return sum(row.status.startswith(status) for row in self.rows)


def decimal_value(value: Any) -> Decimal:
    if value in (None, ""):
        return Decimal()
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return Decimal(str(value).replace(".", "").replace(",", "."))


def currency(value: Decimal | None) -> str:
    if value is None:
        return "—"
    text = f"R$ {value:,.2f}"
    return text.replace(",", "_").replace(".", ",").replace("_", ".")


def integer(value: int) -> str:
    return f"{value:,}".replace(",", ".")


def date_value(value: Any) -> str:
    if value in (None, "", "NULL"):
        return "—"
    if isinstance(value, (date, datetime)):
        return value.strftime("%d/%m/%Y")
    text = str(value).strip()
    for pattern in ("%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, pattern).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return text


def document_type(key: str) -> str:
    model = key[20:22] if len(key) >= 22 else ""
    return "Cupom (NFC-e)" if model == "65" else "Nota (NF-e)" if model == "55" else "Nota (Unknown)"


def workbook_rows(path: Path) -> list[tuple[Any, ...]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook.active
        sheet.reset_dimensions()
        return list(sheet.iter_rows(values_only=True))
    finally:
        workbook.close()


def find_header(rows: list[tuple[Any, ...]], required: set[str]) -> tuple[int, tuple[Any, ...]]:
    for index, row in enumerate(rows[:40]):
        if required.issubset(set(row)):
            return index, row
    raise ValueError(f"Cabeçalhos não encontrados: {', '.join(sorted(required))}.")


def read_file(path: Path) -> tuple[str, dict[str, SourceEntry], int]:
    rows = workbook_rows(path)
    headers = set(value for row in rows[:40] for value in row if value)
    values: defaultdict[str, Decimal] = defaultdict(Decimal)
    dates: dict[str, str] = {}
    cancelled = 0

    if {"Chave da NF", "Valor Total NF", "Status"}.issubset(headers):
        source = "Simphony"
        header_i, header = find_header(rows, {"Chave da NF", "Valor Total NF", "Status"})
        key_i, value_i, status_i = header.index("Chave da NF"), header.index("Valor Total NF"), header.index("Status")
        date_i = header.index("Data")
        for row in rows[header_i + 1:]:
            if len(row) <= max(key_i, value_i, status_i) or not row[key_i]:
                continue
            if str(row[status_i]).strip().lower() != "aprovado":
                cancelled += 1
                continue
            key = str(row[key_i]).strip()
            values[key] += decimal_value(row[value_i])
            dates.setdefault(key, date_value(row[date_i] if len(row) > date_i else None))
    elif {"Chave", "ValorContabil"}.issubset(headers):
        source = "Fiscal"
        header_i, header = find_header(rows, {"Chave", "ValorContabil"})
        key_i, value_i = header.index("Chave"), header.index("ValorContabil")
        date_i = header.index("DataDocumento")
        cancelled_i = header.index("Cancelado") if "Cancelado" in header else None
        for row in rows[header_i + 1:]:
            if len(row) <= max(key_i, value_i) or not row[key_i]:
                continue
            if cancelled_i is not None and len(row) > cancelled_i and str(row[cancelled_i]).lower() == "true":
                continue
            key = str(row[key_i]).strip()
            values[key] += decimal_value(row[value_i])
            dates.setdefault(key, date_value(row[date_i] if len(row) > date_i else None))
    elif {"Chave de acesso", "Valor R$"}.issubset(headers):
        source = "SEFAZ"
        header_i, header = find_header(rows, {"Chave de acesso", "Valor R$"})
        key_i, value_i = header.index("Chave de acesso"), header.index("Valor R$")
        date_i = header.index("Data de emissão")
        for row in rows[header_i + 1:]:
            if len(row) > max(key_i, value_i) and row[key_i]:
                key = str(row[key_i]).strip()
                values[key] += decimal_value(row[value_i])
                dates.setdefault(key, date_value(row[date_i] if len(row) > date_i else None))
    else:
        raise ValueError(f"{path.name}: formato não reconhecido como Simphony, Fiscal ou SEFAZ.")
    return source, {key: SourceEntry(value, dates.get(key, "—")) for key, value in values.items()}, cancelled


def reconcile(paths: list[Path]) -> CouponResult:
    if len(paths) != 3:
        raise ValueError("Selecione exatamente as planilhas Simphony, Fiscal e SEFAZ.")
    sources: dict[str, dict[str, SourceEntry]] = {}
    cancelled = 0
    for path in paths:
        source, values, file_cancelled = read_file(path)
        if source in sources:
            raise ValueError(f"Foram selecionados dois arquivos do tipo {source}.")
        sources[source] = values
        cancelled += file_cancelled
    if set(sources) != {"Simphony", "Fiscal", "SEFAZ"}:
        raise ValueError("É necessário selecionar um arquivo Simphony, um Fiscal e um SEFAZ.")

    simphony, fiscal, sefaz = sources["Simphony"], sources["Fiscal"], sources["SEFAZ"]
    result_rows: list[CouponRow] = []
    for key in set(simphony) | set(fiscal) | set(sefaz):
        entries = (simphony.get(key), fiscal.get(key), sefaz.get(key))
        values = tuple(entry.value if entry else None for entry in entries)
        missing = [name for name, value in zip(("Simphony", "Fiscal", "SEFAZ"), entries) if value is None]
        if missing:
            status = "Ausente: " + "/".join(missing)
        elif max(values) - min(values) <= TOLERANCE:  # type: ignore[arg-type]
            status = "Conciliado"
        else:
            status = "Divergente: valor"
        source_dates = tuple(entry.emission_date if entry else "—" for entry in entries)
        result_rows.append(CouponRow(key, document_type(key), *source_dates, *values, status))
    result_rows.sort(key=lambda row: (row.status == "Conciliado", -row.difference, row.key))
    return CouponResult(
        result_rows, sum((entry.value for entry in simphony.values()), Decimal()),
        sum((entry.value for entry in fiscal.values()), Decimal()),
        sum((entry.value for entry in sefaz.values()), Decimal()), cancelled,
    )


def save_excel(result: CouponResult, path: Path) -> None:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumo"
    missing = sum(row.status.startswith("Ausente") for row in result.rows)
    summary.append(["Indicador", "Valor"])
    for label, value in (
        ("Total Simphony", float(result.simphony_total)), ("Total Fiscal", float(result.fiscal_total)),
        ("Total SEFAZ", float(result.sefaz_total)), ("Chaves analisadas", len(result.rows)),
        ("Conciliadas", result.count("Conciliado")), ("Divergentes", result.count("Divergente")),
        ("Com integração ausente", missing), ("Cupons cancelados ignorados", result.cancelled),
    ):
        summary.append([label, value])
    for cell in summary[1]:
        cell.style = "Headline 4"
    for cell in summary["B"][1:4]:
        cell.number_format = 'R$ #,##0.00'
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 20

    details = workbook.create_sheet("Conciliação")
    details.append(["Tipo", "Chave fiscal", "Data", "Simphony - Valor Total NF", "Fiscal - Valor Contábil",
                    "SEFAZ - Valor", "Diferença", "Status"])
    for row in result.rows:
        details.append([
            row.document_type, row.key, row.reference_date,
            float(row.simphony) if row.simphony is not None else None,
            float(row.fiscal) if row.fiscal is not None else None,
            float(row.sefaz) if row.sefaz is not None else None,
            float(row.difference) if row.comparable else None, row.status,
        ])
    for cell in details[1]:
        cell.style = "Headline 4"
    for column in ("D", "E", "F", "G"):
        for cell in details[column][1:]:
            cell.number_format = 'R$ #,##0.00'
    for column, width in {"A": 18, "B": 48, "C": 14, "D": 25, "E": 25, "F": 22, "G": 18, "H": 24}.items():
        details.column_dimensions[column].width = width
    details.freeze_panes = "A2"
    details.auto_filter.ref = details.dimensions
    workbook.save(path)


def save_pdf(result: CouponResult, path: Path) -> None:
    styles = getSampleStyleSheet()
    document = SimpleDocTemplate(
        str(path), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm, title="Resumo da conferência dos cupons",
    )
    missing = sum(row.status.startswith("Ausente") for row in result.rows)
    data = [
        ["Simphony", "Fiscal", "SEFAZ", "Conciliadas", "Divergentes", "Ausentes"],
        [currency(result.simphony_total), currency(result.fiscal_total), currency(result.sefaz_total),
         integer(result.count("Conciliado")), integer(result.count("Divergente")), integer(missing)],
    ]
    table = Table(data, colWidths=[39 * mm] * 6)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#24588A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    document.build([
        Paragraph("Conferência dos cupons — Simphony x Fiscal x SEFAZ", styles["Title"]),
        Spacer(1, 6 * mm), table, Spacer(1, 6 * mm),
        Paragraph(f"Cupons cancelados ignorados: {integer(result.cancelled)}.", styles["BodyText"]),
    ])


class CouponReconciliationAutomation(Automation):
    name = "Conferência dos Cupons"

    def __init__(self, app, container: ctk.CTkFrame) -> None:
        super().__init__(app, container)
        self.paths: list[Path] = []
        self.result: CouponResult | None = None
        self.output_format = ctk.StringVar(value="Excel")
        self.filter_status = ctk.StringVar(value="Pendências")
        self.filter_type = ctk.StringVar(value="Todos os tipos")
        self.search_text = ctk.StringVar()
        self.page = 0
        self.page_size = 100

    def render(self) -> None:
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(6, weight=1)
        ctk.CTkLabel(self.container, text=self.name, font=ctk.CTkFont(size=26, weight="bold")).grid(
            row=0, column=0, padx=30, pady=(22, 3), sticky="w"
        )
        ctk.CTkLabel(
            self.container,
            text="Compara cupons NFC-e e notas NF-e/Unknown pela chave e pelos valores do Simphony, Fiscal e SEFAZ.",
            text_color="gray70",
        ).grid(row=1, column=0, padx=30, pady=(0, 10), sticky="w")
        controls = ctk.CTkFrame(self.container, fg_color="transparent")
        controls.grid(row=2, column=0, padx=30, sticky="ew")
        self.select_button = ctk.CTkButton(controls, text="Selecionar os três arquivos", command=self._select)
        self.select_button.pack(side="left", padx=(0, 10))
        self.format_selector = ctk.CTkSegmentedButton(controls, values=["Excel", "PDF"], variable=self.output_format)
        self.format_selector.pack(side="left", padx=10)
        self.export_button = ctk.CTkButton(controls, text="Exportar resultado", state="disabled", command=self._export)
        self.export_button.pack(side="left", padx=10)
        self.clear_button = ctk.CTkButton(controls, text="Limpar", fg_color="gray35", command=self._clear)
        self.clear_button.pack(side="left", padx=10)
        self.files_label = ctk.CTkLabel(self.container, text="Selecione os três arquivos.", text_color="gray70", anchor="w")
        self.files_label.grid(row=3, column=0, padx=30, pady=(10, 6), sticky="ew")

        dashboard = ctk.CTkFrame(self.container, fg_color="transparent")
        dashboard.grid(row=4, column=0, padx=30, pady=(0, 8), sticky="ew")
        dashboard.grid_columnconfigure(0, weight=3)
        dashboard.grid_columnconfigure(1, weight=2)
        summary = ctk.CTkFrame(dashboard, fg_color="transparent")
        summary.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        self.summary_labels: dict[str, ctk.CTkLabel] = {}
        cards = (("Simphony", "simphony"), ("Fiscal", "fiscal"), ("SEFAZ", "sefaz"),
                 ("Conciliadas", "ok"), ("Divergentes", "different"), ("Ausentes", "missing"))
        for index, (title, key) in enumerate(cards):
            column, row = index % 3, index // 3
            summary.grid_columnconfigure(column, weight=1)
            card = ctk.CTkFrame(summary)
            card.grid(row=row, column=column, padx=(0 if column == 0 else 5, 0), pady=(0 if row == 0 else 5, 0), sticky="nsew")
            ctk.CTkLabel(card, text=title, text_color="gray70").pack(pady=(6, 0))
            label = ctk.CTkLabel(card, text="—", font=ctk.CTkFont(size=15, weight="bold"))
            label.pack(pady=(0, 6))
            self.summary_labels[key] = label
        chart_frame = ctk.CTkFrame(dashboard)
        chart_frame.grid(row=0, column=1, sticky="nsew")
        chart_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(chart_frame, text="Distribuição da conferência", font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, pady=(6, 0), sticky="ew"
        )
        self.chart = tk.Canvas(chart_frame, height=125, background="#2b2b2b", highlightthickness=0)
        self.chart.grid(row=1, column=0, padx=8, sticky="ew")
        self.chart.bind("<Configure>", lambda _: self._draw_chart())

        filters = ctk.CTkFrame(self.container, fg_color="transparent")
        filters.grid(row=5, column=0, padx=30, pady=(0, 8), sticky="ew")
        filters.grid_columnconfigure(3, weight=1)
        self.status_selector = ctk.CTkSegmentedButton(
            filters, values=["Pendências", "Conciliadas", "Todas"], variable=self.filter_status,
            command=lambda _: self._reset_filter(),
        )
        self.status_selector.grid(row=0, column=0, padx=(0, 12))
        self.type_selector = ctk.CTkOptionMenu(
            filters, values=["Todos os tipos", "Cupons NFC-e", "Notas NF-e/Unknown"],
            variable=self.filter_type, command=lambda _: self._reset_filter(), width=185,
        )
        self.type_selector.grid(row=0, column=1, padx=(0, 12))
        ctk.CTkLabel(filters, text="Buscar chave:").grid(row=0, column=2, padx=(0, 8))
        self.search_entry = ctk.CTkEntry(filters, textvariable=self.search_text, placeholder_text="Digite a chave fiscal")
        self.search_entry.grid(row=0, column=3, sticky="ew")
        self.search_entry.bind("<KeyRelease>", lambda _: self._reset_filter())
        self.preview = create_result_table(self.container, (
            TableColumn("type", "Tipo", 130), TableColumn("date", "Data", 90),
            TableColumn("key", "Chave fiscal", 330), TableColumn("simphony", "Simphony", 115, "e"),
            TableColumn("fiscal", "Fiscal", 115, "e"), TableColumn("sefaz", "SEFAZ", 115, "e"),
            TableColumn("difference", "Diferença", 115, "e"), TableColumn("status", "Situação", 210),
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
        names = filedialog.askopenfilenames(title="Selecionar Simphony, Fiscal e SEFAZ", filetypes=[("Planilhas Excel", "*.xlsx")])
        if not names:
            return
        self.paths = [Path(name) for name in names]
        self._busy(True)
        self.app.set_status("Conferindo cupons...", 0.1)
        self.app.run_background(lambda: reconcile(self.paths), self._completed, self._failed)

    def _completed(self, result: CouponResult) -> None:
        self.result = result
        self.page = 0
        self._busy(False)
        self._render_preview()
        self.app.set_status("Conferência concluída", 1)

    def _failed(self) -> None:
        self._busy(False)

    def _filtered(self) -> list[CouponRow]:
        if not self.result:
            return []
        selected, selected_type, search = self.filter_status.get(), self.filter_type.get(), self.search_text.get().strip()
        return [row for row in self.result.rows if
                (selected == "Todas" or (selected == "Conciliadas" and row.status == "Conciliado")
                 or (selected == "Pendências" and row.status != "Conciliado"))
                and (selected_type == "Todos os tipos"
                     or (selected_type == "Cupons NFC-e" and row.document_type == "Cupom (NFC-e)")
                     or (selected_type == "Notas NF-e/Unknown" and row.document_type != "Cupom (NFC-e)"))
                and (not search or search in row.key)]

    def _render_preview(self) -> None:
        clear_table(self.preview)
        if not self.result:
            self.files_label.configure(text="Selecione os três arquivos.")
            for label in self.summary_labels.values():
                label.configure(text="—")
            self.page_label.configure(text="Página 0 de 0")
            self.previous_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
        else:
            result = self.result
            missing = sum(row.status.startswith("Ausente") for row in result.rows)
            values = {"simphony": currency(result.simphony_total), "fiscal": currency(result.fiscal_total),
                      "sefaz": currency(result.sefaz_total), "ok": integer(result.count("Conciliado")),
                      "different": integer(result.count("Divergente")), "missing": integer(missing)}
            for key, value in values.items():
                self.summary_labels[key].configure(text=value)
            self.files_label.configure(
                text=(
                    "Arquivos carregados: 3 • Fiscal, SEFAZ e Simphony\n"
                    f"Observação: {integer(result.cancelled)} cupons cancelados foram desconsiderados da conferência."
                ),
                justify="left",
            )
            filtered = self._filtered()
            total_pages = max(1, (len(filtered) + self.page_size - 1) // self.page_size)
            self.page = min(self.page, total_pages - 1)
            page_rows = filtered[self.page * self.page_size:(self.page + 1) * self.page_size]
            for row in page_rows:
                self.preview.insert("", "end", values=(row.document_type, row.reference_date, row.key,
                    currency(row.simphony), currency(row.fiscal), currency(row.sefaz),
                    currency(row.difference) if row.comparable else "", row.status), tags=(result_tag(row.status),))
            self.page_label.configure(text=f"Página {self.page + 1} de {total_pages} • {integer(len(filtered))} registros")
            self.previous_button.configure(state="normal" if self.page else "disabled")
            self.next_button.configure(state="normal" if self.page + 1 < total_pages else "disabled")
        self.export_button.configure(state="normal" if self.result else "disabled")
        self.app.after_idle(self._draw_chart)

    def _draw_chart(self) -> None:
        self.chart.delete("all")
        width = max(self.chart.winfo_width(), 300)
        cx, cy, radius = width * .33, 55, 45
        bounds = (cx - radius, cy - radius, cx + radius, cy + radius)
        if not self.result or not self.result.rows:
            self.chart.create_oval(*bounds, outline="#4b5563", width=13)
            self.chart.create_text(cx, cy, text="—", fill="#e5e7eb")
            return
        total = len(self.result.rows)
        segments = ((self.result.count("Conciliado"), "#21a67a", "Conciliadas"),
                    (self.result.count("Divergente"), "#dc5a5a", "Divergentes"),
                    (sum(row.status.startswith("Ausente") for row in self.result.rows), "#e0a83e", "Ausentes"))
        start, legend_y = 90.0, 26
        for count, color, label in segments:
            if not count:
                continue
            rate = count / total
            extent = -360 * rate
            self.chart.create_arc(*bounds, start=start, extent=extent, style="arc", outline=color, width=13)
            self.chart.create_text(width * .57, legend_y, text=f"● {label}: {rate:.1%}", fill=color, anchor="w")
            start += extent
            legend_y += 24
        self.chart.create_text(cx, cy, text=f"{self.result.count('Conciliado') / total:.1%}", fill="#e5e7eb", font=("Segoe UI", 10, "bold"))

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
        name = filedialog.asksaveasfilename(
            title="Salvar conferência", defaultextension=extension,
            initialfile=f"conferencia_cupons{extension}", filetypes=[(f"Arquivo {output_format}", f"*{extension}")],
        )
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
        self.filter_type.set("Todos os tipos")
        self.page = 0
        self._render_preview()
        self.app.set_status("Seleção limpa", 0)

    def _busy(self, busy: bool) -> None:
        state = "disabled" if busy else "normal"
        self.select_button.configure(state=state)
        self.clear_button.configure(state=state)
        self.format_selector.configure(state=state)
        self.status_selector.configure(state=state)
        self.type_selector.configure(state=state)
        self.search_entry.configure(state=state)
        self.export_button.configure(state="disabled" if busy or not self.result else "normal")


AUTOMATION_CLASS = CouponReconciliationAutomation
