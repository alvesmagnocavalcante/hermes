from __future__ import annotations

import re
import unicodedata
import warnings
from dataclasses import dataclass
from datetime import date, datetime
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


STATE_CODES = {
    "ceara": "CE", "sao-paulo": "SP", "parana": "PR", "distrito-federal": "DF",
    "santa-catarina": "SC", "espirito-santo": "ES", "alagoas": "AL", "pernambuco": "PE",
    "paraiba": "PB", "bahia": "BA", "rio-grande-do-sul": "RS", "minas-gerais": "MG",
    "rio-de-janeiro": "RJ", "piaui": "PI", "amazonas": "AM", "pa": "PA",
    "rio-grande-do-norte": "RN",
}


@dataclass(frozen=True)
class NoteResult:
    key: str
    company: str
    supplier: str
    state: str
    emission_date: date | None
    entry_date: date | None
    days: int | None
    limit: int | None
    launch_status: str
    status: str


@dataclass(frozen=True)
class AnalysisResult:
    rows: list[NoteResult]

    def count(self, status: str) -> int:
        return sum(row.status == status for row in self.rows)


def state_code(value: Any) -> str:
    text = str(value or "").lower()
    match = re.search(r"bandeira-([a-z-]+)\.png", text)
    slug = match.group(1) if match else text.strip().lower()
    return STATE_CODES.get(slug, slug.upper() if len(slug) == 2 else "N/I")


def as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for pattern in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(text[:10], pattern).date()
        except ValueError:
            continue
    return None


def normalized_header(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return re.sub(r"[^a-z0-9]", "", "".join(
        character for character in text.lower() if not unicodedata.combining(character)
    ))


def header_index(headers: tuple[Any, ...], *names: str) -> int:
    available = {normalized_header(header): index for index, header in enumerate(headers)}
    for name in names:
        index = available.get(normalized_header(name))
        if index is not None:
            return index
    raise ValueError(f"Coluna não encontrada: {' ou '.join(names)}.")


def note_status(days: int, state: str) -> tuple[int, str]:
    alert, limit = (6, 11) if state == "CE" else (20, 30)
    if days >= limit:
        return limit, "Em atraso"
    if days >= alert:
        return limit, "Alerta"
    return limit, "Em dia"


def open_rows(path: Path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    sheet.reset_dimensions()
    rows = sheet.iter_rows(values_only=True)
    headers = tuple(next(rows))
    return workbook, headers, rows


def analyze(paths: list[Path], reference_date: date | None = None) -> AnalysisResult:
    if len(paths) != 2:
        raise ValueError("Selecione exatamente as planilhas Manifesto e Detalhe das notas recebidas.")

    manifesto_path: Path | None = None
    detalhe_path: Path | None = None
    for path in paths:
        workbook, headers, _ = open_rows(path)
        workbook.close()
        normalized = {normalized_header(header) for header in headers}
        if normalized_header("Chave Manifesto") in normalized:
            manifesto_path = path
        elif {normalized_header("Chave"), normalized_header("Data de Entrada")} <= normalized:
            detalhe_path = path
    if not manifesto_path or not detalhe_path:
        raise ValueError("Não foi possível identificar uma planilha Manifesto e uma planilha Detalhe.")

    workbook, headers, rows = open_rows(manifesto_path)
    key_i = header_index(headers, "Chave Manifesto")
    emission_i = header_index(headers, "Data Emissão", "Data da Emissão")
    state_i = header_index(headers, "Estado")
    company_i = header_index(headers, "Empresa")
    supplier_i = header_index(headers, "Fornecedor", "Razão Social")
    max_i = max(key_i, emission_i, state_i, company_i, supplier_i)
    manifest: dict[str, tuple[date | None, str, str, str]] = {}
    try:
        for row in rows:
            if len(row) <= max_i or not row[key_i]:
                continue
            manifest[str(row[key_i]).strip()] = (
                as_date(row[emission_i]),
                state_code(row[state_i]),
                str(row[company_i] or ""),
                str(row[supplier_i] or ""),
            )
    finally:
        workbook.close()

    workbook, headers, rows = open_rows(detalhe_path)
    key_i = header_index(headers, "Chave")
    entry_i = header_index(headers, "Data de Entrada")
    max_i = max(key_i, entry_i)
    entries: dict[str, date | None] = {}
    try:
        for row in rows:
            if len(row) <= max_i or not row[key_i]:
                continue
            key = str(row[key_i]).strip()
            entry = as_date(row[entry_i])
            previous = entries.get(key)
            if previous is None or (entry is not None and entry < previous):
                entries[key] = entry
    finally:
        workbook.close()

    today = reference_date or date.today()
    results: list[NoteResult] = []
    for key, (emission, state, company, supplier) in manifest.items():
        entry = entries.get(key)
        launch_status = "Lançada" if entry else "Não lançada"
        if not emission:
            results.append(NoteResult(
                key, company, supplier, state, None, entry, None, None,
                f"{launch_status} • emissão ausente", "Alerta",
            ))
            continue
        reference = entry or today
        days = max(0, (reference - emission).days)
        limit, status = note_status(days, state)
        results.append(NoteResult(
            key, company, supplier, state, emission, entry, days, limit, launch_status, status,
        ))
    order = {"Em atraso": 0, "Alerta": 1, "Em dia": 2}
    results.sort(key=lambda item: (order[item.status], -(item.days or 0), item.key))
    return AnalysisResult(results)


def date_text(value: date | None) -> str:
    return value.strftime("%d/%m/%Y") if value else "—"


def save_excel(result: AnalysisResult, path: Path) -> None:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumo"
    summary.append(["Indicador", "Quantidade"])
    summary.append(["Notas analisadas", len(result.rows)])
    summary.append(["Em dia", result.count("Em dia")])
    summary.append(["Em alerta", result.count("Alerta")])
    summary.append(["Em atraso", result.count("Em atraso")])
    summary.append(["Não lançadas", sum(row.entry_date is None for row in result.rows)])
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 16
    for cell in summary[1]:
        cell.style = "Headline 4"

    details = workbook.create_sheet("Análise")
    details.append([
        "Chave", "Empresa", "Fornecedor", "UF", "Emissão", "Entrada", "Dias",
        "Limite para atraso", "Lançamento", "Situação",
    ])
    for row in result.rows:
        details.append([
            row.key, row.company, row.supplier, row.state, row.emission_date, row.entry_date,
            row.days, row.limit, row.launch_status, row.status,
        ])
    for cell in details[1]:
        cell.style = "Headline 4"
    for column in ("E", "F"):
        for cell in details[column][1:]:
            cell.number_format = "DD/MM/YYYY"
    for column, width in {
        "A": 48, "B": 22, "C": 45, "D": 8, "E": 14, "F": 14, "G": 10,
        "H": 18, "I": 24, "J": 16,
    }.items():
        details.column_dimensions[column].width = width
    details.freeze_panes = "A2"
    details.auto_filter.ref = details.dimensions
    workbook.save(path)


def save_pdf(result: AnalysisResult, path: Path) -> None:
    styles = getSampleStyleSheet()
    document = SimpleDocTemplate(
        str(path), pagesize=landscape(A4), leftMargin=12 * mm, rightMargin=12 * mm,
        topMargin=12 * mm, bottomMargin=12 * mm, title="Resumo de notas fiscais em atraso",
    )
    data = [
        ["Analisadas", "Em dia", "Em alerta", "Em atraso", "Não lançadas"],
        [
            len(result.rows), result.count("Em dia"), result.count("Alerta"),
            result.count("Em atraso"), sum(row.entry_date is None for row in result.rows),
        ],
    ]
    table = Table(data, colWidths=[45 * mm] * 5)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#24588A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#CBD5E1")),
        ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    document.build([
        Paragraph("Notas fiscais de entrada de mercadoria em atraso", styles["Title"]),
        Spacer(1, 6 * mm), table,
    ])


class OverdueNotesAutomation(Automation):
    name = "Notas Fiscais de Entrada em Atraso"

    def __init__(self, app, container: ctk.CTkFrame) -> None:
        super().__init__(app, container)
        self.paths: list[Path] = []
        self.result: AnalysisResult | None = None
        self.output_format = ctk.StringVar(value="Excel")
        self.filter_status = ctk.StringVar(value="Em atraso")
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
            text=("Usa o Manifesto como base e classifica notas lançadas ou não lançadas "
                  "como em dia, em alerta ou em atraso."),
            text_color="gray70",
        ).grid(row=1, column=0, padx=30, pady=(0, 10), sticky="w")

        controls = ctk.CTkFrame(self.container, fg_color="transparent")
        controls.grid(row=2, column=0, padx=30, sticky="ew")
        self.select_button = ctk.CTkButton(controls, text="Selecionar Manifesto e Detalhe", command=self._select)
        self.select_button.pack(side="left", padx=(0, 10))
        self.format_selector = ctk.CTkSegmentedButton(controls, values=["Excel", "PDF"], variable=self.output_format)
        self.format_selector.pack(side="left", padx=10)
        self.export_button = ctk.CTkButton(controls, text="Exportar resultado", state="disabled", command=self._export)
        self.export_button.pack(side="left", padx=10)
        self.clear_button = ctk.CTkButton(controls, text="Limpar", fg_color="gray35", command=self._clear)
        self.clear_button.pack(side="left", padx=10)

        self.files_label = ctk.CTkLabel(self.container, text="Selecione os dois arquivos.", text_color="gray70", anchor="w")
        self.files_label.grid(row=3, column=0, padx=30, pady=(10, 6), sticky="ew")

        dashboard = ctk.CTkFrame(self.container, fg_color="transparent")
        dashboard.grid(row=4, column=0, padx=30, pady=(0, 8), sticky="ew")
        dashboard.grid_columnconfigure(0, weight=3)
        dashboard.grid_columnconfigure(1, weight=2)

        summary = ctk.CTkFrame(dashboard, fg_color="transparent")
        summary.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        self.summary_labels: dict[str, ctk.CTkLabel] = {}
        for column, (title, key) in enumerate((
            ("Analisadas", "total"), ("Em dia", "on_time"), ("Em alerta", "alert"),
            ("Em atraso", "late"), ("Não lançadas", "not_posted"),
        )):
            grid_column = column % 3
            grid_row = column // 3
            summary.grid_columnconfigure(grid_column, weight=1)
            card = ctk.CTkFrame(summary)
            card.grid(
                row=grid_row, column=grid_column,
                padx=(0 if grid_column == 0 else 5, 0),
                pady=(0 if grid_row == 0 else 5, 0), sticky="nsew",
            )
            ctk.CTkLabel(card, text=title, text_color="gray70").pack(pady=(7, 0))
            label = ctk.CTkLabel(card, text="—", font=ctk.CTkFont(size=17, weight="bold"))
            label.pack(pady=(0, 7))
            self.summary_labels[key] = label

        chart_frame = ctk.CTkFrame(dashboard)
        chart_frame.grid(row=0, column=1, sticky="nsew")
        chart_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            chart_frame,
            text="Distribuição das notas",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="center",
        ).grid(row=0, column=0, padx=10, pady=(7, 0), sticky="ew")
        self.chart = tk.Canvas(chart_frame, height=125, background="#2b2b2b", highlightthickness=0)
        self.chart.grid(row=1, column=0, padx=8, pady=(0, 6), sticky="ew")
        self.chart.bind("<Configure>", lambda _: self._draw_chart())

        filters = ctk.CTkFrame(self.container, fg_color="transparent")
        filters.grid(row=5, column=0, padx=30, pady=(0, 8), sticky="ew")
        filters.grid_columnconfigure(2, weight=1)
        self.status_selector = ctk.CTkSegmentedButton(
            filters, values=["Em atraso", "Alerta", "Em dia", "Todas"], variable=self.filter_status,
            command=lambda _: self._reset_filter(),
        )
        self.status_selector.grid(row=0, column=0, padx=(0, 14))
        ctk.CTkLabel(filters, text="Buscar chave:").grid(row=0, column=1, padx=(0, 8))
        self.search_entry = ctk.CTkEntry(filters, textvariable=self.search_text, placeholder_text="Digite a chave da nota")
        self.search_entry.grid(row=0, column=2, sticky="ew")
        self.search_entry.bind("<KeyRelease>", lambda _: self._reset_filter())

        self.preview = create_result_table(self.container, (
            TableColumn("key", "Chave", 330), TableColumn("company", "Hotel", 170),
            TableColumn("supplier", "Fornecedor", 260), TableColumn("state", "UF", 55),
            TableColumn("emission", "Emissão", 90), TableColumn("entry", "Entrada", 90),
            TableColumn("days", "Dias", 65, "e"), TableColumn("limit", "Limite", 65, "e"),
            TableColumn("launch", "Lançamento", 150), TableColumn("status", "Situação", 110),
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
        names = filedialog.askopenfilenames(title="Selecionar Manifesto e Detalhe", filetypes=[("Planilhas Excel", "*.xlsx")])
        if not names:
            return
        self.paths = [Path(name) for name in names]
        self._busy(True)
        self.app.set_status("Analisando notas fiscais...", 0.1)
        self.app.run_background(lambda: analyze(self.paths), self._completed, self._failed)

    def _completed(self, result: AnalysisResult) -> None:
        self.result = result
        self.page = 0
        self._busy(False)
        self._render_preview()
        self.app.set_status("Análise concluída", 1)

    def _failed(self) -> None:
        self._busy(False)

    def _filtered(self) -> list[NoteResult]:
        if not self.result:
            return []
        selected = self.filter_status.get()
        search = self.search_text.get().strip()
        return [
            row for row in self.result.rows
            if (selected == "Todas" or row.status == selected)
            and (not search or search in row.key)
        ]

    def _render_preview(self) -> None:
        clear_table(self.preview)
        if not self.result:
            self.files_label.configure(text="Selecione os dois arquivos.")
            for label in self.summary_labels.values():
                label.configure(text="—")
            filtered: list[NoteResult] = []
        else:
            result = self.result
            values = {
                "total": len(result.rows), "on_time": result.count("Em dia"),
                "alert": result.count("Alerta"), "late": result.count("Em atraso"),
                "not_posted": sum(row.entry_date is None for row in result.rows),
            }
            for key, value in values.items():
                self.summary_labels[key].configure(text=f"{value:,}".replace(",", "."))
            self.files_label.configure(text="Arquivos carregados: 2 • Manifesto de notas e Detalhe de notas")
            filtered = self._filtered()
            total_pages = max(1, (len(filtered) + self.page_size - 1) // self.page_size)
            self.page = min(self.page, total_pages - 1)
            page_rows = filtered[self.page * self.page_size : (self.page + 1) * self.page_size]
            for row in page_rows:
                self.preview.insert("", "end", values=(row.key, row.company, row.supplier, row.state,
                    date_text(row.emission_date), date_text(row.entry_date), row.days if row.days is not None else "",
                    row.limit if row.limit is not None else "", row.launch_status, row.status),
                    tags=(result_tag(row.status),))
            self.page_label.configure(text=f"Página {self.page + 1} de {total_pages} • {len(filtered):,} registros".replace(",", "."))
            self.previous_button.configure(state="normal" if self.page > 0 else "disabled")
            self.next_button.configure(state="normal" if self.page + 1 < total_pages else "disabled")
        if not self.result:
            self.page_label.configure(text="Página 0 de 0")
            self.previous_button.configure(state="disabled")
            self.next_button.configure(state="disabled")
        self.export_button.configure(state="normal" if self.result else "disabled")
        self.app.after_idle(self._draw_chart)

    def _draw_chart(self) -> None:
        self.chart.delete("all")
        width = max(self.chart.winfo_width(), 300)
        center_x, center_y, radius = width * 0.34, 58, 48
        bounds = (center_x - radius, center_y - radius, center_x + radius, center_y + radius)
        if not self.result or not self.result.rows:
            self.chart.create_oval(*bounds, outline="#4b5563", width=14)
            self.chart.create_text(center_x, center_y, text="—", fill="#e5e7eb", font=("Segoe UI", 15, "bold"))
            self.chart.create_text(center_x, 125, text="Aguardando análise", fill="#cbd5e1")
            return

        total = len(self.result.rows)
        segments = (
            (self.result.count("Em dia"), "#21a67a", "Em dia"),
            (self.result.count("Alerta"), "#e0a83e", "Em alerta"),
            (self.result.count("Em atraso"), "#dc5a5a", "Em atraso"),
        )
        start = 90.0
        legend_x = width * 0.58
        legend_y = 30
        for count, color, label in segments:
            if not count:
                continue
            rate = count / total
            extent = -360 * rate
            self.chart.create_arc(*bounds, start=start, extent=extent, style="arc", outline=color, width=14)
            self.chart.create_text(
                legend_x, legend_y, text=f"● {label}: {rate:.1%}", fill=color, anchor="w"
            )
            start += extent
            legend_y += 23
        on_time_rate = self.result.count("Em dia") / total
        self.chart.create_text(
            center_x, center_y, text=f"{on_time_rate:.1%}", fill="#e5e7eb", font=("Segoe UI", 11, "bold")
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
        name = filedialog.asksaveasfilename(
            title="Salvar análise", defaultextension=extension,
            initialfile=f"notas_fiscais_entrada_em_atraso{extension}",
            filetypes=[(f"Arquivo {output_format}", f"*{extension}")],
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


AUTOMATION_CLASS = OverdueNotesAutomation
