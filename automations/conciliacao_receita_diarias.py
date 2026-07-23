from __future__ import annotations

import re
import unicodedata
import warnings
from collections import defaultdict
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from automations.legacy_ui import ctk, filedialog, messagebox, tk
from openpyxl import Workbook
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from automations.base import Automation
from automations.excel_reader import load_workbook_compatible as load_workbook
from automations.ui import TableColumn, clear_table, create_result_table


@dataclass(frozen=True)
class DailyRevenueRow:
    trx_code: str
    description: str
    daily: bool
    average_daily: bool
    transactions: int
    value: Decimal

    @property
    def status(self) -> str:
        return "Com movimento" if self.transactions else "Sem movimento"


@dataclass(frozen=True)
class DailyRevenueResult:
    hotel: str
    rows: list[DailyRevenueRow]
    journal_rows: int

    @property
    def daily_total(self) -> Decimal:
        return sum((row.value for row in self.rows if row.daily), Decimal())

    @property
    def average_daily_total(self) -> Decimal:
        return sum((row.value for row in self.rows if row.average_daily), Decimal())

    @property
    def moved(self) -> int:
        return sum(row.transactions > 0 for row in self.rows)


def normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return re.sub(r"[^A-Z0-9]", "", "".join(char for char in text if not unicodedata.combining(char)).upper())


def trx_code(value: Any) -> str:
    text = str(value or "").strip()
    return str(int(float(text))) if re.fullmatch(r"\d+(?:\.0+)?", text) else normalize(text)


def decimal_value(value: Any) -> Decimal:
    if value in (None, "", "NULL"):
        return Decimal()
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return Decimal(str(value).replace(".", "").replace(",", "."))


def money(value: Decimal) -> str:
    return f"R$ {value:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def workbook_headers(path: Path) -> list[tuple[str, set[str]]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    result = []
    for sheet in workbook.worksheets:
        sheet.reset_dimensions()
        header = next(sheet.iter_rows(values_only=True), ())
        result.append((sheet.title, {normalize(value) for value in header if value not in (None, "")}))
    workbook.close()
    return result


def identify_file(path: Path) -> str:
    headers = workbook_headers(path)
    if any({"TRXCODE", "DIARIA", "DIARIAMEDIA"}.issubset(columns) for _, columns in headers):
        return "codes"
    if any({"TRXCODE", "CASHIERDEBIT"}.issubset(columns) for _, columns in headers):
        return "journal"
    return "unknown"


def read_rules(path: Path, hotel: str) -> dict[str, tuple[str, bool, bool]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = next((item for item in workbook.worksheets if normalize(item.title) == normalize(hotel)), None)
        if sheet is None:
            raise ValueError(f"O hotel {hotel} não existe na planilha de códigos de transação.")
        sheet.reset_dimensions()
        rows = sheet.iter_rows(values_only=True)
        header = tuple(next(rows))
        indexes = {normalize(value): index for index, value in enumerate(header)}
        required = {"TRXCODE", "DIARIA", "DIARIAMEDIA"}
        if not required.issubset(indexes):
            raise ValueError("A planilha de códigos não contém TRX_CODE, DIÁRIA e DIÁRIA MÉDIA.")
        description_index = indexes.get("D3")
        result = {}
        for row in rows:
            code_index = indexes["TRXCODE"]
            if code_index >= len(row) or not trx_code(row[code_index]):
                continue
            daily = indexes["DIARIA"] < len(row) and normalize(row[indexes["DIARIA"]]) == "SIM"
            average = indexes["DIARIAMEDIA"] < len(row) and normalize(row[indexes["DIARIAMEDIA"]]) == "SIM"
            if not daily and not average:
                continue
            description = str(row[description_index] or "").strip() if description_index is not None and description_index < len(row) else ""
            result[trx_code(row[code_index])] = description, daily, average
        if not result:
            raise ValueError(f"Nenhum TRX_CODE marcado como SIM foi encontrado para {hotel}.")
        return result
    finally:
        workbook.close()


def read_journal(path: Path) -> tuple[dict[str, tuple[int, Decimal]], int]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        sheet = workbook.active
        sheet.reset_dimensions()
        rows = sheet.iter_rows(values_only=True)
        header = tuple(next(rows))
        indexes = {normalize(value): index for index, value in enumerate(header)}
        if not {"TRXCODE", "CASHIERDEBIT"}.issubset(indexes):
            raise ValueError("O Journal não contém TRX_CODE e CASHIER_DEBIT.")
        grouped: dict[str, list[Any]] = defaultdict(lambda: [0, Decimal()])
        total_rows = 0
        for row in rows:
            code_index, value_index = indexes["TRXCODE"], indexes["CASHIERDEBIT"]
            if code_index >= len(row) or not trx_code(row[code_index]):
                continue
            code = trx_code(row[code_index])
            grouped[code][0] += 1
            grouped[code][1] += decimal_value(row[value_index] if value_index < len(row) else None)
            total_rows += 1
        return {code: (data[0], data[1]) for code, data in grouped.items()}, total_rows
    finally:
        workbook.close()


def analyze(paths: list[Path], hotel: str) -> DailyRevenueResult:
    if len(paths) != 2:
        raise ValueError("Selecione a planilha de códigos de transação e o Journal.")
    identified = [(path, identify_file(path)) for path in paths]
    codes = [path for path, kind in identified if kind == "codes"]
    journals = [path for path, kind in identified if kind == "journal"]
    unknown = [path.name for path, kind in identified if kind == "unknown"]
    if unknown:
        selected = ", ".join(unknown)
        guidance = (
            "Esta conferência aceita somente a planilha 'Códigos de transação' "
            "e o 'Journal Opera - Receita'."
        )
        if any("BIPDV" in normalize(name) for name in unknown):
            guidance += " O arquivo BI PDV pertence à automação 'Cupons Emitidos x Conta do Hóspede'."
        raise ValueError(f"Arquivo incompatível com Receita de Diárias: {selected}. {guidance}")
    if len(codes) != 1 or len(journals) != 1:
        raise ValueError("Envie uma planilha de códigos de transação e um Journal.")
    rules = read_rules(codes[0], hotel)
    journal, journal_rows = read_journal(journals[0])
    rows = [
        DailyRevenueRow(code, description, daily, average, *journal.get(code, (0, Decimal())))
        for code, (description, daily, average) in rules.items()
    ]
    rows.sort(key=lambda row: (row.transactions == 0, -abs(row.value), row.trx_code))
    return DailyRevenueResult(hotel, rows, journal_rows)


def save_excel(result: DailyRevenueResult, path: Path) -> None:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumo"
    summary.append(["Indicador", "Resultado"])
    for label, value in (
        ("Hotel", result.hotel),
        ("Receita de diárias", float(result.daily_total)),
        ("Receita considerada na diária média", float(result.average_daily_total)),
        ("TRX_CODE considerados", len(result.rows)),
        ("TRX_CODE com movimento", result.moved),
        ("TRX_CODE sem movimento", len(result.rows) - result.moved),
        ("Lançamentos lidos no Journal", result.journal_rows),
    ):
        summary.append([label, value])
    for cell in summary[1]:
        cell.style = "Headline 4"
    for cell in summary["B"][2:4]:
        cell.number_format = 'R$ #,##0.00'
    summary.column_dimensions["A"].width = 44
    summary.column_dimensions["B"].width = 24

    details = workbook.create_sheet("Detalhamento")
    details.append(["TRX_CODE", "Descrição", "Diária", "Diária média", "Lançamentos", "CASHIER_DEBIT", "Situação"])
    for row in result.rows:
        details.append([row.trx_code, row.description, "SIM" if row.daily else "NÃO",
                        "SIM" if row.average_daily else "NÃO", row.transactions, float(row.value), row.status])
    for cell in details[1]:
        cell.style = "Headline 4"
    for cell in details["F"][1:]:
        cell.number_format = 'R$ #,##0.00'
    for column, width in {"A": 14, "B": 48, "C": 14, "D": 18, "E": 16, "F": 20, "G": 18}.items():
        details.column_dimensions[column].width = width
    details.freeze_panes = "A2"
    details.auto_filter.ref = details.dimensions
    workbook.save(path)


def save_pdf(result: DailyRevenueResult, path: Path) -> None:
    styles = getSampleStyleSheet()
    data = [
        ["Hotel", "Receita de diárias", "Receita diária média", "Códigos", "Com movimento", "Sem movimento"],
        [result.hotel, money(result.daily_total), money(result.average_daily_total), str(len(result.rows)),
         str(result.moved), str(len(result.rows) - result.moved)],
    ]
    table = Table(data, colWidths=[40 * mm, 48 * mm, 48 * mm, 30 * mm, 35 * mm, 35 * mm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#24588A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), .5, colors.grey),
    ]))
    document = SimpleDocTemplate(str(path), pagesize=landscape(A4), title="Receita de Diárias")
    document.build([
        Paragraph("Conciliação da Receita de Diárias", styles["Title"]),
        Spacer(1, 5 * mm), table,
    ])


class DailyRevenueAutomation(Automation):
    name = "Conciliação da Receita de Diárias"

    def __init__(self, app, container):
        super().__init__(app, container)
        self.paths: list[Path] = []
        self.result: DailyRevenueResult | None = None
        self.hotel = ctk.StringVar(value="Magna")
        self.output_format = ctk.StringVar(value="Excel")
        self.status_filter = ctk.StringVar(value="Todos")
        self.search = ctk.StringVar()
        self.page = 0
        self.page_size = 50

    def render(self):
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(6, weight=1)
        ctk.CTkLabel(self.container, text=self.name, font=ctk.CTkFont(size=26, weight="bold")).grid(row=0, column=0, padx=30, pady=(22, 3), sticky="w")
        ctk.CTkLabel(self.container, text="Classifica e totaliza o CASHIER_DEBIT do Journal pelos TRX_CODE de diárias do hotel.", text_color="gray70").grid(row=1, column=0, padx=30, pady=(0, 10), sticky="w")

        controls = ctk.CTkFrame(self.container, fg_color="transparent")
        controls.grid(row=2, column=0, padx=30, sticky="ew")
        ctk.CTkLabel(controls, text="Hotel:").pack(side="left", padx=(0, 6))
        self.hotel_menu = ctk.CTkOptionMenu(controls, values=["Cumbuco", "Magna", "Taiba", "Charme"], variable=self.hotel, width=125)
        self.hotel_menu.pack(side="left", padx=(0, 12))
        self.select = ctk.CTkButton(controls, text="Selecionar os dois arquivos", command=self._select)
        self.select.pack(side="left", padx=(0, 10))
        ctk.CTkSegmentedButton(controls, values=["Excel", "PDF"], variable=self.output_format).pack(side="left", padx=10)
        self.export = ctk.CTkButton(controls, text="Exportar resultado", state="disabled", command=self._export)
        self.export.pack(side="left", padx=10)
        ctk.CTkButton(controls, text="Limpar", fg_color="gray35", command=self._clear).pack(side="left", padx=10)

        self.info = ctk.CTkLabel(self.container, text="Selecione a planilha de códigos de transação e o Journal; os nomes dos arquivos são livres.", text_color="gray70", anchor="w")
        self.info.grid(row=3, column=0, padx=30, pady=(10, 6), sticky="ew")

        dashboard = ctk.CTkFrame(self.container, fg_color="transparent")
        dashboard.grid(row=4, column=0, padx=30, pady=(0, 8), sticky="ew")
        dashboard.grid_columnconfigure(0, weight=3)
        dashboard.grid_columnconfigure(1, weight=1)
        cards = ctk.CTkFrame(dashboard, fg_color="transparent")
        cards.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        self.cards = {}
        for index, title in enumerate(("Receita de diárias", "Receita diária média", "Códigos considerados", "Com movimento")):
            cards.grid_columnconfigure(index, weight=1)
            card = ctk.CTkFrame(cards)
            card.grid(row=0, column=index, padx=(0 if index == 0 else 5, 0), sticky="nsew")
            ctk.CTkLabel(card, text=title, text_color="gray70").pack(pady=(8, 0))
            label = ctk.CTkLabel(card, text="—", font=ctk.CTkFont(size=16, weight="bold"))
            label.pack(pady=(0, 8))
            self.cards[title] = label
        chart_frame = ctk.CTkFrame(dashboard)
        chart_frame.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(chart_frame, text="Distribuição dos códigos", font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(6, 0))
        self.chart = tk.Canvas(chart_frame, height=125, bg="#2b2b2b", highlightthickness=0)
        self.chart.pack(fill="x", padx=8)
        self.chart.bind("<Configure>", lambda _: self._draw_chart())

        filters = ctk.CTkFrame(self.container, fg_color="transparent")
        filters.grid(row=5, column=0, padx=30, pady=(0, 8), sticky="ew")
        filters.grid_columnconfigure(2, weight=1)
        ctk.CTkSegmentedButton(filters, values=["Todos", "Com movimento", "Sem movimento"], variable=self.status_filter, command=lambda _: self._reset()).grid(row=0, column=0, padx=(0, 12))
        ctk.CTkLabel(filters, text="Buscar:").grid(row=0, column=1, padx=(0, 8))
        entry = ctk.CTkEntry(filters, textvariable=self.search, placeholder_text="TRX_CODE ou descrição")
        entry.grid(row=0, column=2, sticky="ew")
        entry.bind("<KeyRelease>", lambda _: self._reset())
        ctk.CTkLabel(filters, text="Verde: código com lançamentos no Journal  •  Amarelo: código selecionado sem movimento", text_color="gray70", anchor="w").grid(row=1, column=0, columnspan=3, pady=(6, 0), sticky="ew")

        self.preview = create_result_table(self.container, (
            TableColumn("code", "TRX_CODE", 110), TableColumn("description", "Descrição", 380),
            TableColumn("daily", "Diária", 90), TableColumn("average", "Diária média", 110),
            TableColumn("transactions", "Lançamentos", 110, "center"),
            TableColumn("value", "CASHIER_DEBIT", 145, "e"), TableColumn("status", "Situação", 140),
        ), row=6)

        pagination = ctk.CTkFrame(self.container, fg_color="transparent")
        pagination.grid(row=7, column=0, padx=30, pady=(7, 14), sticky="ew")
        pagination.grid_columnconfigure(1, weight=1)
        self.previous = ctk.CTkButton(pagination, text="Anterior", width=100, command=self._previous)
        self.previous.grid(row=0, column=0)
        self.page_label = ctk.CTkLabel(pagination, text="Página 0 de 0")
        self.page_label.grid(row=0, column=1)
        self.next = ctk.CTkButton(pagination, text="Próxima", width=100, command=self._next)
        self.next.grid(row=0, column=2)
        self._show()

    def _select(self):
        names = filedialog.askopenfilenames(title="Arquivos da receita de diárias", filetypes=[("Planilhas Excel", "*.xlsx *.xlsm *.xls *.xltx *.xltm")])
        if not names:
            return
        self.paths = [Path(name) for name in names]
        selected_hotel = self.hotel.get()
        self.select.configure(state="disabled")
        self.hotel_menu.configure(state="disabled")
        self.app.set_status("Analisando receita de diárias...", .1)
        self.app.run_background(lambda: analyze(self.paths, selected_hotel), self._done, self._failed)

    def _done(self, result):
        self.result = result
        self.page = 0
        self.select.configure(state="normal")
        self.hotel_menu.configure(state="normal")
        self.export.configure(state="normal")
        self.info.configure(text=f"Hotel: {result.hotel} • 2 arquivos carregados • {result.journal_rows:,} lançamentos lidos".replace(",", "."))
        self.app.set_status("Apuração de diárias concluída", 1)
        self._show()

    def _failed(self):
        self.select.configure(state="normal")
        self.hotel_menu.configure(state="normal")

    def _filtered(self):
        if not self.result:
            return []
        status, search = self.status_filter.get(), normalize(self.search.get())
        return [row for row in self.result.rows
                if (status == "Todos" or row.status == status)
                and (not search or search in normalize(row.trx_code + row.description))]

    def _show(self):
        clear_table(self.preview)
        if not self.result:
            self.page_label.configure(text="Página 0 de 0")
            self.previous.configure(state="disabled")
            self.next.configure(state="disabled")
        else:
            values = (money(self.result.daily_total), money(self.result.average_daily_total),
                      str(len(self.result.rows)), str(self.result.moved))
            for title, value in zip(self.cards, values):
                self.cards[title].configure(text=value)
            filtered = self._filtered()
            pages = max(1, (len(filtered) + self.page_size - 1) // self.page_size)
            self.page = min(self.page, pages - 1)
            visible = filtered[self.page * self.page_size:(self.page + 1) * self.page_size]
            for row in visible:
                self.preview.insert("", "end", values=(row.trx_code, row.description,
                                    "SIM" if row.daily else "NÃO", "SIM" if row.average_daily else "NÃO",
                                    row.transactions, money(row.value), row.status),
                                    tags=("ok" if row.transactions else "missing",))
            self.page_label.configure(text=f"Página {self.page + 1} de {pages} • {len(filtered)} códigos")
            self.previous.configure(state="normal" if self.page else "disabled")
            self.next.configure(state="normal" if self.page + 1 < pages else "disabled")
        self.app.after_idle(self._draw_chart)

    def _draw_chart(self):
        self.chart.delete("all")
        width = max(self.chart.winfo_width(), 180)
        if not self.result:
            self.chart.create_text(width / 2, 62, text="Aguardando análise", fill="#9ca3af")
            return
        moved, total = self.result.moved, len(self.result.rows)
        ratio = moved / total if total else 0
        size, top = 70, 24
        left = (width - size) / 2
        self.chart.create_oval(left, top, left + size, top + size, outline="#dc5a5a", width=13)
        self.chart.create_arc(left, top, left + size, top + size, start=90, extent=-360 * ratio,
                              style="arc", outline="#21a67a", width=13)
        self.chart.create_text(width / 2, top + size / 2, text=f"{ratio:.1%}", fill="white", font=("Segoe UI", 11, "bold"))
        self.chart.create_text(width / 2, 108, text=f"{moved} com movimento • {total - moved} sem movimento", fill="#d1d5db", font=("Segoe UI", 9))

    def _reset(self):
        self.page = 0
        self._show()

    def _previous(self):
        if self.page:
            self.page -= 1
            self._show()

    def _next(self):
        self.page += 1
        self._show()

    def _export(self):
        if not self.result:
            return
        extension = ".xlsx" if self.output_format.get() == "Excel" else ".pdf"
        name = filedialog.asksaveasfilename(title="Exportar resultado", defaultextension=extension,
                                            initialfile=f"conciliacao_receita_diarias_{normalize(self.result.hotel).lower()}{extension}",
                                            filetypes=[(self.output_format.get(), f"*{extension}")])
        if not name:
            return
        self.app.set_status("Exportando resultado...", .5)
        task = save_excel if extension == ".xlsx" else save_pdf
        self.app.run_background(lambda: task(self.result, Path(name)),
                                lambda _: (self.app.set_status("Resultado exportado", 1), messagebox.showinfo("Concluído", "Resultado exportado com sucesso.")))

    def _clear(self):
        self.paths = []
        self.result = None
        self.search.set("")
        self.status_filter.set("Todos")
        self.page = 0
        self.export.configure(state="disabled")
        self.info.configure(text="Selecione a planilha de códigos de transação e o Journal; os nomes dos arquivos são livres.")
        self._show()
        self.app.set_status("Seleção limpa", 0)


AUTOMATION_CLASS = DailyRevenueAutomation
