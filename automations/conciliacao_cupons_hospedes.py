from __future__ import annotations

import re
import unicodedata
import warnings
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
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
from automations.ui import TableColumn, clear_table, create_result_table


@dataclass(frozen=True)
class CouponResult:
    company: str
    pdv: str
    issue_date: date
    posting_date: date | None
    document: str
    account: str
    room: str
    guest: str
    document_type: str
    pdv_value: Decimal
    journal_value: Decimal | None
    status: str
    detail: str

    @property
    def difference(self) -> Decimal | None:
        return None if self.journal_value is None else self.pdv_value - self.journal_value

    @property
    def reconciled(self) -> bool:
        return self.status.startswith("Conciliado")

    @property
    def incomplete_period(self) -> bool:
        return self.status == "Journal não cobre a data"


@dataclass(frozen=True)
class ReconciliationResult:
    company: str
    mapping: str
    journal_start: date
    journal_end: date
    coupons: list[CouponResult]
    files: tuple[str, str, str]

    @property
    def reconciled(self) -> int:
        return sum(item.reconciled for item in self.coupons)

    @property
    def incomplete(self) -> int:
        return sum(item.incomplete_period for item in self.coupons)

    @property
    def issues(self) -> int:
        return len(self.coupons) - self.reconciled - self.incomplete

    @property
    def total_pdv(self) -> Decimal:
        return sum((item.pdv_value for item in self.coupons), Decimal())


@dataclass
class _Coupon:
    company: str
    pdv: str
    issue_date: date
    document: str
    account: str
    room: str
    guest: str
    document_type: str
    value: Decimal = Decimal()


@dataclass(frozen=True)
class _JournalRow:
    code: str
    check: str
    posting_date: date
    value: Decimal
    room: str


def normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return re.sub(r"[^A-Z0-9]", "", "".join(char for char in text if not unicodedata.combining(char)).upper())


def decimal_value(value: Any) -> Decimal:
    if value in (None, "", "-"):
        return Decimal()
    text = str(value).strip().replace("R$", "").replace(" ", "")
    if "," in text:
        text = text.replace(".", "").replace(",", ".")
    try:
        return Decimal(text)
    except InvalidOperation:
        return Decimal()


def money(value: Decimal | None) -> str:
    if value is None:
        return "—"
    return f"R$ {value:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def parse_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    for pattern in ("%d/%m/%Y", "%d-%m-%y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def _header_map(values: tuple[Any, ...]) -> dict[str, int]:
    result = {}
    for index, value in enumerate(values):
        key = normalize(value)
        if key and key not in result:
            result[key] = index
    return result


def identify_file(path: Path) -> str:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return "unknown"
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        first_headers = _header_map(tuple(next(workbook.active.iter_rows(values_only=True))))
        keys = set(first_headers)
        if {"EMPRESA", "PDV", "VALOR", "DATADEEMISSAO", "NODODOCUMENTO", "CONTA", "DADOSHOSPEDE"}.issubset(keys):
            return "pdv"
        if {"TRXCODE", "REFERENCE", "CASHIERDEBIT", "BUSINESSFORMATDATE", "ROOM"}.issubset(keys):
            return "journal"
        if workbook.sheetnames and all(normalize(next(sheet.iter_rows(values_only=True))[0]) == "TRXCODE" for sheet in workbook.worksheets):
            return "mapping"
        return "unknown"
    finally:
        workbook.close()


def _read_pdv(path: Path) -> dict[tuple[str, str, date, str], _Coupon]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        rows = workbook.active.iter_rows(values_only=True)
        headers = _header_map(tuple(next(rows)))

        def value(row, name):
            index = headers[name]
            return row[index] if index < len(row) else None

        coupons: dict[tuple[str, str, date, str], _Coupon] = {}
        for row in rows:
            company = str(value(row, "EMPRESA") or "").strip()
            account = re.sub(r"\D", "", str(value(row, "CONTA") or ""))
            issued = parse_date(value(row, "DATADEEMISSAO"))
            document = str(value(row, "NODODOCUMENTO") or "").strip()
            if not company or not account or not issued or not document or normalize(company) in {"TOTAL", "FILTROSAPLICADOS"}:
                continue
            guest_data = str(value(row, "DADOSHOSPEDE") or "").strip()
            parts = [part.strip() for part in guest_data.split("/")]
            key = (company, account, issued, document)
            if key not in coupons:
                coupons[key] = _Coupon(
                    company, str(value(row, "PDV") or "").strip(), issued, document, account,
                    parts[0].zfill(4) if parts and parts[0].isdigit() else (parts[0] if parts else ""),
                    parts[1] if len(parts) > 1 else "",
                    str(value(row, "TIPODEDOCUMENTO") or "").strip(),
                )
            coupons[key].value += decimal_value(value(row, "VALOR"))
        for coupon in coupons.values():
            coupon.value = coupon.value.quantize(Decimal("0.01"))
        return coupons
    finally:
        workbook.close()


def _read_journal(path: Path) -> list[_JournalRow]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        rows = workbook.active.iter_rows(values_only=True)
        headers = _header_map(tuple(next(rows)))

        def value(row, name):
            index = headers[name]
            return row[index] if index < len(row) else None

        result = []
        for row in rows:
            reference = str(value(row, "REFERENCE") or "")
            match = re.search(r"CHECK#\s*(\d+)", reference, re.IGNORECASE)
            posting_date = parse_date(value(row, "BUSINESSFORMATDATE"))
            if not match or not posting_date:
                continue
            code_value = value(row, "TRXCODE")
            code = str(int(code_value)) if isinstance(code_value, float) and code_value.is_integer() else str(code_value or "").strip()
            result.append(_JournalRow(code, match.group(1), posting_date,
                                      decimal_value(value(row, "CASHIERDEBIT")),
                                      str(value(row, "ROOM") or "").strip().zfill(4)))
        return result
    finally:
        workbook.close()


def _read_mappings(path: Path) -> dict[str, set[str]]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    try:
        result = {}
        for sheet in workbook.worksheets:
            codes = set()
            for (value,) in sheet.iter_rows(min_row=2, max_col=1, values_only=True):
                if value is not None:
                    codes.add(str(int(value)) if isinstance(value, float) and value.is_integer() else str(value).strip())
            result[sheet.title] = codes
        return result
    finally:
        workbook.close()


def _match_account(check: str, accounts: set[str]) -> str | None:
    if check in accounts:
        return check
    matches = [account for account in accounts if check.endswith(account)]
    return max(matches, key=len) if matches else None


def analyze(paths: list[Path]) -> ReconciliationResult:
    if len(paths) != 3:
        raise ValueError("Selecione o BI/PDV, o Journal e o arquivo de de/para dos TRX_CODE.")
    identified = [(path, identify_file(path)) for path in paths]
    unknown = [path.name for path, kind in identified if kind == "unknown"]
    if unknown:
        raise ValueError(f"Arquivo não reconhecido pelo conteúdo: {', '.join(unknown)}.")
    grouped = {kind: [path for path, current in identified if current == kind] for kind in ("pdv", "journal", "mapping")}
    invalid = [kind for kind, items in grouped.items() if len(items) != 1]
    if invalid:
        raise ValueError("Envie exatamente um BI/PDV, um Journal e um arquivo de de/para.")

    pdv = _read_pdv(grouped["pdv"][0])
    journal = _read_journal(grouped["journal"][0])
    mappings = _read_mappings(grouped["mapping"][0])
    if not pdv or not journal or not mappings:
        raise ValueError("Um dos arquivos não contém registros utilizáveis.")

    companies: dict[str, set[str]] = defaultdict(set)
    for coupon in pdv.values():
        companies[coupon.company].add(coupon.account)
    best: tuple[int, str, str] | None = None
    for mapping_name, codes in mappings.items():
        checks = [row.check for row in journal if row.code in codes]
        for company, accounts in companies.items():
            matched = len({_match_account(check, accounts) for check in checks} - {None})
            candidate = (matched, mapping_name, company)
            if best is None or candidate > best:
                best = candidate
    if not best or best[0] == 0:
        raise ValueError("Não foi encontrada relação entre as contas do BI/PDV e os CHECK# do Journal.")
    _, mapping_name, company = best
    accounts = companies[company]
    selected_journal = [row for row in journal if row.code in mappings[mapping_name]]
    journal_start = min(row.posting_date for row in selected_journal)
    journal_end = max(row.posting_date for row in selected_journal)
    postings: dict[str, dict[date, Decimal]] = defaultdict(lambda: defaultdict(Decimal))
    for row in selected_journal:
        account = _match_account(row.check, accounts)
        if account:
            postings[account][row.posting_date] += row.value

    results = []
    for coupon in (item for item in pdv.values() if item.company == company):
        by_date = postings.get(coupon.account, {})
        posting_date = coupon.issue_date if coupon.issue_date in by_date else None
        detail = "Conta e valor localizados no Journal."
        if posting_date is None and by_date:
            exact = [day for day, value in by_date.items() if abs(value - coupon.value) <= Decimal("0.01")]
            if exact:
                posting_date = min(exact, key=lambda day: (abs((day - coupon.issue_date).days), day))
        if posting_date is not None:
            journal_value = by_date[posting_date].quantize(Decimal("0.01"))
            difference = coupon.value - journal_value
            if journal_value == 0 and coupon.value != 0:
                status, detail = "Não cobrado", "O CHECK# existe, mas o valor líquido lançado no Journal é zero."
            elif abs(difference) > Decimal("0.01"):
                status, detail = "Valor divergente", f"Diferença de {money(difference)} entre o cupom e a conta."
            elif posting_date != coupon.issue_date:
                status = "Conciliado - data diferente"
                detail = f"Cobrado no Journal em {posting_date:%d/%m/%Y}."
            else:
                status = "Conciliado"
        elif coupon.issue_date < journal_start or coupon.issue_date > journal_end:
            journal_value = None
            status = "Journal não cobre a data"
            detail = f"O Journal selecionado cobre {journal_start:%d/%m/%Y} a {journal_end:%d/%m/%Y}."
        elif by_date:
            nearest = min(by_date, key=lambda day: abs((day - coupon.issue_date).days))
            journal_value = by_date[nearest].quantize(Decimal("0.01"))
            posting_date = nearest
            status = "Lançado em outra data"
            detail = f"CHECK# localizado em {nearest:%d/%m/%Y}, mas sem correspondência segura de valor."
        else:
            journal_value = None
            status = "Ausente na conta"
            detail = "Cupom emitido no BI/PDV sem CHECK# correspondente no Journal."
        results.append(CouponResult(company, coupon.pdv, coupon.issue_date, posting_date, coupon.document,
                                    coupon.account, coupon.room, coupon.guest, coupon.document_type,
                                    coupon.value, journal_value, status, detail))
    results.sort(key=lambda item: (item.reconciled, item.incomplete_period, item.issue_date, item.account))
    return ReconciliationResult(company, mapping_name, journal_start, journal_end, results,
                                tuple(path.name for path in paths))


def save_excel(result: ReconciliationResult, path: Path) -> None:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumo"
    summary.append(["Indicador", "Resultado"])
    for label, value in (
        ("Hotel", result.company), ("De/para identificado", result.mapping),
        ("Período do Journal", f"{result.journal_start:%d/%m/%Y} a {result.journal_end:%d/%m/%Y}"),
        ("Cupons analisados", len(result.coupons)), ("Conciliados", result.reconciled),
        ("Pendências", result.issues), ("Fora do período", result.incomplete),
        ("Valor total dos cupons", float(result.total_pdv)),
    ):
        summary.append([label, value])
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 42
    summary["B9"].number_format = 'R$ #,##0.00'

    detail = workbook.create_sheet("Conferencia")
    detail.append(["Hotel", "Data cupom", "Data Journal", "PDV", "Cupom", "Conta/CHECK", "Quarto",
                   "Hóspede", "Valor cupom", "Valor Journal", "Diferença", "Resultado", "Explicação"])
    for item in result.coupons:
        detail.append([item.company, item.issue_date, item.posting_date, item.pdv, item.document, item.account,
                       item.room, item.guest, float(item.pdv_value),
                       None if item.journal_value is None else float(item.journal_value),
                       None if item.difference is None else float(item.difference), item.status, item.detail])
    for column in ("B", "C"):
        for cell in detail[column][1:]:
            cell.number_format = "dd/mm/yyyy"
    for column in ("I", "J", "K"):
        for cell in detail[column][1:]:
            cell.number_format = 'R$ #,##0.00'
    for column, width in {"A": 22, "B": 14, "C": 14, "D": 20, "E": 12, "F": 16, "G": 10,
                          "H": 32, "I": 16, "J": 16, "K": 16, "L": 28, "M": 65}.items():
        detail.column_dimensions[column].width = width
    detail.freeze_panes = "A2"
    detail.auto_filter.ref = detail.dimensions
    workbook.save(path)


def save_pdf(result: ReconciliationResult, path: Path) -> None:
    styles = getSampleStyleSheet()
    data = [
        ["Cupons", "Conciliados", "Pendências", "Fora do período", "Valor total"],
        [str(len(result.coupons)), str(result.reconciled), str(result.issues), str(result.incomplete), money(result.total_pdv)],
    ]
    table = Table(data, colWidths=[35 * mm, 40 * mm, 35 * mm, 45 * mm, 50 * mm])
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#24588A")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), .5, colors.grey),
    ]))
    pending = [item for item in result.coupons if not item.reconciled][:30]
    pending_data = [["Data", "Cupom", "Conta", "Hóspede", "Valor", "Resultado"]]
    pending_data.extend([[item.issue_date.strftime("%d/%m/%Y"), item.document, item.account, item.guest[:28],
                          money(item.pdv_value), item.status] for item in pending])
    pending_table = Table(pending_data, colWidths=[25 * mm, 25 * mm, 30 * mm, 65 * mm, 30 * mm, 55 * mm], repeatRows=1)
    pending_table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#374151")),
                                       ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                                       ("FONTSIZE", (0, 0), (-1, -1), 8),
                                       ("GRID", (0, 0), (-1, -1), .4, colors.grey)]))
    story = [Paragraph("Cupons emitidos x conta do hóspede", styles["Title"]),
             Paragraph(f"{result.company} — Journal de {result.journal_start:%d/%m/%Y} a {result.journal_end:%d/%m/%Y}", styles["BodyText"]),
             Spacer(1, 5 * mm), table]
    if pending:
        story.extend([Spacer(1, 6 * mm), Paragraph("Cupons que exigem atenção", styles["Heading2"]), pending_table])
    SimpleDocTemplate(str(path), pagesize=landscape(A4), title="Conferência de cupons").build(story)


class GuestCouponAutomation(Automation):
    name = "Cupons Emitidos x Conta do Hóspede"

    def __init__(self, app, container):
        super().__init__(app, container)
        self.result: ReconciliationResult | None = None
        self.output_format = ctk.StringVar(value="Excel")
        self.status_filter = ctk.StringVar(value="Pendências")
        self.search = ctk.StringVar()
        self.page = 0
        self.page_size = 50

    def render(self):
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(6, weight=1)
        ctk.CTkLabel(self.container, text=self.name, font=ctk.CTkFont(size=26, weight="bold")).grid(row=0, column=0, padx=30, pady=(22, 3), sticky="w")
        ctk.CTkLabel(self.container, text="Verifica se cada cupom do BI/PDV foi lançado e cobrado na conta do hóspede no Journal.", text_color="gray70").grid(row=1, column=0, padx=30, pady=(0, 10), sticky="w")
        controls = ctk.CTkFrame(self.container, fg_color="transparent")
        controls.grid(row=2, column=0, padx=30, sticky="ew")
        self.select = ctk.CTkButton(controls, text="Selecionar BI/PDV, Journal e de/para", command=self._select)
        self.select.pack(side="left", padx=(0, 10))
        ctk.CTkSegmentedButton(controls, values=["Excel", "PDF"], variable=self.output_format).pack(side="left", padx=10)
        self.export = ctk.CTkButton(controls, text="Exportar resultado", state="disabled", command=self._export)
        self.export.pack(side="left", padx=10)
        ctk.CTkButton(controls, text="Limpar", fg_color="gray35", command=self._clear).pack(side="left", padx=10)
        self.info = ctk.CTkLabel(self.container, text="Os arquivos são reconhecidos pelas colunas; o nome não é utilizado.", text_color="gray70", anchor="w")
        self.info.grid(row=3, column=0, padx=30, pady=(10, 6), sticky="ew")

        dashboard = ctk.CTkFrame(self.container, fg_color="transparent")
        dashboard.grid(row=4, column=0, padx=30, pady=(0, 8), sticky="ew")
        dashboard.grid_columnconfigure(0, weight=3)
        dashboard.grid_columnconfigure(1, weight=1)
        cards = ctk.CTkFrame(dashboard, fg_color="transparent")
        cards.grid(row=0, column=0, padx=(0, 10), sticky="nsew")
        self.cards = {}
        for index, title in enumerate(("Cupons analisados", "Conciliados", "Pendências", "Journal incompleto")):
            cards.grid_columnconfigure(index, weight=1)
            card = ctk.CTkFrame(cards)
            card.grid(row=0, column=index, padx=(0 if index == 0 else 5, 0), sticky="nsew")
            ctk.CTkLabel(card, text=title, text_color="gray70").pack(pady=(8, 0))
            label = ctk.CTkLabel(card, text="—", font=ctk.CTkFont(size=16, weight="bold"))
            label.pack(pady=(0, 8))
            self.cards[title] = label
        chart_frame = ctk.CTkFrame(dashboard)
        chart_frame.grid(row=0, column=1, sticky="nsew")
        ctk.CTkLabel(chart_frame, text="Distribuição da conferência", font=ctk.CTkFont(size=13, weight="bold")).pack(pady=(6, 0))
        self.chart = tk.Canvas(chart_frame, height=125, bg="#2b2b2b", highlightthickness=0)
        self.chart.pack(fill="x", padx=8)
        self.chart.bind("<Configure>", lambda _: self._draw_chart())

        filters = ctk.CTkFrame(self.container, fg_color="transparent")
        filters.grid(row=5, column=0, padx=30, pady=(0, 8), sticky="ew")
        filters.grid_columnconfigure(2, weight=1)
        ctk.CTkSegmentedButton(filters, values=["Pendências", "Conciliados", "Todos"], variable=self.status_filter, command=lambda _: self._reset()).grid(row=0, column=0, padx=(0, 12))
        ctk.CTkLabel(filters, text="Buscar:").grid(row=0, column=1, padx=(0, 8))
        entry = ctk.CTkEntry(filters, textvariable=self.search, placeholder_text="Cupom, conta, hóspede, quarto, PDV ou resultado")
        entry.grid(row=0, column=2, sticky="ew")
        entry.bind("<KeyRelease>", lambda _: self._reset())
        ctk.CTkLabel(filters, text="Verde: lançado e cobrado  •  Vermelho: ausência ou valor incorreto  •  Amarelo: Journal não cobre a data", text_color="gray70", anchor="w").grid(row=1, column=0, columnspan=3, pady=(6, 0), sticky="ew")

        self.preview = create_result_table(self.container, (
            TableColumn("date", "Data cupom", 105), TableColumn("posting", "Data Journal", 105),
            TableColumn("pdv", "PDV", 145), TableColumn("document", "Cupom", 90),
            TableColumn("account", "Conta/CHECK", 120), TableColumn("guest", "Quarto / Hóspede", 255),
            TableColumn("pdv_value", "Valor cupom", 120, "e"), TableColumn("journal_value", "Valor conta", 120, "e"),
            TableColumn("difference", "Diferença", 120, "e"), TableColumn("status", "Resultado", 210),
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
        names = filedialog.askopenfilenames(title="Arquivos da conferência de cupons", filetypes=[("Planilhas Excel", "*.xlsx *.xlsm")])
        if not names:
            return
        self.select.configure(state="disabled")
        self.app.set_status("Conferindo cupons e contas dos hóspedes...", .1)
        self.app.run_background(lambda: analyze([Path(name) for name in names]), self._done, self._failed)

    def _done(self, result):
        self.result = result
        self.page = 0
        self.select.configure(state="normal")
        self.export.configure(state="normal")
        self.info.configure(text=f"{result.company} • De/para {result.mapping} • Journal {result.journal_start:%d/%m/%Y} a {result.journal_end:%d/%m/%Y} • {money(result.total_pdv)} em cupons")
        self.app.set_status("Conferência de cupons concluída", 1)
        self._show()

    def _failed(self):
        self.select.configure(state="normal")

    def _filtered(self):
        if not self.result:
            return []
        status, search = self.status_filter.get(), normalize(self.search.get())
        return [item for item in self.result.coupons
                if (status == "Todos" or status == "Conciliados" and item.reconciled or status == "Pendências" and not item.reconciled)
                and (not search or search in normalize(item.pdv + item.document + item.account + item.room + item.guest + item.status + item.detail))]

    def _show(self):
        clear_table(self.preview)
        if not self.result:
            self.page_label.configure(text="Página 0 de 0")
            self.previous.configure(state="disabled")
            self.next.configure(state="disabled")
        else:
            values = (str(len(self.result.coupons)), str(self.result.reconciled), str(self.result.issues), str(self.result.incomplete))
            for title, value in zip(self.cards, values):
                self.cards[title].configure(text=value)
            filtered = self._filtered()
            pages = max(1, (len(filtered) + self.page_size - 1) // self.page_size)
            self.page = min(self.page, pages - 1)
            for item in filtered[self.page * self.page_size:(self.page + 1) * self.page_size]:
                tag = "ok" if item.reconciled else "missing" if item.incomplete_period else "error"
                self.preview.insert("", "end", values=(item.issue_date.strftime("%d/%m/%Y"),
                                    item.posting_date.strftime("%d/%m/%Y") if item.posting_date else "—",
                                    item.pdv, item.document, item.account, f"{item.room} / {item.guest}",
                                    money(item.pdv_value), money(item.journal_value), money(item.difference), item.status), tags=(tag,))
            self.page_label.configure(text=f"Página {self.page + 1} de {pages} • {len(filtered)} cupons")
            self.previous.configure(state="normal" if self.page else "disabled")
            self.next.configure(state="normal" if self.page + 1 < pages else "disabled")
        self.app.after_idle(self._draw_chart)

    def _draw_chart(self):
        self.chart.delete("all")
        width = max(self.chart.winfo_width(), 180)
        if not self.result:
            self.chart.create_text(width / 2, 62, text="Aguardando análise", fill="#9ca3af")
            return
        total = len(self.result.coupons) or 1
        ok_ratio = self.result.reconciled / total
        incomplete_ratio = self.result.incomplete / total
        size, top, left = 70, 24, (width - 70) / 2
        base = "#21a67a" if self.result.reconciled == total else "#dc5a5a"
        self.chart.create_oval(left, top, left + size, top + size, outline=base, width=13)
        if 0 < ok_ratio < 1:
            self.chart.create_arc(left, top, left + size, top + size, start=90, extent=-360 * ok_ratio, style="arc", outline="#21a67a", width=13)
        if incomplete_ratio:
            self.chart.create_arc(left, top, left + size, top + size, start=90 - 360 * ok_ratio,
                                  extent=-360 * incomplete_ratio, style="arc", outline="#e0a83e", width=13)
        self.chart.create_text(width / 2, top + size / 2, text=f"{ok_ratio:.1%}", fill="white", font=("Segoe UI", 11, "bold"))
        self.chart.create_text(width / 2, 108, text=f"{self.result.reconciled} conciliados • {self.result.issues} pendências", fill="#d1d5db", font=("Segoe UI", 9))

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
        name = filedialog.asksaveasfilename(title="Exportar conferência", defaultextension=extension,
                                            initialfile=f"cupons_conta_hospede_{self.result.journal_end:%Y_%m_%d}{extension}",
                                            filetypes=[(self.output_format.get(), f"*{extension}")])
        if not name:
            return
        result = self.result
        task = save_excel if extension == ".xlsx" else save_pdf
        self.app.set_status("Exportando conferência de cupons...", .5)
        self.app.run_background(lambda: task(result, Path(name)),
                                lambda _: (self.app.set_status("Resultado exportado", 1), messagebox.showinfo("Concluído", "Resultado exportado com sucesso.")))

    def _clear(self):
        self.result = None
        self.search.set("")
        self.status_filter.set("Pendências")
        self.page = 0
        self.export.configure(state="disabled")
        self.info.configure(text="Os arquivos são reconhecidos pelas colunas; o nome não é utilizado.")
        self._show()
        self.app.set_status("Seleção limpa", 0)


AUTOMATION_CLASS = GuestCouponAutomation
