from __future__ import annotations

import csv
import re
import unicodedata
import warnings
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from difflib import SequenceMatcher
from html.parser import HTMLParser
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

TOLERANCE = Decimal("0.01")


@dataclass(frozen=True)
class CapNote:
    provider: str
    cnpj: str
    number: str
    emission_date: str
    gross: Decimal
    bpm: str
    hotel: str


@dataclass(frozen=True)
class TaxEntry:
    gross: Decimal
    iss: Decimal


@dataclass(frozen=True)
class ResultRow:
    source: str
    provider: str
    cnpj: str
    number: str
    emission_date: str
    gross: Decimal | None
    iss: Decimal | None
    bpm: str
    cap_provider: str
    cap_date: str
    cap_gross: Decimal | None
    cap_iss: Decimal | None
    cap_hotel: str
    status: str

    @property
    def reconciled(self) -> bool:
        return self.status == "Conciliada"

    @property
    def situation(self) -> str:
        if self.reconciled:
            return "Conciliada"
        if "Não escriturada" in self.status:
            return "Não escriturada"
        if "Ausente" in self.status:
            return "Informação ausente"
        return "Divergente"


@dataclass(frozen=True)
class AnalysisResult:
    rows: list[ResultRow]
    external_count: int
    cap_count: int
    matched_count: int
    approved_count: int
    retained_count: int
    cap_retained_count: int
    expected_hotel: str


def normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return re.sub(r"[^A-Z0-9]", "", "".join(c for c in text if not unicodedata.combining(c)).upper())


def cnpj(value: Any) -> str:
    return re.sub(r"\D", "", str(value or ""))


def note_number(value: Any) -> str:
    text = str(value or "").strip()
    return str(int(float(text))) if re.fullmatch(r"\d+(?:\.0+)?", text) else normalize(text)


def decimal_value(value: Any) -> Decimal:
    if value in (None, "", "NULL"):
        return Decimal()
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return Decimal(str(value).replace(".", "").replace(",", "."))


def date_text(value: Any) -> str:
    if value in (None, "", "NULL"):
        return "—"
    if isinstance(value, (datetime, date)):
        return value.strftime("%d/%m/%Y")
    text = str(value).strip().split()[0]
    for pattern in ("%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, pattern).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return text


def dates_equal(first: str, second: str) -> bool:
    if first == second:
        return True
    try:
        a, b = datetime.strptime(first, "%d/%m/%Y"), datetime.strptime(second, "%d/%m/%Y")
        return a.year == b.year and a.day == b.month and a.month == b.day
    except ValueError:
        return False


def money(value: Decimal) -> str:
    return f"R$ {value:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def xlsx_rows(path: Path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    sheet.reset_dimensions()
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    return rows[0], rows[1:]


def xlsx_columns(path: Path) -> set[str]:
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    sheet.reset_dimensions()
    header = next(sheet.iter_rows(values_only=True), ())
    workbook.close()
    return {normalize(value) for value in header if value not in (None, "")}


def identify_file(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix in {".csv", ".xls"}:
        return "external"
    if suffix != ".xlsx":
        return "unknown"
    columns = xlsx_columns(path)
    if {"RAZAOSOCIALFORNECEDOR", "DOCUMENTOPRINCIPALFORNECEDOR", "NUMERO", "VALORBRUTO", "STATUSBPM"}.issubset(columns):
        return "cap"
    if {"DOCUMENTOPRINCIPALFORNECEDOR", "NUMERODOCUMENTO", "VALORBASECALCULO", "VALOR"}.issubset(columns):
        return "tax"
    if "NUMERONFSE" in columns or {"CNPJ", "PRESTADOR", "VALORSERVICOS"}.issubset(columns):
        return "external"
    return "unknown"


def row_dict(header, row):
    return {name: row[index] if index < len(row) else None for index, name in enumerate(header)}


def read_cap(path: Path) -> list[CapNote]:
    header, rows = xlsx_rows(path)
    result = []
    for row in rows:
        data = row_dict(header, row)
        if data.get("Numero") in (None, ""):
            continue
        result.append(CapNote(
            str(data.get("RazaoSocialFornecedor") or "").strip(), cnpj(data.get("DocumentoPrincipalFornecedor")),
            note_number(data.get("Numero")), date_text(data.get("DataEmissao")), decimal_value(data.get("ValorBruto")),
            str(data.get("StatusBPM") or "Sem BPM").strip(), str(data.get("EmpresaNomeResumido") or "").strip(),
        ))
    return result


def read_tax(path: Path) -> dict[tuple[str, str], TaxEntry]:
    header, rows = xlsx_rows(path)
    result = {}
    for row in rows:
        data = row_dict(header, row)
        key = cnpj(data.get("DocumentoPrincipalFornecedor")), note_number(data.get("NumeroDocumento"))
        if all(key):
            result[key] = TaxEntry(decimal_value(data.get("ValorBaseCalculo")), abs(decimal_value(data.get("Valor"))))
    return result


def external_xlsx(path: Path) -> list[dict[str, Any]]:
    header, rows = xlsx_rows(path)
    result = []
    if "Número NFS-e" in header:
        for row in rows:
            data = row_dict(header, row)
            if data.get("Número NFS-e") in (None, ""):
                continue
            retained = str(data.get("Retenção ISSQN") or "").startswith("2")
            result.append({"source": "Portal Nacional", "number": data["Número NFS-e"],
                           "date": data.get("Data Geração"), "cnpj": data.get("CNPJ/CPF Prestador"),
                           "provider": data.get("Nome Prestador"), "gross": data.get("Valor do Serviço (R$)"),
                           "iss": data.get("Valor do ISSQN (R$)") if retained else 0})
    elif "N°" in header:
        for row in rows:
            data = row_dict(header, row)
            if data.get("N°") in (None, ""):
                continue
            result.append({"source": "Prefeitura Caucaia", "number": data["N°"], "date": data.get("DATA"),
                           "cnpj": data.get("CNPJ"), "provider": data.get("PRESTADOR"),
                           "gross": data.get("Valor Serviços"), "iss": data.get("Valor ISS")})
    return result


def external_csv(path: Path) -> list[dict[str, Any]]:
    text = path.read_text(encoding="latin1")
    result = []
    for data in csv.DictReader(text.splitlines(), delimiter=";"):
        if not data.get("Nº NFS-e") or data.get("Nº NFS-e") == "Total":
            continue
        retained = str(data.get("ISS Retido") or "").strip().upper() == "S"
        result.append({"source": "Prefeitura SP", "number": data.get("Nº NFS-e"),
                       "date": data.get("Data Hora NFE"), "cnpj": data.get("CPF/CNPJ do Prestador"),
                       "provider": data.get("Razão Social do Prestador"), "gross": data.get("Valor dos Serviços"),
                       "iss": data.get("ISS devido") if retained else 0})
    return result


class HtmlTableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows: list[list[str]] = []
        self._row: list[str] | None = None
        self._cell: list[str] | None = None

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() == "tr":
            self._row = []
        elif tag.lower() in {"th", "td"} and self._row is not None:
            self._cell = []

    def handle_data(self, data: str) -> None:
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"th", "td"} and self._row is not None and self._cell is not None:
            self._row.append(" ".join("".join(self._cell).split()))
            self._cell = None
        elif tag.lower() == "tr" and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None


def external_html_xls(path: Path) -> list[dict[str, Any]]:
    raw = path.read_bytes()
    if not raw.lstrip().lower().startswith((b"<html", b"<!doctype")):
        raise ValueError(f"O arquivo {path.name} não é um relatório HTML compatível.")
    parser = HtmlTableParser()
    parser.feed(raw.decode("utf-8-sig", errors="replace"))
    if len(parser.rows) < 2:
        raise ValueError(f"Nenhuma nota foi encontrada em {path.name}.")

    header = {normalize(value): index for index, value in enumerate(parser.rows[0])}

    def value(row: list[str], *names: str) -> str:
        index = next((header[name] for name in names if name in header), None)
        return row[index] if index is not None and index < len(row) else ""

    required_groups = (
        ("NUMERO", "NUMERONFSE", "NUMERODANOTA", "NOTA"),
        ("PRESTADORDOSERVICO", "PRESTADOR", "RAZAOSOCIAL"),
        ("DATADEEMISSAO", "DATAEMISSAO", "DATA"),
        ("VALORDOSERVICO", "VALORBRUTO", "VALOR"),
    )
    if any(not any(name in header for name in group) for group in required_groups):
        raise ValueError(f"Colunas obrigatórias não encontradas em {path.name}.")

    result = []
    for row in parser.rows[1:]:
        number = value(row, "NUMERO", "NUMERONFSE", "NUMERODANOTA", "NOTA")
        provider_field = value(row, "PRESTADORDOSERVICO", "PRESTADOR", "RAZAOSOCIAL")
        match = re.match(r"\s*([\d./-]{14,})\s+-\s+(.+)", provider_field)
        document = match.group(1) if match else value(row, "CNPJ", "CPFCNPJ", "CNPJCPF")
        provider = match.group(2) if match else provider_field
        if not number or not cnpj(document) or not provider:
            continue
        provider = re.sub(r"^\s*[\d./-]{8,}\s+", "", provider).strip()
        result.append({
            "source": "Relatório externo",
            "number": number,
            "date": value(row, "DATADEEMISSAO", "DATAEMISSAO", "DATA"),
            "cnpj": document,
            "provider": provider,
            "gross": value(row, "VALORDOSERVICO", "VALORBRUTO", "VALOR"),
            "iss": value(row, "ISSDEVIDO", "VALORDOISS", "VALORISS", "ISSRETIDO"),
        })
    if not result:
        raise ValueError(f"Nenhuma nota com CNPJ e prestador foi encontrada em {path.name}.")
    return result


def expected_hotel(paths: list[Path]) -> str:
    tokens = [match.group(1) for path in paths if (match := re.search(r"\(([^)]+)\)", path.name))]
    return normalize(max(set(tokens), key=tokens.count)) if tokens else ""


def analyze(paths: list[Path]) -> AnalysisResult:
    if len(paths) < 3:
        raise ValueError("Selecione o CAP, o Alterador ISS e pelo menos uma fonte externa.")
    identified = [(path, identify_file(path)) for path in paths]
    cap_paths = [path for path, kind in identified if kind == "cap"]
    tax_paths = [path for path, kind in identified if kind == "tax"]
    external_paths = [path for path, kind in identified if kind == "external"]
    unknown = [path.name for path, kind in identified if kind == "unknown"]
    if unknown:
        raise ValueError(f"Formato não reconhecido: {', '.join(unknown)}.")
    if len(cap_paths) != 1 or len(tax_paths) != 1 or not external_paths:
        raise ValueError("Envie um arquivo CAP, um arquivo de ISS retido e pelo menos uma fonte externa.")
    cap_notes, taxes = read_cap(cap_paths[0]), read_tax(tax_paths[0])
    hotel = expected_hotel(paths)
    if not hotel and cap_notes:
        hotel = normalize(max({item.hotel for item in cap_notes}, key=lambda name: sum(row.hotel == name for row in cap_notes)))
    cap_map = {(item.cnpj, item.number): item for item in cap_notes}
    external = []
    for path in external_paths:
        suffix = path.suffix.lower()
        if suffix == ".csv":
            external.extend(external_csv(path))
        elif suffix == ".xls":
            external.extend(external_html_xls(path))
        else:
            external.extend(external_xlsx(path))

    grouped_external = defaultdict(list)
    for raw in external:
        grouped_external[(cnpj(raw["cnpj"]), note_number(raw["number"]))].append(raw)

    result = []
    retained_count = 0
    for key, occurrences in grouped_external.items():
        raw = occurrences[0]
        sources = " + ".join(sorted({str(item["source"]) for item in occurrences}))
        provider, issued, gross = str(raw["provider"] or "").strip(), date_text(raw["date"]), decimal_value(raw["gross"])
        gross_values = [decimal_value(item["gross"]) for item in occurrences]
        dates = [date_text(item["date"]) for item in occurrences]
        prefeitura = [item for item in occurrences if str(item["source"]) != "Portal Nacional"]
        iss = max((decimal_value(item["iss"]) for item in prefeitura), default=Decimal()) if prefeitura else None
        cap = cap_map.get(key)
        tax = taxes.get(key)
        issues = []
        if max(gross_values) - min(gross_values) > TOLERANCE:
            issues.append("Valor divergente entre fontes externas")
        if any(not dates_equal(dates[0], item) for item in dates[1:]):
            issues.append("Data divergente entre fontes externas")
        bpm = "—"
        if not cap:
            issues.append("Ausente no CAP")
        else:
            sources = f"CAP + {sources}"
            bpm = cap.bpm
            if normalize(cap.bpm) != "BMAPROVADO" and normalize(cap.bpm) != "BPMAPROVADO":
                issues.append("Não escriturada: BPM não aprovado")
            if hotel and hotel not in normalize(cap.hotel):
                issues.append(f"Hotel divergente ({cap.hotel})")
            if abs(gross - cap.gross) > TOLERANCE:
                issues.append("Valor bruto divergente")
            if not dates_equal(issued, cap.emission_date):
                issues.append("Data divergente")
            similarity = SequenceMatcher(None, normalize(provider), normalize(cap.provider)).ratio()
            if similarity < .60:
                issues.append("Razão social divergente")
        if iss is not None and iss > 0:
            retained_count += 1
            if not tax:
                issues.append("ISS retido ausente no CAP")
            elif abs(iss - tax.iss) > TOLERANCE:
                issues.append("ISS retido divergente")
        elif tax and tax.iss > 0:
            issues.append("ISS retido ausente na prefeitura")
        result.append(ResultRow(
            sources, provider, key[0], key[1], issued, gross, iss, bpm,
            cap.provider if cap else "—", cap.emission_date if cap else "—", cap.gross if cap else None,
            tax.iss if tax else None, cap.hotel if cap else "—", "Conciliada" if not issues else " • ".join(issues),
        ))
    for key, cap in cap_map.items():
        if key in grouped_external:
            continue
        tax = taxes.get(key)
        issues = ["Ausente nas fontes externas"]
        if tax and tax.iss > 0:
            issues.append("ISS retido ausente na prefeitura")
        if normalize(cap.bpm) != "BPMAPROVADO":
            issues.append("Não escriturada: BPM não aprovado")
        if hotel and hotel not in normalize(cap.hotel):
            issues.append(f"Hotel divergente ({cap.hotel})")
        result.append(ResultRow(
            "CAP", "—", cap.cnpj, cap.number, "—", None, None, cap.bpm, cap.provider,
            cap.emission_date, cap.gross, tax.iss if tax else None, cap.hotel, " • ".join(issues),
        ))
    result.sort(key=lambda row: (row.reconciled, row.source, row.provider, row.number))
    return AnalysisResult(result, len(grouped_external), len(cap_notes),
                          sum(row.cap_gross is not None and row.source != "CAP" for row in result),
                          sum(normalize(x.bpm) == "BPMAPROVADO" for x in cap_notes), retained_count,
                          sum(entry.iss > 0 for entry in taxes.values()), hotel)


def save_excel(result: AnalysisResult, path: Path) -> None:
    workbook = Workbook()
    summary = workbook.active
    summary.title = "Resumo"
    reconciled = sum(row.reconciled for row in result.rows)
    summary.append(["Indicador", "Resultado"])
    for label, value in (
        ("Hotel esperado", result.expected_hotel), ("Notas externas únicas", result.external_count),
        ("Notas existentes no arquivo CAP", result.cap_count), ("Notas externas encontradas no CAP", result.matched_count),
        ("Notas externas ausentes no CAP", result.external_count - result.matched_count),
        ("Notas existentes somente no CAP", result.cap_count - result.matched_count),
        ("BPM aprovadas no arquivo CAP", result.approved_count),
        ("Notas com ISS retido na prefeitura", result.retained_count),
        ("Notas com ISS retido no CAP", result.cap_retained_count),
        ("Totalmente conciliadas", reconciled), ("Com pendências", len(result.rows) - reconciled),
    ):
        summary.append([label, value])
    for cell in summary[1]: cell.style = "Headline 4"
    summary.column_dimensions["A"].width = 40; summary.column_dimensions["B"].width = 24

    sheet = workbook.create_sheet("Comparação detalhada")
    sheet.append(["Fonte", "CNPJ", "Número da nota", "Prestador externo", "Data externa", "Valor externo",
                  "ISS prefeitura", "Prestador CAP", "Data CAP", "Valor CAP", "ISS CAP", "BPM", "Hotel CAP",
                  "Situação", "Detalhes"])
    for row in result.rows:
        sheet.append([row.source, row.cnpj, row.number, row.provider, row.emission_date,
                      float(row.gross) if row.gross is not None else None,
                      float(row.iss) if row.iss is not None else None,
                      row.cap_provider, row.cap_date, float(row.cap_gross) if row.cap_gross is not None else None,
                      float(row.cap_iss) if row.cap_iss is not None else None, row.bpm, row.cap_hotel,
                      row.situation, row.status])
    for cell in sheet[1]: cell.style = "Headline 4"
    for column in ("F", "G", "J", "K"):
        for cell in sheet[column][1:]: cell.number_format = 'R$ #,##0.00'
    for column, width in {"A":30,"B":20,"C":18,"D":48,"E":14,"F":18,"G":18,"H":48,"I":14,
                          "J":18,"K":18,"L":20,"M":22,"N":22,"O":70}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"; sheet.auto_filter.ref = sheet.dimensions; workbook.save(path)


def save_pdf(result: AnalysisResult, path: Path) -> None:
    styles = getSampleStyleSheet(); total=len(result.rows); ok=sum(r.reconciled for r in result.rows)
    data=[["Notas analisadas","Encontradas no CAP","BPM aprovadas","Conciliadas","Pendências","ISS Pref./CAP"],
          [str(total),str(result.matched_count),str(result.approved_count),str(ok),str(total-ok),
           f"{result.retained_count} / {result.cap_retained_count}"]]
    table=Table(data,colWidths=[38*mm]*6); table.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor("#24588A")),("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),("ALIGN",(0,0),(-1,-1),"CENTER"),("GRID",(0,0),(-1,-1),.5,colors.grey)]))
    doc=SimpleDocTemplate(str(path),pagesize=landscape(A4),title="Conferência de Notas de Serviços Tomados")
    doc.build([Paragraph("Conferência de Notas de Serviços Tomados",styles["Title"]),Spacer(1,5*mm),table])


class ServiceNotesAutomation(Automation):
    name = "Conferência de Notas de Serviços Tomados"

    def __init__(self, app, container):
        super().__init__(app, container); self.paths=[]; self.result=None; self.output_format=ctk.StringVar(value="Excel")
        self.filter_status=ctk.StringVar(value="Pendências"); self.filter_source=ctk.StringVar(value="Todas as fontes"); self.search=ctk.StringVar()
        self.page=0; self.page_size=25

    def render(self):
        self.container.grid_columnconfigure(0,weight=1); self.container.grid_rowconfigure(6,weight=1)
        ctk.CTkLabel(self.container,text=self.name,font=ctk.CTkFont(size=26,weight="bold")).grid(row=0,column=0,padx=30,pady=(22,3),sticky="w")
        ctk.CTkLabel(self.container,text="Compara notas das prefeituras com CAP, hotel, BPM e ISS retido.",text_color="gray70").grid(row=1,column=0,padx=30,pady=(0,10),sticky="w")
        controls=ctk.CTkFrame(self.container,fg_color="transparent");controls.grid(row=2,column=0,padx=30,sticky="ew")
        self.select=ctk.CTkButton(controls,text="Selecionar arquivos",command=self._select);self.select.pack(side="left",padx=(0,10))
        self.format=ctk.CTkSegmentedButton(controls,values=["Excel","PDF"],variable=self.output_format);self.format.pack(side="left",padx=10)
        self.export=ctk.CTkButton(controls,text="Exportar resultado",state="disabled",command=self._export);self.export.pack(side="left",padx=10)
        self.clear=ctk.CTkButton(controls,text="Limpar",fg_color="gray35",command=self._clear);self.clear.pack(side="left",padx=10)
        self.info=ctk.CTkLabel(self.container,text="Obrigatórios: CAP, Alterador ISS e uma ou mais fontes externas.",text_color="gray70",anchor="w");self.info.grid(row=3,column=0,padx=30,pady=(10,6),sticky="ew")
        dash=ctk.CTkFrame(self.container,fg_color="transparent");dash.grid(row=4,column=0,padx=30,pady=(0,8),sticky="ew");dash.grid_columnconfigure(0,weight=3);dash.grid_columnconfigure(1,weight=2)
        cards=ctk.CTkFrame(dash,fg_color="transparent");cards.grid(row=0,column=0,padx=(0,10),sticky="nsew");self.cards={}
        for index,title in enumerate(("Notas externas","Encontradas no CAP","BPM aprovadas no CAP","ISS Prefeitura / CAP")):
            cards.grid_columnconfigure(index,weight=1);card=ctk.CTkFrame(cards);card.grid(row=0,column=index,padx=(0 if index==0 else 5,0),sticky="nsew")
            ctk.CTkLabel(card,text=title,text_color="gray70").pack(pady=(8,0));label=ctk.CTkLabel(card,text="—",font=ctk.CTkFont(size=16,weight="bold"));label.pack(pady=(0,8));self.cards[title]=label
        chart_frame=ctk.CTkFrame(dash);chart_frame.grid(row=0,column=1,sticky="nsew");ctk.CTkLabel(chart_frame,text="Distribuição da conferência",font=ctk.CTkFont(size=13,weight="bold")).pack(pady=(6,0))
        self.chart=tk.Canvas(chart_frame,height=125,bg="#2b2b2b",highlightthickness=0);self.chart.pack(fill="x",padx=8);self.chart.bind("<Configure>",lambda _:self._chart())
        filters=ctk.CTkFrame(self.container,fg_color="transparent");filters.grid(row=5,column=0,padx=30,pady=(0,8),sticky="ew");filters.grid_columnconfigure(3,weight=1)
        self.status=ctk.CTkSegmentedButton(filters,values=["Pendências","Conciliadas","Todas"],variable=self.filter_status,command=lambda _:self._reset());self.status.grid(row=0,column=0,padx=(0,12))
        self.source=ctk.CTkOptionMenu(filters,values=["Todas as fontes"],variable=self.filter_source,command=lambda _:self._reset(),width=180);self.source.grid(row=0,column=1,padx=(0,12))
        ctk.CTkLabel(filters,text="Buscar:").grid(row=0,column=2,padx=(0,8));entry=ctk.CTkEntry(filters,textvariable=self.search,placeholder_text="CNPJ, prestador ou nota");entry.grid(row=0,column=3,sticky="ew");entry.bind("<KeyRelease>",lambda _:self._reset())
        ctk.CTkLabel(filters,text="Verde: conciliada  •  Amarelo: informação ausente  •  Vermelho: divergência ou não escriturada  •  Célula vazia: dado não encontrado na fonte",text_color="gray70",anchor="w").grid(row=1,column=0,columnspan=4,pady=(6,0),sticky="ew")
        self.preview=create_result_table(self.container,(
            TableColumn("source","Fonte",190),TableColumn("provider","Prestador",300),TableColumn("cnpj","CNPJ",125),
            TableColumn("note","Nota",95),TableColumn("date","Data",90),TableColumn("external","Valor externo",115),
            TableColumn("cap","Valor CAP",115),TableColumn("iss_external","ISS prefeitura",105),
            TableColumn("iss_cap","ISS CAP",105),TableColumn("bpm","BPM",130),TableColumn("hotel","Hotel CAP",140),
            TableColumn("situation","Situação",135),TableColumn("result","Detalhes",430),
        ),row=6)
        pagination=ctk.CTkFrame(self.container,fg_color="transparent");pagination.grid(row=7,column=0,padx=30,pady=(7,14),sticky="ew");pagination.grid_columnconfigure(1,weight=1)
        self.previous=ctk.CTkButton(pagination,text="Anterior",width=100,command=self._previous);self.previous.grid(row=0,column=0)
        self.page_label=ctk.CTkLabel(pagination,text="Página 0 de 0");self.page_label.grid(row=0,column=1)
        self.next=ctk.CTkButton(pagination,text="Próxima",width=100,command=self._next);self.next.grid(row=0,column=2);self._show()

    def _select(self):
        names=filedialog.askopenfilenames(title="Arquivos da Atividade 7",filetypes=[("Planilhas e CSV","*.xlsx *.xls *.csv")])
        if not names:return
        self.paths=[Path(x) for x in names];self.select.configure(state="disabled");self.app.set_status("Conferindo notas tomadas...",.1);self.app.run_background(lambda:analyze(self.paths),self._done,self._failed)

    def _done(self,result):
        self.result=result;self.page=0;self.select.configure(state="normal");self.export.configure(state="normal");self._update_source_filter();self.info.configure(text=f"Arquivos carregados: {len(self.paths)} • Hotel esperado: {result.expected_hotel} • {result.external_count} notas externas • {result.cap_count} no CAP");self.app.set_status("Conferência concluída",1);self._show()

    def _failed(self):self.select.configure(state="normal")

    def _update_source_filter(self):
        sources = sorted({source for row in self.result.rows for source in row.source.split(" + ")}) if self.result else []
        self.filter_source.set("Todas as fontes")
        self.source.configure(values=["Todas as fontes", *sources])

    def _filtered(self):
        if not self.result:return []
        text=normalize(self.search.get());status=self.filter_status.get();source=self.filter_source.get()
        return [r for r in self.result.rows if (status=="Todas" or status=="Conciliadas" and r.reconciled or status=="Pendências" and not r.reconciled) and (source=="Todas as fontes" or source in r.source) and (not text or text in normalize(r.provider+r.cnpj+r.number))]

    def _show(self):
        clear_table(self.preview)
        if not self.result:
            self.page_label.configure(text="Página 0 de 0");self.previous.configure(state="disabled");self.next.configure(state="disabled")
        else:
            filtered=self._filtered();total_pages=max(1,(len(filtered)+self.page_size-1)//self.page_size);self.page=min(self.page,total_pages-1);rows=filtered[self.page*self.page_size:(self.page+1)*self.page_size]
            values=(self.result.external_count,self.result.matched_count,self.result.approved_count,
                    f"{self.result.retained_count} / {self.result.cap_retained_count}")
            for title,value in zip(self.cards,values):
                self.cards[title].configure(text=value if isinstance(value,str) else f"{value:,}".replace(",","."))
            for r in rows:
                tag="ok" if r.reconciled else "missing" if r.situation=="Informação ausente" else "error"
                display_date=r.emission_date if r.emission_date!="—" else r.cap_date if r.cap_date!="—" else ""
                self.preview.insert("", "end", values=(r.source,r.provider if r.provider != "—" else r.cap_provider,
                                    r.cnpj,r.number,display_date,money(r.gross) if r.gross is not None else "",
                                    money(r.cap_gross) if r.cap_gross is not None else "",
                                    money(r.iss) if r.iss is not None else "",
                                    money(r.cap_iss) if r.cap_iss is not None else "",
                                    "" if r.bpm=="—" else r.bpm,"" if r.cap_hotel=="—" else r.cap_hotel,
                                    r.situation,r.status),tags=(tag,))
            self.page_label.configure(text=f"Página {self.page+1} de {total_pages} • {len(filtered)} notas");self.previous.configure(state="normal" if self.page else "disabled");self.next.configure(state="normal" if self.page+1<total_pages else "disabled")
        self.app.after_idle(self._chart)

    def _reset(self):self.page=0;self._show()

    def _previous(self):
        if self.page:self.page-=1;self._show()

    def _next(self):
        if (self.page+1)*self.page_size<len(self._filtered()):self.page+=1;self._show()

    def _chart(self):
        self.chart.delete("all");w=max(self.chart.winfo_width(),300);cx,cy,r=w*.32,55,43;b=(cx-r,cy-r,cx+r,cy+r)
        if not self.result or not self.result.rows:self.chart.create_oval(*b,outline="#4b5563",width=13);return
        total=len(self.result.rows);ok=sum(x.reconciled for x in self.result.rows);bad=total-ok;rate=ok/total;split=-360*rate
        self.chart.create_arc(*b,start=90,extent=split,style="arc",outline="#21a67a",width=13);self.chart.create_arc(*b,start=90+split,extent=-360*(bad/total),style="arc",outline="#dc5a5a",width=13);self.chart.create_text(cx,cy,text=f"{rate:.1%}",fill="white",font=("Segoe UI",10,"bold"));self.chart.create_text(w*.57,40,text=f"● Conciliadas: {ok}",fill="#21a67a",anchor="w");self.chart.create_text(w*.57,67,text=f"● Pendências: {bad}",fill="#dc5a5a",anchor="w")

    def _export(self):
        if not self.result:return
        ext=".pdf" if self.output_format.get()=="PDF" else ".xlsx";name=filedialog.asksaveasfilename(defaultextension=ext,initialfile=f"conferencia_notas_servicos_tomados{ext}")
        if not name:return
        task=lambda:save_pdf(self.result,Path(name)) if ext==".pdf" else save_excel(self.result,Path(name));self.app.run_background(task,lambda _:messagebox.showinfo("Exportação concluída",f"Arquivo salvo em:\n{name}"),self._failed)

    def _clear(self):self.paths=[];self.result=None;self.search.set("");self.page=0;self.export.configure(state="disabled");self._update_source_filter();self.info.configure(text="Obrigatórios: CAP, Alterador ISS e uma ou mais fontes externas.");self._show();self.app.set_status("Seleção limpa",0)


AUTOMATION_CLASS = ServiceNotesAutomation
