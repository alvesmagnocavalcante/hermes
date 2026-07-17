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
ENTRY_ACCOUNTS = {
    "ALIMENTOS": "Alimentos", "VINHOECHAMPANHE": "Vinhos & Champanhe",
    "BEBIDASALCOOLICAS": "Alcoolicos", "BEBIDASNAOALCOOLICAS": "Não Alcoolicos", "FRIGOBAR": "Frigobar",
}
INVENTORY_CODES = {
    "Alimentos": ("01",), "Vinhos & Champanhe": ("02",), "Alcoolicos": ("03",),
    "Não Alcoolicos": ("04",), "Frigobar": ("05",), "Mimos Hospedes": ("06", "0706"),
    "Amenitees": ("0701",), "Material de Higiene e Limpeza": ("0702",),
    "Material de Escritório/Informatica": ("0704", "0711"), "Decoracao": ("0708",),
    "Eletroeletronicos": ("0709",), "Suprimentos de uso do Hospedes": ("0713", "0806"),
    "Material de Copa e Cozinha": ("0802", "0804", "2001"), "Uniforme": ("0805",),
    "Material de Manutenção de Edifícios e Instalações": ("0901", "0902", "0904", "0909"),
    "Material de Manutenção de Maquinas e Equipamentos": ("0908", "0910"),
    "Material de Manutenção da Piscina": ("0903",), "Materal de Reposicao": ("0705",),
    "Equipamento de Protecao": ("0907",), "SPA": ("10",),
}


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


def normalize(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return re.sub(r"[^A-Z0-9]", "", "".join(c for c in text if not unicodedata.combining(c)).upper())


def number(value: Any) -> Decimal:
    if value in (None, "", "NULL"):
        return Decimal()
    try:
        return Decimal(str(value))
    except InvalidOperation:
        return Decimal(str(value).replace(".", "").replace(",", "."))


def money(value: Decimal) -> str:
    return f"R$ {value:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def read(path: Path):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        workbook = load_workbook(path, data_only=True, read_only=True)
    sheet = workbook.active
    sheet.reset_dimensions()
    rows = list(sheet.iter_rows(values_only=True))
    workbook.close()
    return rows[0], rows[1:]


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


def analyze(paths: list[Path]) -> list[Row]:
    if len(paths) != 4:
        raise ValueError("Selecione exatamente os quatro arquivos da Atividade 10.")
    files = identify(paths)
    documents_raw = grouped(files["documents"], "DESCRICAOTDESEMB", "VALORLANÇADO")
    documents = {normalize(key): value for key, value in documents_raw.items()}
    entry_ledger = grouped(files["entry_ledger"], "DescricaoConta", "Debito")
    entry_normalized = {normalize(key): value for key, value in entry_ledger.items()}
    result = [Row("Entradas", account, documents.get(key, Decimal()), entry_normalized.get(normalize(account), Decimal()))
              for key, account in ENTRY_ACCOUNTS.items()]

    inventory = grouped(files["inventory"], "GrupoCodigo", "SaldoValor")
    header, ledger_rows = files["stock_ledger"]
    name_i, balance_i = header.index("DescricaoConta"), header.index("SaldoAtual")
    final_balances = {}
    for row in ledger_rows:
        if len(row) > max(name_i, balance_i) and row[name_i] not in (None, "", "NULL"):
            final_balances[str(row[name_i]).strip()] = number(row[balance_i])
    for account, codes in INVENTORY_CODES.items():
        inventory_value = sum((value for code, value in inventory.items() if any(str(code).startswith(prefix) for prefix in codes)), Decimal())
        ledger_value = next((value for name, value in final_balances.items() if normalize(name) == normalize(account)), Decimal())
        result.append(Row("Saldo final", account, inventory_value, ledger_value))
    return result


def export_excel(rows: list[Row], path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Conferência de Custos"
    sheet.append(["Análise", "Conta", "CAP / Inventário", "Contabilidade", "Diferença", "Status"])
    for row in rows:
        sheet.append([row.analysis, row.account, float(row.source), float(row.accounting), float(row.difference), row.status])
    for cell in sheet[1]:
        cell.style = "Headline 4"
    for column in "CDE":
        for cell in sheet[column][1:]:
            cell.number_format = 'R$ #,##0.00'
    for column, width in {"A": 18, "B": 52, "C": 22, "D": 22, "E": 20, "F": 16}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    workbook.save(path)


def export_pdf(rows: list[Row], path: Path) -> None:
    styles = getSampleStyleSheet()
    data = [["Análise", "Conta", "CAP / Inventário", "Contabilidade", "Diferença", "Status"]]
    data.extend([[r.analysis, r.account, money(r.source), money(r.accounting), money(r.difference), r.status] for r in rows])
    table = Table(data, repeatRows=1, colWidths=[27*mm, 64*mm, 38*mm, 38*mm, 36*mm, 28*mm])
    table.setStyle(TableStyle([("BACKGROUND", (0,0), (-1,0), colors.HexColor("#24588A")),
                               ("TEXTCOLOR", (0,0), (-1,0), colors.white), ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
                               ("FONTSIZE", (0,0), (-1,-1), 7), ("GRID", (0,0), (-1,-1), .4, colors.grey)]))
    doc = SimpleDocTemplate(str(path), pagesize=landscape(A4), title="Conferência dos Custos da Mercadoria Vendida")
    doc.build([Paragraph("Conferência dos Custos da Mercadoria Vendida", styles["Title"]), Spacer(1, 5*mm), table])


class CostAutomation(Automation):
    name = "Conferência dos Custos da Mercadoria Vendida"

    def __init__(self, app, container):
        super().__init__(app, container)
        self.paths = []
        self.rows = []
        self.analysis = ctk.StringVar(value="Entradas")
        self.output_format = ctk.StringVar(value="Excel")

    def render(self):
        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(6, weight=1)
        ctk.CTkLabel(self.container, text=self.name, font=ctk.CTkFont(size=26, weight="bold")).grid(row=0,column=0,padx=30,pady=(22,3),sticky="w")
        ctk.CTkLabel(self.container, text="Confere valores lançados no CAP e o saldo final do inventário contra a contabilidade.", text_color="gray70").grid(row=1,column=0,padx=30,pady=(0,10),sticky="w")
        controls = ctk.CTkFrame(self.container, fg_color="transparent"); controls.grid(row=2,column=0,padx=30,sticky="ew")
        self.select = ctk.CTkButton(controls,text="Selecionar os quatro arquivos",command=self._select); self.select.pack(side="left",padx=(0,10))
        self.format = ctk.CTkSegmentedButton(controls,values=["Excel","PDF"],variable=self.output_format); self.format.pack(side="left",padx=10)
        self.export = ctk.CTkButton(controls,text="Exportar resultado",state="disabled",command=self._export); self.export.pack(side="left",padx=10)
        ctk.CTkButton(controls,text="Limpar",fg_color="gray35",command=self._clear).pack(side="left",padx=10)
        self.info = ctk.CTkLabel(self.container,text="Selecione os quatro arquivos da atividade 10.",text_color="gray70",anchor="w"); self.info.grid(row=3,column=0,padx=30,pady=(10,6),sticky="ew")

        dashboard=ctk.CTkFrame(self.container,fg_color="transparent"); dashboard.grid(row=4,column=0,padx=30,pady=(0,8),sticky="ew")
        dashboard.grid_columnconfigure(0,weight=3); dashboard.grid_columnconfigure(1,weight=2)
        cards=ctk.CTkFrame(dashboard,fg_color="transparent"); cards.grid(row=0,column=0,padx=(0,10),sticky="nsew")
        self.summary_labels={}
        for column,title in enumerate(("Entradas","Saldo final")):
            cards.grid_columnconfigure(column,weight=1)
            card=ctk.CTkFrame(cards); card.grid(row=0,column=column,padx=(0 if column==0 else 7,0),sticky="nsew")
            ctk.CTkLabel(card,text=title,font=ctk.CTkFont(size=14,weight="bold")).pack(pady=(9,2))
            counts=ctk.CTkLabel(card,text="—",text_color="gray70",justify="center"); counts.pack()
            status=ctk.CTkLabel(card,text="Aguardando",font=ctk.CTkFont(weight="bold")); status.pack(pady=(2,9))
            self.summary_labels[title]=(counts,status)
        chart_frame=ctk.CTkFrame(dashboard); chart_frame.grid(row=0,column=1,sticky="nsew"); chart_frame.grid_columnconfigure(0,weight=1)
        ctk.CTkLabel(chart_frame,text="Distribuição da conciliação",font=ctk.CTkFont(size=13,weight="bold")).grid(row=0,column=0,pady=(6,0),sticky="ew")
        self.chart=tk.Canvas(chart_frame,height=125,bg="#2b2b2b",highlightthickness=0); self.chart.grid(row=1,column=0,padx=8,sticky="ew"); self.chart.bind("<Configure>",lambda _:self._chart())
        tabs = ctk.CTkSegmentedButton(self.container,values=["Entradas","Saldo final"],variable=self.analysis,command=lambda _:self._show()); tabs.grid(row=5,column=0,padx=30,pady=(0,8),sticky="w")
        self.preview = create_result_table(self.container, (
            TableColumn("account", "Conta", 430), TableColumn("source", "CAP / Inventário", 160, "e"),
            TableColumn("accounting", "Contabilidade", 160, "e"),
            TableColumn("difference", "Diferença", 150, "e"), TableColumn("status", "Situação", 130),
        ), row=6, pady=(0,14))
        self._show()

    def _select(self):
        names=filedialog.askopenfilenames(title="Arquivos da Atividade 10",filetypes=[("Planilhas Excel","*.xlsx")])
        if not names:return
        self.paths=[Path(x) for x in names]; self.select.configure(state="disabled"); self.app.set_status("Conferindo custos...",.1)
        self.app.run_background(lambda:analyze(self.paths),self._done,self._failed)

    def _done(self, rows):
        self.rows=rows; self.select.configure(state="normal"); self.export.configure(state="normal"); self.info.configure(text=f"Arquivos carregados: {len(self.paths)} • Contas analisadas: {len(rows)}"); self.app.set_status("Conferência concluída",1); self._show()

    def _failed(self): self.select.configure(state="normal")

    def _show(self):
        clear_table(self.preview)
        shown=[r for r in self.rows if r.analysis==self.analysis.get()]
        for row in shown:
            self.preview.insert("", "end", values=(row.account, money(row.source), money(row.accounting),
                money(row.difference), row.status), tags=(result_tag(row.status),))
        for title,(counts,status) in self.summary_labels.items():
            group=[r for r in self.rows if r.analysis==title]
            source_total=sum((r.source for r in group),Decimal()); accounting_total=sum((r.accounting for r in group),Decimal())
            difference=source_total-accounting_total; reconciled=bool(group) and abs(difference)<=TOLERANCE
            source_name="CAP" if title=="Entradas" else "Inventário"
            counts.configure(text=f"{source_name}: {money(source_total)}\nContabilidade: {money(accounting_total)}" if group else "—")
            status.configure(text=(f"{'Conciliado' if reconciled else 'Divergente'} • diferença {money(difference)}" if group else "Aguardando"),
                             text_color="#21a67a" if reconciled else "#dc5a5a" if group else "gray70")
        self.app.after_idle(self._chart)

    def _chart(self):
        self.chart.delete("all"); width=max(self.chart.winfo_width(),300); center_x,center_y,radius=width*.34,53,43; bounds=(center_x-radius,center_y-radius,center_x+radius,center_y+radius)
        total=len(self.rows)
        if not total:
            self.chart.create_oval(*bounds,outline="#4b5563",width=13); self.chart.create_text(center_x,center_y,text="—",fill="#e5e7eb"); self.chart.create_text(width*.58,53,text="Aguardando conferência",fill="#cbd5e1",anchor="w"); return
        ok=sum(r.status=="Conciliado" for r in self.rows); bad=total-ok; rate=ok/total; split=-360*rate
        self.chart.create_arc(*bounds,start=90,extent=split,style="arc",outline="#21a67a",width=13)
        self.chart.create_arc(*bounds,start=90+split,extent=-360*(bad/total),style="arc",outline="#dc5a5a",width=13)
        self.chart.create_text(center_x,center_y,text=f"{rate:.1%}",fill="#e5e7eb",font=("Segoe UI",10,"bold"))
        self.chart.create_text(width*.58,40,text=f"● Conciliadas: {ok}",fill="#21a67a",anchor="w")
        self.chart.create_text(width*.58,67,text=f"● Divergentes: {bad}",fill="#dc5a5a",anchor="w")

    def _clear(self):
        self.paths=[]; self.rows=[]; self.info.configure(text="Selecione os quatro arquivos da atividade 10."); self.export.configure(state="disabled"); self._show(); self.app.set_status("Seleção limpa",0)

    def _export(self):
        ext=".pdf" if self.output_format.get()=="PDF" else ".xlsx"; name=filedialog.asksaveasfilename(defaultextension=ext,initialfile=f"conferencia_custos_mercadoria{ext}")
        if not name:return
        task=lambda:export_pdf(self.rows,Path(name)) if ext==".pdf" else export_excel(self.rows,Path(name))
        self.app.run_background(task,lambda _:messagebox.showinfo("Exportação concluída",f"Arquivo salvo em:\n{name}"),self._failed)


AUTOMATION_CLASS = CostAutomation
