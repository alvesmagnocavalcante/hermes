from __future__ import annotations

import asyncio
import base64
import math
import os
import unicodedata
from pathlib import Path
from decimal import Decimal
from typing import Any

import flet as ft

from hermes_ui.registry import (
    SPECS,
    AutomationSpec,
    format_value,
    record_status,
    record_value,
    result_details,
    searchable,
)


BLUE = "#2383C4"
SURFACE = "#171A1F"
CARD = "#20242B"
BORDER = "#303640"
TEXT = "#F4F7FB"
MUTED = "#AAB4C3"
GREEN = "#21A67A"
YELLOW = "#E0A83E"
RED = "#DC5A5A"
ROOT_DIR = Path(__file__).resolve().parent.parent
ASSETS_DIR = Path(os.environ.get("FLET_ASSETS_DIR", ROOT_DIR / "assets")).resolve()
LOGO_PATH = ASSETS_DIR / "deus.png"
LOGO_SOURCE = "data:image/png;base64," + base64.b64encode(LOGO_PATH.read_bytes()).decode("ascii")
WINDOW_ICON_PATH = ASSETS_DIR / "icon_windows.ico"


def normalized(value: str) -> str:
    text = unicodedata.normalize("NFKD", value)
    return "".join(char for char in text if not unicodedata.combining(char)).casefold()


def status_kind(status: str) -> str:
    value = normalized(status)
    if any(word in value for word in ("conciliad", "no prazo", "em dia", "com movimento", "pronto", "processado", "cobrado", "dados completos")):
        return "ok"
    if any(word in value for word in ("cancelad", "fora do periodo", "sem movimento", "informativ", "alerta")):
        return "info"
    return "error"


class AutomationView:
    PAGE_SIZE = 25

    def __init__(self, page: ft.Page, spec: AutomationSpec, set_status):
        self.page = page
        self.spec = spec
        self.set_status = set_status
        self.result: Any = None
        self.records: list[Any] = []
        self.current_page = 0
        self.search = ft.TextField(
            label="Buscar nos resultados",
            prefix_icon=ft.Icons.SEARCH,
            dense=True,
            expand=True,
            on_change=self._filter_changed,
        )
        self.status_filter = ft.Dropdown(
            value="Todos",
            label="Exibir",
            width=170,
            dense=True,
            options=(
                [ft.DropdownOption(key="Pendências", text="De/para incompleto"),
                 ft.DropdownOption(key="Conciliados", text="Prontos para importar"),
                 ft.DropdownOption(key="Todos", text="Todos")]
                if spec.key == "folha"
                else [ft.DropdownOption(key=value, text=value)
                      for value in (("Em atraso", "Alerta", "Em dia", "Todos")
                                    if spec.key == "entrada"
                                    else ("Pendências", "Conciliados", "Todos"))]
            ),
            on_select=self._filter_changed,
        )
        self.hotel = ft.Dropdown(
            value="Cumbuco",
            label="Hotel",
            width=150,
            dense=True,
            visible=spec.hotel_option,
            options=[ft.DropdownOption(key=value, text=value) for value in ("Cumbuco", "Magna", "Taiba", "Charme")],
        )
        self.output_format = ft.Dropdown(
            value=spec.formats[0],
            label="Formato",
            width=130,
            dense=True,
            options=[ft.DropdownOption(key=value, text=value) for value in spec.formats],
        )
        self.select_button = ft.FilledButton("Selecionar arquivos", icon=ft.Icons.UPLOAD_FILE, bgcolor=BLUE,
                                             color=ft.Colors.WHITE, on_click=self._select_files)
        self.export_button = ft.Button("Exportar resultado", icon=ft.Icons.DOWNLOAD, color="#80C8FA",
                                       disabled=True, on_click=self._export)
        self.file_summary = ft.Text("Arquivos necessários", weight=ft.FontWeight.W_600)
        self.file_hint = ft.Text(self._file_guidance(), color=MUTED, size=11)
        self.file_info = ft.Container(
            content=ft.Row([
                ft.Container(ft.Icon(ft.Icons.FOLDER_OPEN, color="#80C8FA", size=20), bgcolor="#173E59",
                             border_radius=8, padding=8),
                ft.Column([self.file_summary, self.file_hint], spacing=2),
            ], spacing=10),
            bgcolor="#141920", border=ft.Border.all(1, BORDER), border_radius=10, padding=10,
        )
        self.card_titles = [ft.Text(value, color=MUTED, size=12) for value in
                            ("Registros", "Conciliados", "Pendências", "Informativos")]
        self.card_values = [ft.Text("—", size=22, weight=ft.FontWeight.BOLD) for _ in range(4)]
        self.ring = ft.ProgressRing(value=0, stroke_width=9, color=GREEN, bgcolor="#39414D", width=68, height=68)
        self.ring_text = ft.Text("0%", size=15, weight=ft.FontWeight.BOLD)
        self.chart_title = ft.Text("Qualidade", weight=ft.FontWeight.BOLD, size=12)
        self.entry_segments = [
            ft.Container(width=0, height=14, bgcolor=color)
            for color in (GREEN, YELLOW, RED)
        ]
        self.entry_chart_legend = ft.Text("Aguardando análise", color=MUTED, size=10, text_align=ft.TextAlign.CENTER)
        self.entry_chart = ft.Column([
            ft.Stack([
                ft.Container(width=150, height=14, bgcolor="#39414D", border_radius=7),
                ft.Row(self.entry_segments, spacing=0, width=150),
            ], width=150, height=14),
            self.entry_chart_legend,
        ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=7)
        self.quality_chart = ft.Stack(
            [self.ring, ft.Container(self.ring_text, alignment=ft.Alignment.CENTER)], width=68, height=68,
        )
        self.table_host = ft.Container(expand=True)
        self.page_text = ft.Text("Página 0 de 0", color=MUTED)
        self.previous = ft.IconButton(ft.Icons.CHEVRON_LEFT, tooltip="Página anterior", disabled=True, on_click=self._previous)
        self.next = ft.IconButton(ft.Icons.CHEVRON_RIGHT, tooltip="Próxima página", disabled=True, on_click=self._next)
        self.details = ft.Column(spacing=3)
        self.control = self._build()
        self._render_table()

    def _build(self) -> ft.Control:
        metrics = ft.Row(
            [self._metric_card(self.card_titles[index], self.card_values[index], color) for index, color in
             enumerate((BLUE, GREEN, RED, YELLOW))] + [
                ft.Container(
                    content=ft.Column([
                        self.chart_title,
                        self.entry_chart if self.spec.key == "entrada" else self.quality_chart,
                    ], horizontal_alignment=ft.CrossAxisAlignment.CENTER, alignment=ft.MainAxisAlignment.CENTER, spacing=5),
                    bgcolor=CARD, border=ft.Border.all(1, BORDER), border_radius=14,
                    padding=10, width=190, height=112,
                )
            ],
            spacing=12,
        )
        controls = ft.Row(
            [self.select_button, self.hotel, self.output_format, self.export_button,
             ft.Button("Limpar", icon=ft.Icons.DELETE_OUTLINE, on_click=self._clear)],
            spacing=10,
            wrap=True,
        )
        filters = ft.Row([self.status_filter, self.search], spacing=12)
        legend = ft.Row(
            [self._legend(color, label) for color, label in self._legend_items()],
            spacing=22,
            scroll=ft.ScrollMode.AUTO,
        )
        pagination = ft.Row([self.previous, self.page_text, self.next], alignment=ft.MainAxisAlignment.CENTER)
        return ft.Container(
            bgcolor=SURFACE,
            border_radius=18,
            padding=24,
            expand=True,
            content=ft.Column([
                ft.Text(self.spec.name, size=28, weight=ft.FontWeight.BOLD),
                ft.Text(self.spec.description, color=MUTED),
                controls,
                self.file_info,
                self.details,
                metrics,
                filters,
                legend,
                self.table_host,
                pagination,
            ], spacing=13, expand=True),
        )

    @staticmethod
    def _metric_card(name: ft.Text, value: ft.Text, color: str) -> ft.Container:
        return ft.Container(
            content=ft.Column([name, value], spacing=5),
            bgcolor=CARD,
            border=ft.Border(left=ft.BorderSide(4, color), top=ft.BorderSide(1, BORDER),
                             right=ft.BorderSide(1, BORDER), bottom=ft.BorderSide(1, BORDER)),
            border_radius=14,
            padding=16,
            height=112,
            expand=1,
        )

    @staticmethod
    def _legend(color: str, label: str) -> ft.Control:
        return ft.Row([ft.Container(width=9, height=9, bgcolor=color, shape=ft.BoxShape.CIRCLE), ft.Text(label, color=MUTED, size=12)], spacing=6)

    def _legend_items(self) -> tuple[tuple[str, str], ...]:
        return {
            "folha": ((GREEN, "Pronto para importar"), (RED, "De/para incompleto")),
            "receita": ((GREEN, "Conciliado"), (RED, "Valor divergente")),
            "diarias": ((GREEN, "Com movimento"), (YELLOW, "Sem movimento")),
            "cupons_hospede": ((GREEN, "Cupom cobrado"), (RED, "Ausente ou não cobrado")),
            "rps": ((GREEN, "Integrado nas três fontes"), (YELLOW, "Cancelado ou fora do período"),
                    (RED, "Ausente, irregular ou divergente")),
            "debito": ((GREEN, "Nota processada"),),
            "entrada": ((GREEN, "Em dia"), (YELLOW, "Em alerta"), (RED, "Em atraso")),
            "cupons": ((GREEN, "Conciliado nas três fontes"), (RED, "Ausência, data ou valor divergente")),
            "servicos": ((GREEN, "Nota conciliada"), (YELLOW, "Informação complementar"),
                         (RED, "Pendência no CAP, hotel ou ISS")),
            "receber": ((GREEN, "Conciliado"), (RED, "Divergente")),
            "pagar": ((GREEN, "Conciliado"), (RED, "Divergente")),
            "custos": ((GREEN, "Conciliado"), (RED, "Divergente")),
        }[self.spec.key]

    def _file_guidance(self) -> str:
        return {
            "receita": "Razão Analítico da Contabilidade + Journal de Receita do Opera.",
            "diarias": "Códigos de transação + Journal Opera - Receita. BI PDV não pertence a esta conferência.",
            "folha": "Sete relatórios do DP; a planilha modelo pode ser incluída para atualizar o de/para.",
            "cupons_hospede": "BI PDV + Journal do Opera + planilha de de/para.",
            "rps": "XML de encerramentos do Opera + Fiscal do CMFlex + Prefeitura.",
            "debito": "Uma ou mais planilhas de notas de débito.",
            "entrada": "Manifesto de notas (base completa) + Detalhe de notas recebidas (notas já lançadas).",
            "cupons": "Simphony + Fiscal do CMFlex + SEFAZ.",
            "servicos": "CAP + Portal Nacional + arquivos de Prefeitura/ISS aplicáveis ao hotel.",
            "receber": "Seis relatórios: balancete, posição por cliente, borderô, razão a faturar, agregados e razão.",
            "pagar": "Oito relatórios de fornecedores, adiantamentos e impostos.",
            "custos": "Documentos por tipo de desembolso + Razão + Inventário + Razão de estoques.",
        }[self.spec.key]

    async def _select_files(self, _event=None) -> None:
        picker = ft.FilePicker()
        files = await picker.pick_files(
            dialog_title=f"Arquivos — {self.spec.name}",
            allow_multiple=True,
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=list(self.spec.extensions),
        )
        if not files:
            return
        paths = [Path(file.path) for file in files if file.path]
        self.select_button.disabled = True
        self.export_button.disabled = True
        self.set_status("Processando arquivos...", None)
        self.page.update()
        try:
            self.result = await asyncio.to_thread(self.spec.analyze, paths, self.hotel.value)
            self.records = self.spec.rows(self.result)
        except Exception as error:
            self._show_error(error)
            self.set_status("Falha no processamento", 0)
        else:
            self.current_page = 0
            self.export_button.disabled = False
            names = "\n".join(path.name for path in paths)
            self.file_summary.value = f"{len(paths)} arquivo(s) reconhecido(s) com sucesso"
            self.file_summary.tooltip = names
            self._show_result_details()
            self.set_status("Conferência concluída", 1)
            self._update_metrics()
            self._render_table()
        finally:
            self.select_button.disabled = False
            self.page.update()

    async def _export(self, _event=None) -> None:
        if self.result is None:
            return
        extensions = {"Excel": "xlsx", "PDF": "pdf", "CSV": "csv"}
        extension = extensions[self.output_format.value]
        picker = ft.FilePicker()
        output = await picker.save_file(
            dialog_title="Exportar resultado",
            file_name=f"{self.spec.key}_resultado.{extension}",
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=[extension],
        )
        if not output:
            return
        path = Path(output)
        if path.suffix.lower() != f".{extension}":
            path = path.with_suffix(f".{extension}")
        self.export_button.disabled = True
        self.set_status("Exportando resultado...", None)
        self.page.update()
        try:
            await asyncio.to_thread(self.spec.export, self.result, path, self.output_format.value)
        except Exception as error:
            self._show_error(error)
            self.set_status("Falha na exportação", 0)
        else:
            self.set_status(f"Resultado exportado: {path.name}", 1)
            self.page.show_dialog(
                ft.SnackBar("Resultado exportado com sucesso.", bgcolor=GREEN, show_close_icon=True)
            )
        finally:
            self.export_button.disabled = False
            self.page.update()

    def _filtered(self) -> list[Any]:
        query = self.search.value.strip().casefold()
        selected = self.status_filter.value
        records = []
        for record in self.records:
            status = record_status(record)
            kind = status_kind(status)
            if self.spec.key == "entrada":
                if selected != "Todos" and status != selected:
                    continue
            else:
                if selected == "Conciliados" and kind != "ok":
                    continue
                if selected == "Pendências" and kind == "ok":
                    continue
            if query and query not in searchable(record, self.spec.columns):
                continue
            records.append(record)
        return records

    def _show_result_details(self) -> None:
        result = self.result
        if self.spec.key == "folha":
            self.details.controls = [ft.Text(
                f"{result.company}  •  Competência {result.period_end:%m/%Y}  •  "
                f"Proventos {format_value(result.earnings)}  •  Descontos {format_value(result.deductions)}  •  "
                f"Líquido a pagar {format_value(result.net_payable)}  •  "
                f"Férias: {result.vacation_employees} funcionário(s), "
                f"{result.vacation_entries} lançamento(s) individualizado(s)  •  "
                f"Excluídos: {result.ignored_rows} totalizadores e "
                f"{result.excluded_rows} eventos duplicados do resumo mensal",
                color=MUTED, size=12,
            )]
            return
        if self.spec.key == "entrada":
            not_posted = sum(row.entry_date is None for row in result.rows)
            self.details.controls = [ft.Text(
                f"{not_posted} nota(s) ainda não lançada(s)  •  "
                "CE: alerta de 6 a 10 dias e atraso a partir de 11 dias  •  "
                "Demais estados: alerta de 20 a 29 dias e atraso a partir de 30 dias",
                color=MUTED, size=12,
            )]
            return
        if self.spec.key == "receber":
            client_difference = result.client_accounting_total - result.client_financial_total
            self.details.controls = [
                self._comparison_line(
                    "Clientes", "Balancete", result.client_accounting_total,
                    "Posição por cliente", result.client_financial_total, client_difference,
                    "Conciliado" if abs(client_difference) <= Decimal("0.01") else "Divergente",
                ),
                self._comparison_line(
                    "Notas a faturar", "Borderô", result.billing.source_value,
                    "Razão a faturar", result.billing.accounting_value,
                    result.billing.difference, result.billing.status,
                ),
                self._comparison_line(
                    "Comissões", "Agregados lançados", result.commissions.source_value,
                    "Razão", result.commissions.accounting_value,
                    result.commissions.difference, result.commissions.status,
                ),
            ]
            return
        if self.spec.key == "receita":
            difference = result.cmflex_total - result.opera_total
            self.details.controls = [self._comparison_line(
                "Receita", "Contabilidade", result.cmflex_total, "Opera", result.opera_total,
                difference, self._difference_status(difference),
            )]
            return
        if self.spec.key == "cupons_hospede":
            journal_total = sum((row.journal_value or Decimal() for row in result.coupons), Decimal())
            difference = result.total_pdv - journal_total
            self.details.controls = [self._comparison_line(
                "Cupons cobrados", "BI/PDV", result.total_pdv, "Conta do hóspede", journal_total,
                difference, self._difference_status(difference),
            )]
            return
        if self.spec.key == "rps":
            opera = sum((row.opera_value or Decimal() for row in result.rows), Decimal())
            fiscal = sum((row.fiscal_value or Decimal() for row in result.rows), Decimal())
            city = sum((row.city_value or Decimal() for row in result.rows), Decimal())
            opera_fiscal = opera - fiscal
            fiscal_city = fiscal - city
            self.details.controls = [
                self._comparison_line("Integração Fiscal", "Opera", opera, "Fiscal", fiscal,
                                      opera_fiscal, self._difference_status(opera_fiscal)),
                self._comparison_line("Emissão de NFS-e", "Fiscal", fiscal, "Prefeitura", city,
                                      fiscal_city, self._difference_status(fiscal_city)),
            ]
            return
        if self.spec.key == "cupons":
            simphony_fiscal = result.simphony_total - result.fiscal_total
            fiscal_sefaz = result.fiscal_total - result.sefaz_total
            self.details.controls = [
                self._comparison_line("Integração Fiscal", "Simphony", result.simphony_total,
                                      "Fiscal", result.fiscal_total, simphony_fiscal,
                                      self._difference_status(simphony_fiscal)),
                self._comparison_line("Integração SEFAZ", "Fiscal", result.fiscal_total,
                                      "SEFAZ", result.sefaz_total, fiscal_sefaz,
                                      self._difference_status(fiscal_sefaz)),
            ]
            return
        if self.spec.key == "servicos":
            external_notes: dict[tuple[str, str], Decimal] = {}
            cap_notes: dict[tuple[str, str], Decimal] = {}
            city_taxes: dict[tuple[str, str], Decimal] = {}
            cap_taxes: dict[tuple[str, str], Decimal] = {}
            for row in result.rows:
                key = (row.cnpj or row.provider, row.number)
                if row.gross is not None:
                    external_notes.setdefault(key, row.gross)
                if row.cap_gross is not None:
                    cap_notes.setdefault(key, row.cap_gross)
                if row.iss is not None:
                    city_taxes.setdefault(key, row.iss)
                if row.cap_iss is not None:
                    cap_taxes.setdefault(key, row.cap_iss)
            external_gross = sum(external_notes.values(), Decimal())
            cap_gross = sum(cap_notes.values(), Decimal())
            city_iss = sum(city_taxes.values(), Decimal())
            cap_iss = sum(cap_taxes.values(), Decimal())
            gross_difference = external_gross - cap_gross
            iss_difference = city_iss - cap_iss
            self.details.controls = [
                self._comparison_line("Notas de serviço", "Fontes externas", external_gross,
                                      "CAP", cap_gross, gross_difference,
                                      self._difference_status(gross_difference)),
                self._comparison_line("ISS retido", "Prefeitura", city_iss, "CAP", cap_iss,
                                      iss_difference, self._difference_status(iss_difference)),
            ]
            return
        if self.spec.key == "pagar":
            checks = (result.suppliers, result.advances, *result.taxes)
            self.details.controls = [
                self._comparison_line(check.name, "Financeiro", check.financial,
                                      "Contabilidade", check.accounting, check.difference, check.status)
                for check in checks
            ]
            return
        if self.spec.key == "custos":
            lines = []
            for analysis, source_name in (("Entradas", "CAP"), ("Saldo final", "Inventário")):
                rows = [row for row in result if row.analysis == analysis]
                source = sum((row.source for row in rows), Decimal())
                accounting = sum((row.accounting for row in rows), Decimal())
                difference = source - accounting
                lines.append(self._comparison_line(
                    analysis, source_name, source, "Contabilidade", accounting,
                    difference, self._difference_status(difference),
                ))
            self.details.controls = lines
            return
        details = result_details(result)
        self.details.controls = [ft.Text(
            "  •  ".join(f"{key}: {value}" for key, value in details.items()),
            color=MUTED, size=12, max_lines=2, overflow=ft.TextOverflow.ELLIPSIS,
        )] if details else []

    @staticmethod
    def _difference_status(difference: Decimal) -> str:
        return "Conciliado" if abs(difference) <= Decimal("0.01") else "Divergente"

    @staticmethod
    def _comparison_line(title: str, first_label: str, first_value, second_label: str,
                         second_value, difference, status: str) -> ft.Text:
        color = GREEN if status_kind(status) == "ok" else RED
        return ft.Text(
            spans=[
                ft.TextSpan(f"{title}: ", style=ft.TextStyle(weight=ft.FontWeight.BOLD, color=TEXT)),
                ft.TextSpan(f"{first_label} {format_value(first_value)}  •  "),
                ft.TextSpan(f"{second_label} {format_value(second_value)}  •  "),
                ft.TextSpan(f"Diferença {format_value(difference)}  •  "),
                ft.TextSpan(status, style=ft.TextStyle(weight=ft.FontWeight.BOLD, color=color)),
            ],
            color=MUTED,
            size=12,
            no_wrap=True,
            overflow=ft.TextOverflow.ELLIPSIS,
        )

    def _render_table(self) -> None:
        records = self._filtered()
        pages = max(1, math.ceil(len(records) / self.PAGE_SIZE))
        self.current_page = min(self.current_page, pages - 1)
        start = self.current_page * self.PAGE_SIZE
        visible_records = records[start:start + self.PAGE_SIZE]

        header = ft.Row(
            [self._table_cell(column, column.label, TEXT, header=True) for column in self.spec.columns],
            spacing=0,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        rows: list[ft.Control] = []
        for index, record in enumerate(visible_records):
            status = record_status(record)
            color = {"ok": GREEN, "info": YELLOW, "error": RED}[status_kind(status)]
            cells = [
                self._table_cell(
                    column,
                    format_value(status if column.key == "status" else record_value(record, column.key)),
                    color,
                )
                for column in self.spec.columns
            ]
            rows.append(ft.Container(
                content=ft.Row(cells, spacing=0, vertical_alignment=ft.CrossAxisAlignment.CENTER),
                height=42,
                bgcolor="#15191E" if index % 2 == 0 else "#121519",
                border=ft.Border(bottom=ft.BorderSide(1, BORDER)),
            ))

        table = ft.Container(
            content=ft.Column([
                ft.Container(header, height=48, bgcolor=CARD,
                             border=ft.Border(bottom=ft.BorderSide(1, "#46505D"))),
                *rows,
            ], spacing=0),
            border=ft.Border.all(1, BORDER),
            border_radius=10,
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
            expand=True,
        )
        self.table_host.content = ft.Column(
            [table],
            scroll=ft.ScrollMode.AUTO,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            expand=True,
        )
        self.page_text.value = f"Página {self.current_page + 1} de {pages} • {len(records)} registro(s)"
        self.previous.disabled = self.current_page == 0
        self.next.disabled = self.current_page + 1 >= pages

    def _table_cell(self, column, value: str, color: str, *, header: bool = False) -> ft.Container:
        wide = {"description", "detail", "customer", "provider", "supplier", "organogram", "name", "guest"}
        medium = {"document", "key", "company", "source", "analysis", "status", "item", "comprador"}
        weight = 3 if column.key in wide else 2 if column.key in medium else 1
        alignment = ft.Alignment.CENTER_RIGHT if column.numeric else ft.Alignment.CENTER_LEFT
        return ft.Container(
            content=ft.Text(
                value,
                color=color,
                size=12 if len(self.spec.columns) <= 7 else 11,
                weight=ft.FontWeight.BOLD if header else ft.FontWeight.NORMAL,
                no_wrap=True,
                overflow=ft.TextOverflow.ELLIPSIS,
                tooltip=None if header else value,
                text_align=ft.TextAlign.RIGHT if column.numeric else ft.TextAlign.LEFT,
            ),
            padding=ft.Padding(left=14, top=0, right=14, bottom=0),
            alignment=alignment,
            expand=weight,
        )

    def _update_metrics(self) -> None:
        kinds = [status_kind(record_status(record)) for record in self.records]
        total, ok, info = len(kinds), kinds.count("ok"), kinds.count("info")
        error = total - ok - info
        labels = ("Registros", "Conciliados", "Pendências", "Informativos")
        values = (total, ok, error, info)
        if self.spec.key == "folha":
            labels = ("Lançamentos", "Prontos para importar", "De/para incompleto", "Exclusões aplicadas")
            values = (
                total, self.result.ready, total - self.result.ready,
                self.result.ignored_rows + self.result.excluded_rows,
            )
            ok, error, info = self.result.ready, total - self.result.ready, 0
            self.chart_title.value = "Qualidade dos lançamentos"
        elif self.spec.key == "entrada":
            labels = ("Notas", "Em dia", "Em atraso", "Em alerta")
            values = (total, ok, error, info)
            self.chart_title.value = "Distribuição das notas"
            for segment, amount in zip(self.entry_segments, (ok, info, error)):
                segment.width = 150 * amount / total if total else 0
            self.entry_chart_legend.value = f"Em dia {ok}  •  Alerta {info}  •  Atraso {error}"
        else:
            self.chart_title.value = "Qualidade da conferência"
        for title, value_control, label, amount in zip(self.card_titles, self.card_values, labels, values):
            title.value = label
            value_control.value = str(amount)
        ratio = ok / total if total else 0
        self.ring.value = ratio
        self.ring_text.value = f"{ratio:.1%}"

    def _filter_changed(self, _event=None) -> None:
        self.current_page = 0
        self._render_table()
        self.page.update()

    def _previous(self, _event=None) -> None:
        if self.current_page:
            self.current_page -= 1
            self._render_table()
            self.page.update()

    def _next(self, _event=None) -> None:
        self.current_page += 1
        self._render_table()
        self.page.update()

    def _clear(self, _event=None) -> None:
        self.result = None
        self.records = []
        self.current_page = 0
        self.search.value = ""
        self.status_filter.value = "Todos"
        self.export_button.disabled = True
        self.file_summary.value = "Arquivos necessários"
        self.file_hint.value = self._file_guidance()
        self.file_summary.tooltip = None
        self.details.controls = []
        for title, label in zip(self.card_titles, ("Registros", "Conciliados", "Pendências", "Informativos")):
            title.value = label
        for value in self.card_values:
            value.value = "—"
        self.ring.value = 0
        self.ring_text.value = "0%"
        if self.spec.key == "entrada":
            self.chart_title.value = "Distribuição das notas"
            for segment in self.entry_segments:
                segment.width = 0
            self.entry_chart_legend.value = "Aguardando análise"
        self._render_table()
        self.set_status("Seleção limpa", 0)
        self.page.update()

    def _show_error(self, error: Exception) -> None:
        dialog = ft.AlertDialog(
            modal=True,
            title="Não foi possível concluir",
            content=ft.Text(str(error)),
            actions=[ft.Button("Fechar", on_click=lambda _: self.page.pop_dialog())],
        )
        self.page.show_dialog(dialog)


class HermesApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.current: AutomationView | None = None
        self.selected_index = 0
        self.sidebar_expanded = False
        self.content = ft.Container(expand=True)
        self.status_text = ft.Text("Pronto", color=MUTED, size=12)
        self.progress = ft.ProgressBar(value=0, color=BLUE, bgcolor="#343A45", bar_height=4)
        self.navigation_icons = (
            ft.Icons.COMPARE_ARROWS, ft.Icons.HOTEL, ft.Icons.PAYMENTS, ft.Icons.RECEIPT_LONG,
            ft.Icons.CLOUD_SYNC, ft.Icons.DESCRIPTION, ft.Icons.SCHEDULE, ft.Icons.POINT_OF_SALE,
            ft.Icons.REQUEST_QUOTE, ft.Icons.ACCOUNT_BALANCE_WALLET, ft.Icons.ACCOUNT_BALANCE,
            ft.Icons.INVENTORY_2,
        )
        self.sidebar = ft.Container(
            width=126,
            bgcolor="#101318",
            border=ft.Border(right=ft.BorderSide(1, BORDER)),
            animate_size=180,
        )
        self._render_sidebar()

    def build(self) -> None:
        self.page.title = "HERMES — Painel de Automação"
        self.page.window.icon = str(WINDOW_ICON_PATH)
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.bgcolor = "#0D1014"
        self.page.padding = 0
        self.page.window.width = 1500
        self.page.window.height = 900
        self.page.window.min_width = 1000
        self.page.window.min_height = 680
        self.page.theme = ft.Theme(color_scheme_seed=BLUE, font_family="Segoe UI")
        footer = ft.Container(
            content=ft.Column([self.status_text, self.progress], spacing=6),
            bgcolor="#111419", padding=ft.Padding(left=18, top=9, right=18, bottom=11),
        )
        body = ft.Row([
            self.sidebar,
            ft.Container(self.content, padding=ft.Padding(left=14, top=14, right=14, bottom=10), expand=True),
        ], spacing=0, expand=True)
        self.page.add(ft.Column([body, footer], spacing=0, expand=True))
        self._show(0)

    def _show(self, index: int) -> None:
        self.selected_index = index
        self.current = AutomationView(self.page, SPECS[index], self.set_status)
        self.content.content = self.current.control
        self._render_sidebar()
        self.set_status(f"{SPECS[index].name} carregada", 0)
        self.page.update()

    def _toggle_rail(self, _event=None) -> None:
        self.sidebar_expanded = not self.sidebar_expanded
        self.sidebar.width = 270 if self.sidebar_expanded else 126
        self._render_sidebar()
        self.page.update()

    def _show_about(self, _event=None) -> None:
        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row([
                ft.Image(src=LOGO_SOURCE, width=34, height=34, fit=ft.BoxFit.CONTAIN),
                ft.Text("Sobre o HERMES", weight=ft.FontWeight.BOLD),
            ], spacing=10, tight=True),
            content=ft.Column([
                ft.Text("Painel de Automação de Planilhas", color=MUTED),
                ft.Text("Versão 1.0.0", size=12, color=MUTED),
                ft.Divider(color=BORDER),
                ft.Text("Desenvolvido por Magno Alves", weight=ft.FontWeight.W_600),
            ], spacing=8, tight=True, width=330),
            actions=[ft.Button("Fechar", on_click=lambda _: self.page.pop_dialog())],
        )
        self.page.show_dialog(dialog)

    def _render_sidebar(self) -> None:
        logo_size = 38 if self.sidebar_expanded else 28
        brand = ft.Row([
            ft.Image(
                src=LOGO_SOURCE,
                width=logo_size,
                height=logo_size,
                fit=ft.BoxFit.CONTAIN,
                tooltip="HERMES",
            ),
            ft.Text("HERMES", size=18, weight=ft.FontWeight.BOLD, color=TEXT),
        ], spacing=7 if self.sidebar_expanded else 5, tight=True)
        toggle = ft.IconButton(
            ft.Icons.CHEVRON_LEFT if self.sidebar_expanded else ft.Icons.MENU,
            tooltip="Recolher menu" if self.sidebar_expanded else "Abrir menu",
            on_click=self._toggle_rail,
        )
        header = (ft.Row([brand, toggle], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
                  if self.sidebar_expanded else
                  ft.Column([brand, toggle], horizontal_alignment=ft.CrossAxisAlignment.CENTER, spacing=4))
        buttons = []
        for index, (spec, icon) in enumerate(zip(SPECS, self.navigation_icons)):
            selected = index == self.selected_index
            icon_control = ft.Icon(icon, size=20, color="#9BD5FA" if selected else MUTED)
            content = (ft.Row([
                icon_control,
                ft.Text(spec.name, size=12, weight=ft.FontWeight.W_600, color=TEXT),
            ], spacing=12, alignment=ft.MainAxisAlignment.START)
                       if self.sidebar_expanded else icon_control)
            buttons.append(ft.Container(
                content=content,
                height=43,
                padding=ft.Padding(left=18 if self.sidebar_expanded else 0, top=0,
                                   right=12 if self.sidebar_expanded else 0, bottom=0),
                alignment=ft.Alignment.CENTER_LEFT if self.sidebar_expanded else ft.Alignment.CENTER,
                bgcolor="#174B6D" if selected else ft.Colors.TRANSPARENT,
                border_radius=10,
                tooltip=None if self.sidebar_expanded else spec.name,
                on_click=lambda _event, current=index: self._show(current),
            ))
        about_icon = ft.Icon(ft.Icons.INFO_OUTLINE, size=19, color=MUTED)
        about_content = (ft.Row([
            about_icon,
            ft.Text("Sobre", size=12, color=MUTED),
        ], spacing=12, tight=True) if self.sidebar_expanded else about_icon)
        about = ft.Container(
            content=about_content,
            height=38,
            padding=ft.Padding(left=18 if self.sidebar_expanded else 0, top=0,
                               right=12 if self.sidebar_expanded else 0, bottom=0),
            alignment=ft.Alignment.CENTER_LEFT if self.sidebar_expanded else ft.Alignment.CENTER,
            border_radius=9,
            tooltip=None if self.sidebar_expanded else "Sobre o HERMES",
            on_click=self._show_about,
        )
        self.sidebar.content = ft.Column([
            ft.Container(header, padding=ft.Padding(left=16, top=14, right=10, bottom=8)),
            ft.Divider(height=1, color=BORDER),
            ft.Column(buttons, spacing=5, scroll=ft.ScrollMode.AUTO, expand=True),
            ft.Divider(height=1, color=BORDER),
            ft.Container(about, padding=ft.Padding(left=8, top=0, right=8, bottom=10)),
        ], spacing=10, expand=True)

    def set_status(self, message: str, progress: float | None) -> None:
        self.status_text.value = message
        self.progress.value = progress


def main(page: ft.Page) -> None:
    HermesApp(page).build()
