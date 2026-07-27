from __future__ import annotations

import asyncio
import base64
import logging
import math
import os
import tempfile
import time
import unicodedata
from decimal import Decimal
from pathlib import Path
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
from hermes_ui.runtime import JOB_LIMITER, validate_upload

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
LOGO_SOURCE = "data:image/png;base64," + base64.b64encode(
    LOGO_PATH.read_bytes()
).decode("ascii")
WINDOW_ICON_PATH = ASSETS_DIR / "icon_windows.ico"
LOGGER = logging.getLogger("hermes")


def normalized(value: str) -> str:
    text = unicodedata.normalize("NFKD", value)
    return "".join(char for char in text if not unicodedata.combining(char)).casefold()


def status_kind(status: str) -> str:
    value = normalized(status)
    if any(
        word in value
        for word in (
            "conciliad",
            "no prazo",
            "em dia",
            "com movimento",
            "pronto",
            "processado",
            "cobrado",
            "dados completos",
        )
    ):
        return "ok"
    if any(
        word in value
        for word in (
            "cancelad",
            "fora do periodo",
            "sem movimento",
            "informativ",
            "alerta",
        )
    ):
        return "info"
    return "error"


class AutomationView:
    MIN_PAGE_SIZE = 10

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
                [
                    ft.DropdownOption(key="Pendências", text="De/para incompleto"),
                    ft.DropdownOption(key="Conciliados", text="Prontos para importar"),
                    ft.DropdownOption(key="Todos", text="Todos"),
                ]
                if spec.key == "folha"
                else [
                    ft.DropdownOption(key=value, text=value)
                    for value in (
                        ("Em atraso", "Alerta", "Em dia", "Todos")
                        if spec.key == "entrada"
                        else ("Pendências", "Conciliados", "Todos")
                    )
                ]
            ),
            on_select=self._filter_changed,
        )
        self.hotel = ft.Dropdown(
            value="Cumbuco",
            label="Hotel",
            width=150,
            dense=True,
            visible=spec.hotel_option,
            options=[
                ft.DropdownOption(key=value, text=value)
                for value in ("Cumbuco", "Magna", "Taiba", "Charme", "Wind")
            ],
        )
        self.output_format = ft.Dropdown(
            value=spec.formats[0],
            label="Formato",
            width=130,
            dense=True,
            options=[
                ft.DropdownOption(key=value, text=value) for value in spec.formats
            ],
        )
        self.select_button = ft.FilledButton(
            "Arquivos",
            icon=ft.Icons.UPLOAD_FILE,
            bgcolor=BLUE,
            color=ft.Colors.WHITE,
            on_click=self._select_files,
        )
        self.export_button = ft.Button(
            "Exportar",
            icon=ft.Icons.DOWNLOAD,
            color="#80C8FA",
            disabled=True,
            on_click=self._export,
        )
        self.details_button = ft.Button(
            "Resumo",
            icon=ft.Icons.SUMMARIZE_OUTLINED,
            color="#B7C7D9",
            visible=False,
            on_click=self._show_summary,
        )
        self.file_summary = ft.Text("Arquivos necessários", weight=ft.FontWeight.W_600)
        self.file_hint = ft.Text(self._file_guidance(), color=MUTED, size=11)
        self.file_info = ft.Container(
            content=ft.Row(
                [
                    ft.Container(
                        ft.Icon(ft.Icons.FOLDER_OPEN, color="#80C8FA", size=18),
                        bgcolor="#173E59",
                        border_radius=8,
                        padding=6,
                    ),
                    ft.Column([self.file_summary, self.file_hint], spacing=2),
                ],
                spacing=10,
            ),
            bgcolor="#141920",
            border=ft.Border.all(1, BORDER),
            border_radius=10,
            padding=6,
        )
        self.card_titles = [
            ft.Text(value, color=MUTED, size=12)
            for value in ("Registros", "Conciliados", "Pendências", "Informativos")
        ]
        self.card_values = [
            ft.Text("—", size=18, weight=ft.FontWeight.BOLD) for _ in range(4)
        ]
        self.quality_title = ft.Text("Qualidade", color=MUTED, size=12)
        self.quality_value = ft.Text("0%", size=18, weight=ft.FontWeight.BOLD)
        self.table_host = ft.Container(expand=True)
        self.page_text = ft.Text("Página 0 de 0", color=MUTED)
        self.previous = ft.IconButton(
            ft.Icons.CHEVRON_LEFT,
            tooltip="Página anterior",
            disabled=True,
            on_click=self._previous,
        )
        self.next = ft.IconButton(
            ft.Icons.CHEVRON_RIGHT,
            tooltip="Próxima página",
            disabled=True,
            on_click=self._next,
        )
        self.details = ft.Column(spacing=3)
        self.control = self._build()
        self._render_table()

    def _build(self) -> ft.Control:
        self.metrics = ft.Row(
            [
                self._metric_card(
                    self.card_titles[index], self.card_values[index], color
                )
                for index, color in enumerate((BLUE, GREEN, RED, YELLOW))
            ]
            + [self._metric_card(self.quality_title, self.quality_value, GREEN)],
            spacing=8,
            visible=False,
        )
        actions = ft.Row(
            [
                self.select_button,
                self.hotel,
                self.output_format,
                self.export_button,
                self.details_button,
                ft.Button("Limpar", icon=ft.Icons.DELETE_OUTLINE, on_click=self._clear),
            ],
            spacing=6,
        )
        header = ft.Row(
            [
                ft.Column(
                    [
                        ft.Text(self.spec.name, size=24, weight=ft.FontWeight.BOLD),
                        ft.Text(self.spec.description, color=MUTED, size=12),
                    ],
                    spacing=3,
                    expand=True,
                ),
                actions,
            ],
            spacing=12,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        filters = ft.Row(
            [self.status_filter, self.search, self.previous, self.page_text, self.next],
            spacing=6,
            vertical_alignment=ft.CrossAxisAlignment.CENTER,
        )
        return ft.Container(
            bgcolor=SURFACE,
            border_radius=18,
            padding=14,
            expand=True,
            content=ft.Column(
                [
                    header,
                    self.file_info,
                    self.metrics,
                    filters,
                    self.table_host,
                ],
                spacing=8,
                expand=True,
            ),
        )

    @staticmethod
    def _metric_card(name: ft.Text, value: ft.Text, color: str) -> ft.Container:
        return ft.Container(
            content=ft.Column([name, value], spacing=5),
            bgcolor=CARD,
            border=ft.Border(
                left=ft.BorderSide(4, color),
                top=ft.BorderSide(1, BORDER),
                right=ft.BorderSide(1, BORDER),
                bottom=ft.BorderSide(1, BORDER),
            ),
            border_radius=14,
            padding=8,
            height=58,
            expand=1,
        )

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
            # Browsers do not expose the user's local file path. In web mode,
            # Flet must return the contents so they can be processed server-side.
            with_data=self.page.web,
            cancel_upload_on_window_blur=False,
        )
        if not files:
            return
        self.select_button.disabled = True
        self.export_button.disabled = True
        started_at = time.monotonic()
        self.page.update()
        try:
            total_size = validate_upload(files, self.page.web)
            self.set_status("Aguardando capacidade de processamento...", None)
            self.page.update()
            async with JOB_LIMITER:
                self.set_status("Processando arquivos...", None)
                self.page.update()
                if self.page.web:
                    with tempfile.TemporaryDirectory(
                        prefix="hermes_upload_"
                    ) as temp_dir:
                        paths = self._save_web_files(files, Path(temp_dir))
                        self.result = await asyncio.to_thread(
                            self.spec.analyze, paths, self.hotel.value
                        )
                else:
                    paths = [Path(file.path) for file in files if file.path]
                    self.result = await asyncio.to_thread(
                        self.spec.analyze, paths, self.hotel.value
                    )
            self.records = self.spec.rows(self.result)
        except Exception as error:
            LOGGER.exception("Falha na automação %s", self.spec.key)
            self._show_error(error)
            self.set_status("Falha no processamento", 0)
        else:
            LOGGER.info(
                "Automação concluída: key=%s files=%d bytes=%d records=%d duration=%.3fs",
                self.spec.key,
                len(files),
                total_size,
                len(self.records),
                time.monotonic() - started_at,
            )
            self.current_page = 0
            self.export_button.disabled = False
            names = "\n".join(path.name for path in paths)
            self.file_summary.value = (
                f"{len(paths)} arquivo(s) reconhecido(s) com sucesso"
            )
            self.file_summary.tooltip = names
            self._show_result_details()
            self.file_info.visible = False
            self.metrics.visible = True
            self.details_button.visible = True
            self.set_status("Conferência concluída", 1)
            self._update_metrics()
            self._render_table()
        finally:
            self.select_button.disabled = False
            self.page.update()

    @staticmethod
    def _save_web_files(files: list[ft.FilePickerFile], directory: Path) -> list[Path]:
        paths: list[Path] = []
        used_names: set[str] = set()
        for index, selected in enumerate(files, start=1):
            if selected.bytes is None:
                raise ValueError(
                    f"O navegador não enviou o conteúdo de {selected.name}."
                )

            original = Path(selected.name).name or f"arquivo_{index}"
            candidate = original
            suffix = Path(original).suffix
            stem = Path(original).stem
            duplicate = 2
            while candidate.casefold() in used_names:
                candidate = f"{stem}_{duplicate}{suffix}"
                duplicate += 1

            used_names.add(candidate.casefold())
            path = directory / candidate
            path.write_bytes(selected.bytes)
            paths.append(path)
        return paths

    async def _export(self, _event=None) -> None:
        if self.result is None:
            return
        extensions = {"Excel": "xlsx", "PDF": "pdf", "CSV": "csv"}
        extension = extensions[self.output_format.value]
        file_name = f"{self.spec.key}_resultado.{extension}"
        picker = ft.FilePicker()
        self.export_button.disabled = True
        started_at = time.monotonic()
        self.set_status("Aguardando capacidade para exportar...", None)
        self.page.update()
        try:
            async with JOB_LIMITER:
                self.set_status("Exportando resultado...", None)
                self.page.update()
                if self.page.web:
                    with tempfile.TemporaryDirectory(
                        prefix="hermes_export_"
                    ) as temp_dir:
                        path = Path(temp_dir) / file_name
                        await asyncio.to_thread(
                            self.spec.export,
                            self.result,
                            path,
                            self.output_format.value,
                        )
                        content = await asyncio.to_thread(path.read_bytes)
                    await picker.save_file(
                        dialog_title="Baixar resultado",
                        file_name=file_name,
                        file_type=ft.FilePickerFileType.CUSTOM,
                        allowed_extensions=[extension],
                        src_bytes=content,
                    )
                else:
                    output = await picker.save_file(
                        dialog_title="Exportar resultado",
                        file_name=file_name,
                        file_type=ft.FilePickerFileType.CUSTOM,
                        allowed_extensions=[extension],
                    )
                    if not output:
                        return
                    path = Path(output)
                    if path.suffix.lower() != f".{extension}":
                        path = path.with_suffix(f".{extension}")
                    await asyncio.to_thread(
                        self.spec.export, self.result, path, self.output_format.value
                    )
        except Exception as error:
            LOGGER.exception("Falha na exportação %s", self.spec.key)
            self._show_error(error)
            self.set_status("Falha na exportação", 0)
        else:
            LOGGER.info(
                "Exportação concluída: key=%s format=%s duration=%.3fs",
                self.spec.key,
                self.output_format.value,
                time.monotonic() - started_at,
            )
            action = "baixado" if self.page.web else "exportado"
            self.set_status(f"Resultado {action}: {file_name}", 1)
            self.page.show_dialog(
                ft.SnackBar(
                    f"Resultado {action} com sucesso.",
                    bgcolor=GREEN,
                    show_close_icon=True,
                )
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
            self.details.controls = [
                ft.Text(
                    f"{result.company}  •  Competência {result.period_end:%m/%Y}  •  "
                    f"Proventos {format_value(result.earnings)}  •  Descontos {format_value(result.deductions)}  •  "
                    f"Líquido a pagar {format_value(result.net_payable)}  •  "
                    f"Férias: {result.vacation_employees} funcionário(s), "
                    f"{result.vacation_entries} lançamento(s) individualizado(s)  •  "
                    f"Excluídos: {result.ignored_rows} totalizadores e "
                    f"{result.excluded_rows} eventos duplicados do resumo mensal",
                    color=MUTED,
                    size=12,
                )
            ]
            return
        if self.spec.key == "entrada":
            not_posted = sum(row.entry_date is None for row in result.rows)
            self.details.controls = [
                ft.Text(
                    f"{not_posted} nota(s) ainda não lançada(s)  •  "
                    "CE: alerta de 6 a 10 dias e atraso a partir de 11 dias  •  "
                    "Demais estados: alerta de 20 a 29 dias e atraso a partir de 30 dias",
                    color=MUTED,
                    size=12,
                )
            ]
            return
        if self.spec.key == "receber":
            client_difference = (
                result.client_accounting_total - result.client_financial_total
            )
            self.details.controls = [
                self._comparison_line(
                    "Clientes",
                    "Balancete",
                    result.client_accounting_total,
                    "Posição por cliente",
                    result.client_financial_total,
                    client_difference,
                    "Conciliado"
                    if abs(client_difference) <= Decimal("0.01")
                    else "Divergente",
                ),
                self._comparison_line(
                    "Notas a faturar",
                    "Borderô",
                    result.billing.source_value,
                    "Razão a faturar",
                    result.billing.accounting_value,
                    result.billing.difference,
                    result.billing.status,
                ),
                self._comparison_line(
                    "Comissões",
                    "Agregados lançados",
                    result.commissions.source_value,
                    "Razão",
                    result.commissions.accounting_value,
                    result.commissions.difference,
                    result.commissions.status,
                ),
            ]
            return
        if self.spec.key == "receita":
            difference = result.cmflex_total - result.opera_total
            self.details.controls = [
                self._comparison_line(
                    "Receita",
                    "Contabilidade",
                    result.cmflex_total,
                    "Opera",
                    result.opera_total,
                    difference,
                    self._difference_status(difference),
                )
            ]
            return
        if self.spec.key == "cupons_hospede":
            journal_total = sum(
                (row.journal_value or Decimal() for row in result.coupons), Decimal()
            )
            difference = result.total_pdv - journal_total
            self.details.controls = [
                self._comparison_line(
                    "Cupons cobrados",
                    "BI/PDV",
                    result.total_pdv,
                    "Conta do hóspede",
                    journal_total,
                    difference,
                    self._difference_status(difference),
                )
            ]
            return
        if self.spec.key == "rps":
            opera = sum(
                (row.opera_value or Decimal() for row in result.rows), Decimal()
            )
            fiscal = sum(
                (row.fiscal_value or Decimal() for row in result.rows), Decimal()
            )
            city = sum((row.city_value or Decimal() for row in result.rows), Decimal())
            opera_fiscal = opera - fiscal
            fiscal_city = fiscal - city
            self.details.controls = [
                self._comparison_line(
                    "Integração Fiscal",
                    "Opera",
                    opera,
                    "Fiscal",
                    fiscal,
                    opera_fiscal,
                    self._difference_status(opera_fiscal),
                ),
                self._comparison_line(
                    "Emissão de NFS-e",
                    "Fiscal",
                    fiscal,
                    "Prefeitura",
                    city,
                    fiscal_city,
                    self._difference_status(fiscal_city),
                ),
            ]
            return
        if self.spec.key == "cupons":
            simphony_fiscal = result.simphony_total - result.fiscal_total
            fiscal_sefaz = result.fiscal_total - result.sefaz_total
            self.details.controls = [
                self._comparison_line(
                    "Integração Fiscal",
                    "Simphony",
                    result.simphony_total,
                    "Fiscal",
                    result.fiscal_total,
                    simphony_fiscal,
                    self._difference_status(simphony_fiscal),
                ),
                self._comparison_line(
                    "Integração SEFAZ",
                    "Fiscal",
                    result.fiscal_total,
                    "SEFAZ",
                    result.sefaz_total,
                    fiscal_sefaz,
                    self._difference_status(fiscal_sefaz),
                ),
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
                self._comparison_line(
                    "Notas de serviço",
                    "Prefeitura/Portal",
                    external_gross,
                    "CAP",
                    cap_gross,
                    gross_difference,
                    self._difference_status(gross_difference),
                ),
                self._comparison_line(
                    "ISS retido",
                    "Prefeitura",
                    city_iss,
                    "CAP",
                    cap_iss,
                    iss_difference,
                    self._difference_status(iss_difference),
                ),
            ]
            return
        if self.spec.key == "pagar":
            checks = (result.suppliers, result.advances, *result.taxes)
            self.details.controls = [
                self._comparison_line(
                    check.name,
                    "Financeiro",
                    check.financial,
                    "Contabilidade",
                    check.accounting,
                    check.difference,
                    check.status,
                )
                for check in checks
            ]
            return
        if self.spec.key == "custos":
            lines = []
            for analysis, source_name in (
                ("Entradas", "CAP"),
                ("Saldo final", "Inventário"),
            ):
                rows = [row for row in result if row.analysis == analysis]
                source = sum((row.source for row in rows), Decimal())
                accounting = sum((row.accounting for row in rows), Decimal())
                difference = source - accounting
                lines.append(
                    self._comparison_line(
                        analysis,
                        source_name,
                        source,
                        "Contabilidade",
                        accounting,
                        difference,
                        self._difference_status(difference),
                    )
                )
            self.details.controls = lines
            return
        details = result_details(result)
        self.details.controls = (
            [
                ft.Text(
                    "  •  ".join(f"{key}: {value}" for key, value in details.items()),
                    color=MUTED,
                    size=12,
                    max_lines=2,
                    overflow=ft.TextOverflow.ELLIPSIS,
                )
            ]
            if details
            else []
        )

    @staticmethod
    def _difference_status(difference: Decimal) -> str:
        return "Conciliado" if abs(difference) <= Decimal("0.01") else "Divergente"

    @staticmethod
    def _comparison_line(
        title: str,
        first_label: str,
        first_value,
        second_label: str,
        second_value,
        difference,
        status: str,
    ) -> ft.Text:
        color = GREEN if status_kind(status) == "ok" else RED
        return ft.Text(
            spans=[
                ft.TextSpan(
                    f"{title}: ",
                    style=ft.TextStyle(weight=ft.FontWeight.BOLD, color=TEXT),
                ),
                ft.TextSpan(f"{first_label} {format_value(first_value)}  •  "),
                ft.TextSpan(f"{second_label} {format_value(second_value)}  •  "),
                ft.TextSpan(f"Diferença {format_value(difference)}  •  "),
                ft.TextSpan(
                    status, style=ft.TextStyle(weight=ft.FontWeight.BOLD, color=color)
                ),
            ],
            color=MUTED,
            size=12,
        )

    def _render_table(self) -> None:
        records = self._filtered()
        viewport_height = self.page.height or 900
        page_size = max(
            self.MIN_PAGE_SIZE,
            min(25, int((viewport_height - 300) // 38)),
        )
        pages = max(1, math.ceil(len(records) / page_size))
        self.current_page = min(self.current_page, pages - 1)
        start = self.current_page * page_size
        visible_records = records[start : start + page_size]

        header = ft.Row(
            [
                self._table_cell(column, column.label, TEXT, header=True)
                for column in self.spec.columns
            ],
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
                    format_value(
                        status
                        if column.key == "status"
                        else record_value(record, column.key)
                    ),
                    color if column.key in {"status", "difference"} else TEXT,
                )
                for column in self.spec.columns
            ]
            rows.append(
                ft.Container(
                    content=ft.Row(
                        cells,
                        spacing=0,
                        vertical_alignment=ft.CrossAxisAlignment.CENTER,
                    ),
                    height=38,
                    bgcolor="#15191E" if index % 2 == 0 else "#121519",
                    border=ft.Border(bottom=ft.BorderSide(1, BORDER)),
                )
            )

        table = ft.Container(
            content=ft.Column(
                [
                    ft.Container(
                        header,
                        height=42,
                        bgcolor=CARD,
                        border=ft.Border(bottom=ft.BorderSide(1, "#46505D")),
                    ),
                    *rows,
                ],
                spacing=0,
            ),
            border=ft.Border.all(1, BORDER),
            border_radius=10,
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
            width=(
                None
                if len(self.spec.columns) <= 7
                else sum(
                    self._table_cell_width(column) for column in self.spec.columns
                )
            ),
            expand=len(self.spec.columns) <= 7,
        )
        if len(self.spec.columns) <= 7:
            self.table_host.content = ft.Column(
                [table],
                scroll=ft.ScrollMode.AUTO,
                expand=True,
                spacing=0,
                horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            )
        else:
            self.table_host.content = ft.Column(
                [
                    ft.Row(
                        [table],
                        scroll=ft.ScrollMode.AUTO,
                        vertical_alignment=ft.CrossAxisAlignment.START,
                    )
                ],
                scroll=ft.ScrollMode.AUTO,
                expand=True,
                spacing=0,
            )
        self.page_text.value = (
            f"Página {self.current_page + 1} de {pages} • {len(records)} registro(s)"
        )
        self.previous.disabled = self.current_page == 0
        self.next.disabled = self.current_page + 1 >= pages

    def _table_cell(
        self, column, value: str, color: str, *, header: bool = False
    ) -> ft.Container:
        alignment = (
            ft.Alignment.CENTER_RIGHT if column.numeric else ft.Alignment.CENTER_LEFT
        )
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
            width=(
                None
                if len(self.spec.columns) <= 7
                else self._table_cell_width(column)
            ),
            expand=(
                2
                if len(self.spec.columns) <= 7
                and column.key in {"identification", "detail"}
                else len(self.spec.columns) <= 7
            ),
        )

    @staticmethod
    def _table_cell_width(column) -> int:
        wide = {
            "description",
            "detail",
            "customer",
            "provider",
            "supplier",
            "organogram",
            "name",
            "guest",
        }
        medium = {
            "document",
            "key",
            "company",
            "source",
            "analysis",
            "status",
            "item",
            "comprador",
            "category",
        }
        if column.key in wide:
            return 300
        if column.key in medium:
            return 210
        return 150

    def _update_metrics(self) -> None:
        kinds = [status_kind(record_status(record)) for record in self.records]
        total, ok, info = len(kinds), kinds.count("ok"), kinds.count("info")
        error = total - ok - info
        labels = ("Registros", "Conciliados", "Pendências", "Informativos")
        values = (total, ok, error, info)
        if self.spec.key == "folha":
            labels = (
                "Lançamentos",
                "Prontos para importar",
                "De/para incompleto",
                "Exclusões aplicadas",
            )
            values = (
                total,
                self.result.ready,
                total - self.result.ready,
                self.result.ignored_rows + self.result.excluded_rows,
            )
            ok, error, info = self.result.ready, total - self.result.ready, 0
        elif self.spec.key == "entrada":
            labels = ("Notas", "Em dia", "Em atraso", "Em alerta")
            values = (total, ok, error, info)
        for title, value_control, label, amount in zip(
            self.card_titles, self.card_values, labels, values
        ):
            title.value = label
            value_control.value = str(amount)
        ratio = ok / total if total else 0
        self.quality_value.value = f"{ratio:.1%}"

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
        self.details_button.visible = False
        self.metrics.visible = False
        self.file_info.visible = True
        self.file_summary.value = "Arquivos necessários"
        self.file_hint.value = self._file_guidance()
        self.file_summary.tooltip = None
        self.details.controls = []
        for title, label in zip(
            self.card_titles, ("Registros", "Conciliados", "Pendências", "Informativos")
        ):
            title.value = label
        for value in self.card_values:
            value.value = "—"
        self.quality_value.value = "0%"
        self._render_table()
        self.set_status("Seleção limpa", 0)
        self.page.update()

    def _show_summary(self, _event=None) -> None:
        if self.result is None:
            return
        metrics = "  •  ".join(
            f"{title.value}: {value.value}"
            for title, value in zip(
                [*self.card_titles, self.quality_title],
                [*self.card_values, self.quality_value],
            )
        )
        content: list[ft.Control] = [
            ft.Text(self.file_summary.value, weight=ft.FontWeight.BOLD),
            ft.Text(self.file_summary.tooltip or "", color=MUTED, size=11),
            ft.Divider(color=BORDER),
            ft.Text(metrics, color=MUTED),
        ]
        if self.details.controls:
            content.extend([ft.Divider(color=BORDER), *self.details.controls])
        dialog = ft.AlertDialog(
            modal=True,
            title="Resumo da conferência",
            content=ft.Column(
                content, spacing=8, tight=True, scroll=ft.ScrollMode.AUTO
            ),
            actions=[ft.Button("Fechar", on_click=lambda _: self.page.pop_dialog())],
        )
        self.page.show_dialog(dialog)

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
        self.sidebar_animating = False
        self.content = ft.Container(expand=True)
        self.status_text = ft.Text("Pronto", color=MUTED, size=12)
        self.progress = ft.ProgressBar(
            value=0, color=BLUE, bgcolor="#343A45", bar_height=4
        )
        self.navigation_icons = (
            ft.Icons.COMPARE_ARROWS,
            ft.Icons.HOTEL,
            ft.Icons.PAYMENTS,
            ft.Icons.RECEIPT_LONG,
            ft.Icons.CLOUD_SYNC,
            ft.Icons.DESCRIPTION,
            ft.Icons.SCHEDULE,
            ft.Icons.POINT_OF_SALE,
            ft.Icons.REQUEST_QUOTE,
            ft.Icons.ACCOUNT_BALANCE_WALLET,
            ft.Icons.ACCOUNT_BALANCE,
            ft.Icons.INVENTORY_2,
        )
        self.sidebar_body = ft.Container(
            expand=True,
            opacity=1,
            animate_opacity=ft.Animation(
                140, ft.AnimationCurve.EASE_IN_OUT_CUBIC
            ),
        )
        self.sidebar = ft.Container(
            content=self.sidebar_body,
            width=126,
            bgcolor="#101318",
            border=ft.Border(right=ft.BorderSide(1, BORDER)),
            animate_size=ft.Animation(300, ft.AnimationCurve.EASE_IN_OUT_CUBIC),
            clip_behavior=ft.ClipBehavior.HARD_EDGE,
        )
        self._render_sidebar()

    def build(self) -> None:
        self.page.title = "HERMES"
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
            bgcolor="#111419",
            padding=ft.Padding(left=18, top=9, right=18, bottom=11),
        )
        body = ft.Row(
            [
                self.sidebar,
                ft.Container(
                    self.content,
                    padding=ft.Padding(left=14, top=14, right=14, bottom=10),
                    expand=True,
                ),
            ],
            spacing=0,
            expand=True,
        )
        self.page.add(ft.Column([body, footer], spacing=0, expand=True))
        self._show(0)

    def _show(self, index: int) -> None:
        self.selected_index = index
        self.current = AutomationView(self.page, SPECS[index], self.set_status)
        self.content.content = self.current.control
        self._render_sidebar()
        self.set_status(f"{SPECS[index].name} carregada", 0)
        self.page.update()

    async def _toggle_rail(self, _event=None) -> None:
        if self.sidebar_animating:
            return

        self.sidebar_animating = True
        self.sidebar_body.opacity = 0
        self.page.update()
        await asyncio.sleep(0.09)

        self.sidebar_expanded = not self.sidebar_expanded
        self.sidebar.width = 270 if self.sidebar_expanded else 126
        self._render_sidebar()
        self.page.update()
        await asyncio.sleep(0.08)

        self.sidebar_body.opacity = 1
        self.page.update()
        await asyncio.sleep(0.14)
        self.sidebar_animating = False

    def _show_about(self, _event=None) -> None:
        dialog = ft.AlertDialog(
            modal=True,
            title=ft.Row(
                [
                    ft.Image(
                        src=LOGO_SOURCE, width=34, height=34, fit=ft.BoxFit.CONTAIN
                    ),
                    ft.Text("Sobre o HERMES", weight=ft.FontWeight.BOLD),
                ],
                spacing=10,
                tight=True,
            ),
            content=ft.Column(
                [
                    ft.Text("Painel de Automação de Planilhas", color=MUTED),
                    ft.Text("Versão 1.0.0", size=12, color=MUTED),
                    ft.Divider(color=BORDER),
                    ft.Text("Desenvolvido por Magno Alves", weight=ft.FontWeight.W_600),
                ],
                spacing=8,
                tight=True,
                width=330,
            ),
            actions=[ft.Button("Fechar", on_click=lambda _: self.page.pop_dialog())],
        )
        self.page.show_dialog(dialog)

    def _render_sidebar(self) -> None:
        logo_size = 38 if self.sidebar_expanded else 28
        brand = ft.Row(
            [
                ft.Image(
                    src=LOGO_SOURCE,
                    width=logo_size,
                    height=logo_size,
                    fit=ft.BoxFit.CONTAIN,
                    tooltip="HERMES",
                ),
                ft.Text("HERMES", size=18, weight=ft.FontWeight.BOLD, color=TEXT),
            ],
            spacing=7 if self.sidebar_expanded else 5,
            tight=True,
        )
        toggle = ft.IconButton(
            ft.Icons.CHEVRON_LEFT if self.sidebar_expanded else ft.Icons.MENU,
            tooltip="Recolher menu" if self.sidebar_expanded else "Abrir menu",
            on_click=self._toggle_rail,
        )
        header = (
            ft.Row([brand, toggle], alignment=ft.MainAxisAlignment.SPACE_BETWEEN)
            if self.sidebar_expanded
            else ft.Column(
                [brand, toggle],
                horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                spacing=4,
            )
        )
        buttons = []
        for index, (spec, icon) in enumerate(zip(SPECS, self.navigation_icons)):
            selected = index == self.selected_index
            icon_control = ft.Icon(
                icon, size=20, color="#9BD5FA" if selected else MUTED
            )
            content = (
                ft.Row(
                    [
                        icon_control,
                        ft.Text(
                            spec.name, size=12, weight=ft.FontWeight.W_600, color=TEXT
                        ),
                    ],
                    spacing=12,
                    alignment=ft.MainAxisAlignment.START,
                )
                if self.sidebar_expanded
                else icon_control
            )
            buttons.append(
                ft.Container(
                    content=content,
                    height=43,
                    padding=ft.Padding(
                        left=18 if self.sidebar_expanded else 0,
                        top=0,
                        right=12 if self.sidebar_expanded else 0,
                        bottom=0,
                    ),
                    alignment=ft.Alignment.CENTER_LEFT
                    if self.sidebar_expanded
                    else ft.Alignment.CENTER,
                    bgcolor="#174B6D" if selected else ft.Colors.TRANSPARENT,
                    border_radius=10,
                    tooltip=None if self.sidebar_expanded else spec.name,
                    on_click=lambda _event, current=index: self._show(current),
                )
            )
        about_icon = ft.Icon(ft.Icons.INFO_OUTLINE, size=19, color=MUTED)
        about_content = (
            ft.Row(
                [
                    about_icon,
                    ft.Text("Sobre", size=12, color=MUTED),
                ],
                spacing=12,
                tight=True,
            )
            if self.sidebar_expanded
            else about_icon
        )
        about = ft.Container(
            content=about_content,
            height=38,
            padding=ft.Padding(
                left=18 if self.sidebar_expanded else 0,
                top=0,
                right=12 if self.sidebar_expanded else 0,
                bottom=0,
            ),
            alignment=ft.Alignment.CENTER_LEFT
            if self.sidebar_expanded
            else ft.Alignment.CENTER,
            border_radius=9,
            tooltip=None if self.sidebar_expanded else "Sobre o HERMES",
            on_click=self._show_about,
        )
        self.sidebar_body.content = ft.Column(
            [
                ft.Container(
                    header, padding=ft.Padding(left=16, top=14, right=10, bottom=8)
                ),
                ft.Divider(height=1, color=BORDER),
                ft.Column(buttons, spacing=5, scroll=ft.ScrollMode.AUTO, expand=True),
                ft.Divider(height=1, color=BORDER),
                ft.Container(
                    about, padding=ft.Padding(left=8, top=0, right=8, bottom=10)
                ),
            ],
            spacing=10,
            expand=True,
        )

    def set_status(self, message: str, progress: float | None) -> None:
        self.status_text.value = message
        self.progress.value = progress


def main(page: ft.Page) -> None:
    HermesApp(page).build()
