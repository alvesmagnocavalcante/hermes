from __future__ import annotations

import base64
import csv
import logging
import os
import re
import shutil
import time
import zipfile
from collections.abc import Callable, Iterable
from contextlib import closing
from dataclasses import dataclass
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Protocol

from DrissionPage import Chromium, ChromiumOptions
from DrissionPage.errors import ElementLostError, NoRectError

from automations.common import normalize_key
from automations.excel_reader import load_workbook_compatible as load_workbook


URL_CONSULTA = "https://consultadanfe.com/"
CAMPO_CHAVE = 'xpath://*[@id="accessKey"]'
BOTAO_IMPRIMIR = (
    'xpath://*[@id="main-content"]/section[1]/div[2]/div/div[1]/form[1]/div/button'
)
BOTAO_VISUALIZAR = (
    "xpath:/html/body/div[3]/div[2]/div/div[2]/div[1]/div[2]/button[1]"
)
KEY_HEADERS = {
    "CHAVE",
    "CHAVEACESSO",
    "CHAVEDEACESSO",
    "CHAVENFE",
    "CHAVEFISCAL",
}
LOG = logging.getLogger(__name__)


@dataclass(frozen=True)
class KeyEntry:
    key: str
    row: int


@dataclass(frozen=True)
class DanfeRecord:
    key: str
    status: str
    detail: str
    source_row: int | None = None


@dataclass(frozen=True)
class SpreadsheetKeys:
    entries: tuple[KeyEntry, ...]
    invalid: tuple[DanfeRecord, ...]
    duplicates: int
    sheet: str
    column: str


@dataclass(frozen=True)
class DanfeResult:
    rows: tuple[DanfeRecord, ...]
    requested: int
    downloaded: int
    failed: int
    invalid: int
    duplicates: int = 0


@dataclass(frozen=True)
class DanfeProgress:
    processed: int
    total: int
    downloaded: int
    failed: int
    invalid: int
    duplicates: int
    last_record: DanfeRecord | None = None


class DanfeGateway(Protocol):
    def download(self, key: str, destination: Path) -> Path: ...

    def close(self) -> None: ...


def normalize_access_key(value: object) -> str:
    """Valida formato e dígito verificador de uma chave NF-e numérica."""
    if isinstance(value, int):
        raw = str(value)
    elif isinstance(value, str):
        raw = re.sub(r"[\s.\-]", "", value)
    else:
        raw = ""
    if not re.fullmatch(r"[0-9]{44}", raw):
        raise ValueError("a chave deve conter exatamente 44 dígitos")
    total = sum(int(number) * (2 + index % 8) for index, number in enumerate(reversed(raw[:-1])))
    digit = 11 - total % 11
    expected = digit if digit < 10 else 0
    if int(raw[-1]) != expected:
        raise ValueError("dígito verificador inválido")
    return raw


def _column_letter(index: int) -> str:
    result = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        result = chr(65 + remainder) + result
    return result


def read_spreadsheet_keys(path: Path, header_scan_limit: int = 20) -> SpreadsheetKeys:
    """Localiza automaticamente a coluna de chave e retorna valores únicos válidos."""
    with closing(load_workbook(path, data_only=False, read_only=True)) as workbook:
        selected = None
        for sheet in workbook.worksheets:
            for row_number, row in enumerate(
                sheet.iter_rows(min_row=1, max_row=header_scan_limit, values_only=True),
                start=1,
            ):
                for column_number, value in enumerate(row, start=1):
                    if normalize_key(str(value or "")) in KEY_HEADERS:
                        selected = (sheet, row_number, column_number)
                        break
                if selected:
                    break
            if selected:
                break
        if not selected:
            raise ValueError(
                "Não foi encontrada uma coluna CHAVE, CHAVE DE ACESSO ou CHAVE NF-E."
            )

        sheet, header_row, column_number = selected
        entries: list[KeyEntry] = []
        invalid: list[DanfeRecord] = []
        seen: set[str] = set()
        duplicates = 0
        rows = sheet.iter_rows(
            min_row=header_row + 1,
            min_col=column_number,
            max_col=column_number,
            values_only=True,
        )
        for row_number, (value,) in enumerate(rows, start=header_row + 1):
            if value is None or str(value).strip() == "":
                continue
            try:
                key = normalize_access_key(value)
            except ValueError as error:
                invalid.append(
                    DanfeRecord("", "Chave inválida", str(error), row_number)
                )
                continue
            if key in seen:
                duplicates += 1
                continue
            seen.add(key)
            entries.append(KeyEntry(key, row_number))

        return SpreadsheetKeys(
            tuple(entries),
            tuple(invalid),
            duplicates,
            sheet.title,
            _column_letter(column_number),
        )


def _valid_pdf(path: Path) -> bool:
    if not path.is_file() or path.stat().st_size <= 5:
        return False
    with path.open("rb") as stream:
        return stream.read(5) == b"%PDF-"


def _masked_key(key: str) -> str:
    return f"***{key[-6:]}"


class BrowserDanfeGateway:
    """Adaptador isolado para o site externo Consulta DANFE."""

    def __init__(
        self,
        timeout: float = 30,
        note_timeout: float = 90,
        pause: float = 1,
    ):
        self.timeout = timeout
        self.note_timeout = note_timeout
        self.pause = pause
        options = ChromiumOptions(read_file=False).auto_port().headless()
        browser_path = os.environ.get("HERMES_CHROMIUM_PATH") or next(
            (
                path
                for executable in ("chromium", "chromium-browser", "google-chrome")
                if (path := shutil.which(executable))
            ),
            None,
        )
        if browser_path:
            options.set_browser_path(browser_path)
        for argument in (
            "--no-sandbox",
            "--disable-dev-shm-usage",
            "--disable-gpu",
            "--window-size=1280,900",
        ):
            options.set_argument(argument)
        self.browser = Chromium(options)
        self.tab = self.browser.latest_tab
        self.tab.set.timeouts(base=timeout, page_load=timeout)
        self.tab.set.when_download_file_exists("overwrite")

    @staticmethod
    def _interact(tab, selector: str, timeout: float, action):
        limit = time.monotonic() + timeout
        while time.monotonic() < limit:
            try:
                for element in tab.eles(selector, timeout=0):
                    if element.states.is_clickable:
                        return action(element)
            except (ElementLostError, NoRectError):
                pass
            time.sleep(min(0.2, max(0, limit - time.monotonic())))
        raise TimeoutError(
            "o site não disponibilizou o campo ou botão esperado; "
            "verifique indisponibilidade ou CAPTCHA"
        )

    def download(self, key: str, destination: Path) -> Path:
        destination.mkdir(parents=True, exist_ok=True)
        output = destination / f"{key}.pdf"
        limit = time.monotonic() + self.note_timeout

        def remaining() -> float:
            seconds = limit - time.monotonic()
            if seconds <= 0:
                raise TimeoutError(
                    f"a consulta excedeu {self.note_timeout:g} segundos"
                )
            return seconds

        def wait_time() -> float:
            return min(self.timeout, remaining())

        if not self.tab.get(URL_CONSULTA, timeout=wait_time()):
            raise RuntimeError("não foi possível carregar o Consulta DANFE")
        time.sleep(min(self.pause, remaining()))
        self._interact(
            self.tab,
            CAMPO_CHAVE,
            wait_time(),
            lambda element: element.input(key, clear=True),
        )
        time.sleep(min(self.pause, remaining()))
        clicked = self._interact(
            self.tab,
            BOTAO_IMPRIMIR,
            wait_time(),
            lambda element: element.click(by_js=True),
        )
        if not clicked:
            raise RuntimeError("não foi possível solicitar a impressão do DANFE")
        time.sleep(min(self.pause, remaining()))
        previous_tabs = {tab.tab_id for tab in self.browser.get_tabs()}
        clicked = self._interact(
            self.tab,
            BOTAO_VISUALIZAR,
            wait_time(),
            lambda element: element.click(by_js=True),
        )
        if not clicked:
            raise RuntimeError("não foi possível abrir o PDF gerado")

        # O site cria o DANFE como uma URL blob: em uma nova aba. O blob só
        # existe na sessão que o gerou e é validado antes de ser persistido.
        pdf_tab = None
        while time.monotonic() < limit:
            pdf_tab = next(
                (
                    tab
                    for tab in self.browser.get_tabs()
                    if tab.tab_id not in previous_tabs and tab.url.startswith("blob:")
                ),
                None,
            )
            if pdf_tab:
                break
            time.sleep(0.1)
        if not pdf_tab:
            raise TimeoutError("o site não abriu a visualização do PDF")

        try:
            data_url = self.tab.run_js(
                """
                return fetch(arguments[0])
                    .then(response => {
                        if (!response.ok) throw new Error(`HTTP ${response.status}`);
                        return response.blob();
                    })
                    .then(blob => new Promise((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onload = () => resolve(reader.result);
                        reader.onerror = () => reject(reader.error);
                        reader.readAsDataURL(blob);
                    }));
                """,
                pdf_tab.url,
                timeout=wait_time(),
            )
            if not isinstance(data_url, str) or "," not in data_url:
                raise RuntimeError("o site retornou um PDF ilegível")
            output.write_bytes(base64.b64decode(data_url.split(",", 1)[1]))
            if not _valid_pdf(output):
                output.unlink(missing_ok=True)
                raise RuntimeError("o conteúdo recebido não é um PDF válido")
            return output
        finally:
            pdf_tab.close()

    def close(self) -> None:
        self.browser.quit()


GatewayFactory = Callable[[], DanfeGateway]
ProgressCallback = Callable[[DanfeProgress], None]


def _notify_progress(
    callback: ProgressCallback | None,
    rows: list[DanfeRecord],
    total: int,
    invalid: int,
    duplicates: int,
    last_record: DanfeRecord | None = None,
) -> None:
    if callback is None:
        return
    progress = DanfeProgress(
        processed=len(rows),
        total=total,
        downloaded=sum(row.status == "Baixado" for row in rows),
        failed=sum(row.status == "Falha" for row in rows),
        invalid=invalid,
        duplicates=duplicates,
        last_record=last_record,
    )
    try:
        callback(progress)
    except Exception:
        LOG.exception("Falha ao publicar o andamento do lote de DANFEs")


def download_keys(
    entries: Iterable[KeyEntry],
    destination: Path,
    *,
    invalid: Iterable[DanfeRecord] = (),
    duplicates: int = 0,
    attempts: int = 2,
    interval: float = 2,
    gateway_factory: GatewayFactory = BrowserDanfeGateway,
    progress_callback: ProgressCallback | None = None,
) -> DanfeResult:
    """Baixa chaves em uma única sessão e devolve resultado auditável."""
    items = tuple(entries)
    invalid_rows = tuple(invalid)
    rows: list[DanfeRecord] = list(invalid_rows)
    total = len(items) + len(invalid_rows)
    _notify_progress(
        progress_callback,
        rows,
        total,
        len(invalid_rows),
        duplicates,
    )
    if not items:
        return DanfeResult(
            rows=tuple(rows),
            requested=0,
            downloaded=0,
            failed=0,
            invalid=len(rows),
            duplicates=duplicates,
        )

    gateway = gateway_factory()
    try:
        for index, entry in enumerate(items):
            if index:
                time.sleep(interval)
            error_detail = "falha não identificada"
            for attempt in range(1, attempts + 1):
                try:
                    path = gateway.download(entry.key, destination)
                    rows.append(
                        DanfeRecord(entry.key, "Baixado", path.name, entry.row)
                    )
                    LOG.info("DANFE baixado: %s", _masked_key(entry.key))
                    break
                except Exception as error:  # A falha externa precisa virar resultado.
                    error_detail = str(error)
                    LOG.warning(
                        "Falha no DANFE %s, tentativa %d/%d: %s",
                        _masked_key(entry.key),
                        attempt,
                        attempts,
                        error_detail,
                    )
                    if attempt < attempts:
                        time.sleep(interval)
            else:
                rows.append(
                    DanfeRecord(entry.key, "Falha", error_detail, entry.row)
                )
            _notify_progress(
                progress_callback,
                rows,
                total,
                len(invalid_rows),
                duplicates,
                rows[-1],
            )
    finally:
        gateway.close()

    downloaded = sum(row.status == "Baixado" for row in rows)
    failed = sum(row.status == "Falha" for row in rows)
    return DanfeResult(
        rows=tuple(rows),
        requested=len(items),
        downloaded=downloaded,
        failed=failed,
        invalid=len(invalid_rows),
        duplicates=duplicates,
    )


def _write_results_csv(result: DanfeResult, path: Path) -> None:
    with path.open("w", newline="", encoding="utf-8-sig") as stream:
        writer = csv.writer(stream, delimiter=";")
        writer.writerow(["linha", "chave", "status", "detalhe"])
        for row in result.rows:
            writer.writerow([row.source_row or "", row.key, row.status, row.detail])


def download_batch(
    spreadsheet: Path,
    output_zip: Path,
    *,
    batch_limit: int | None = None,
    gateway_factory: GatewayFactory = BrowserDanfeGateway,
    progress_callback: ProgressCallback | None = None,
) -> DanfeResult:
    batch = read_spreadsheet_keys(spreadsheet)
    if not batch.entries:
        raise ValueError("A planilha não contém nenhuma chave válida para consulta.")
    try:
        limit = batch_limit or int(os.environ.get("HERMES_DANFE_BATCH_LIMIT", "300"))
    except ValueError as error:
        raise ValueError("HERMES_DANFE_BATCH_LIMIT deve ser um inteiro positivo") from error
    if limit <= 0:
        raise ValueError("O limite por lote deve ser um inteiro positivo.")
    if len(batch.entries) > limit:
        raise ValueError(
            f"A planilha contém {len(batch.entries)} chaves únicas; "
            f"o limite por lote é {limit}."
        )
    output_zip.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="hermes_danfe_") as directory:
        downloads = Path(directory)
        result = download_keys(
            batch.entries,
            downloads,
            invalid=batch.invalid,
            duplicates=batch.duplicates,
            gateway_factory=gateway_factory,
            progress_callback=progress_callback,
        )
        csv_path = downloads / "resultados.csv"
        _write_results_csv(result, csv_path)
        with zipfile.ZipFile(
            output_zip, "w", compression=zipfile.ZIP_DEFLATED
        ) as archive:
            archive.write(csv_path, csv_path.name)
            for record in result.rows:
                if record.status != "Baixado":
                    continue
                pdf = downloads / record.detail
                archive.write(pdf, pdf.name)
    return result


def download_individual(
    key: str,
    output_pdf: Path,
    *,
    gateway_factory: GatewayFactory = BrowserDanfeGateway,
) -> DanfeResult:
    normalized = normalize_access_key(key)
    output_pdf.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix="hermes_danfe_") as directory:
        result = download_keys(
            (KeyEntry(normalized, 1),),
            Path(directory),
            gateway_factory=gateway_factory,
        )
        if result.downloaded != 1:
            detail = result.rows[-1].detail if result.rows else "download não concluído"
            raise RuntimeError(detail)
        (Path(directory) / f"{normalized}.pdf").replace(output_pdf)
    return result
