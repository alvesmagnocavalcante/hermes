from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from zipfile import ZipFile

from openpyxl import Workbook

from automations.download_danfe import (
    KeyEntry,
    download_batch,
    download_individual,
    download_keys,
    normalize_access_key,
    read_spreadsheet_keys,
)


def valid_key(seed: str = "1") -> str:
    base = seed * 43
    total = sum(
        int(number) * (2 + index % 8)
        for index, number in enumerate(reversed(base))
    )
    digit = 11 - total % 11
    return base + str(digit if digit < 10 else 0)


class FakeGateway:
    def __init__(self):
        self.closed = False

    def download(self, key: str, destination: Path) -> Path:
        path = destination / f"{key}.pdf"
        path.write_bytes(b"%PDF-test")
        return path

    def close(self) -> None:
        self.closed = True


class DanfeDownloadTest(TestCase):
    def test_validates_access_key_format_and_check_digit(self):
        key = valid_key()

        self.assertEqual(normalize_access_key(key), key)
        self.assertEqual(normalize_access_key(f"{key[:4]}.{key[4:] }"), key)
        with self.assertRaisesRegex(ValueError, "dígito verificador"):
            normalize_access_key(key[:-1] + str((int(key[-1]) + 1) % 10))

    def test_reads_key_column_and_reports_invalid_and_duplicate_rows(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "lote.xlsx"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Relatório de DANFEs"])
            sheet.append(["Nota", "Chave de acesso"])
            first = valid_key("1")
            second = valid_key("2")
            sheet.append([1, first])
            sheet.append([2, first])
            sheet.append([3, "123"])
            sheet.append([4, second])
            workbook.save(path)

            result = read_spreadsheet_keys(path)

        self.assertEqual([entry.key for entry in result.entries], [first, second])
        self.assertEqual(result.duplicates, 1)
        self.assertEqual(result.invalid[0].source_row, 5)
        self.assertEqual(result.column, "B")

    def test_downloads_keys_with_injected_gateway(self):
        gateway = FakeGateway()
        with TemporaryDirectory() as directory:
            result = download_keys(
                (KeyEntry(valid_key(), 2),),
                Path(directory),
                interval=0,
                gateway_factory=lambda: gateway,
            )

            self.assertTrue((Path(directory) / f"{valid_key()}.pdf").exists())

        self.assertEqual(result.downloaded, 1)
        self.assertEqual(result.failed, 0)
        self.assertTrue(gateway.closed)

    def test_reports_progress_after_each_processed_key(self):
        progress = []
        entries = (
            KeyEntry(valid_key("1"), 2),
            KeyEntry(valid_key("2"), 3),
        )

        with TemporaryDirectory() as directory:
            download_keys(
                entries,
                Path(directory),
                interval=0,
                gateway_factory=FakeGateway,
                progress_callback=progress.append,
            )

        self.assertEqual([item.processed for item in progress], [0, 1, 2])
        self.assertEqual([item.total for item in progress], [2, 2, 2])
        self.assertEqual(progress[-1].downloaded, 2)
        self.assertEqual(progress[-1].last_record.key, valid_key("2"))

    def test_batch_creates_zip_with_pdfs_and_results_csv(self):
        with TemporaryDirectory() as directory:
            directory_path = Path(directory)
            spreadsheet = directory_path / "lote.xlsx"
            output = directory_path / "danfes.zip"
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["CHAVE"])
            sheet.append([valid_key("1")])
            sheet.append([valid_key("2")])
            workbook.save(spreadsheet)

            result = download_batch(
                spreadsheet,
                output,
                gateway_factory=FakeGateway,
            )
            with ZipFile(output) as archive:
                names = set(archive.namelist())

        self.assertEqual(result.downloaded, 2)
        self.assertIn("resultados.csv", names)
        self.assertIn(f"{valid_key('1')}.pdf", names)
        self.assertIn(f"{valid_key('2')}.pdf", names)

    def test_individual_writes_requested_pdf(self):
        with TemporaryDirectory() as directory:
            output = Path(directory) / "nota.pdf"

            result = download_individual(
                valid_key(), output, gateway_factory=FakeGateway
            )

            self.assertEqual(output.read_bytes(), b"%PDF-test")
            self.assertEqual(result.downloaded, 1)
