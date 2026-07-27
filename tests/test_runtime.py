import asyncio
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest import IsolatedAsyncioTestCase, TestCase
from unittest.mock import patch

from hermes_ui import runtime


class UploadValidationTest(TestCase):
    def test_counts_web_upload_bytes(self):
        files = [
            SimpleNamespace(name="a.xlsx", bytes=b"123", path=None),
            SimpleNamespace(name="b.xlsx", bytes=b"4567", path=None),
        ]

        self.assertEqual(runtime.validate_upload(files, web=True), 7)

    def test_rejects_missing_web_content(self):
        files = [SimpleNamespace(name="a.xlsx", bytes=None, path=None)]

        with self.assertRaisesRegex(ValueError, "não enviou"):
            runtime.validate_upload(files, web=True)

    def test_rejects_total_above_limit(self):
        files = [SimpleNamespace(name="a.xlsx", bytes=b"1234", path=None)]

        with patch.object(runtime, "MAX_TOTAL_UPLOAD_SIZE", 3):
            with self.assertRaisesRegex(ValueError, "limite total"):
                runtime.validate_upload(files, web=True)

    def test_counts_desktop_file_bytes(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "a.xlsx"
            path.write_bytes(b"12345")
            files = [SimpleNamespace(name=path.name, bytes=None, path=str(path))]

            self.assertEqual(runtime.validate_upload(files, web=False), 5)


class JobLimiterTest(IsolatedAsyncioTestCase):
    async def test_limits_concurrent_jobs(self):
        limiter = asyncio.Semaphore(2)
        active = 0
        peak = 0

        async def job():
            nonlocal active, peak
            async with limiter:
                active += 1
                peak = max(peak, active)
                await asyncio.sleep(0.01)
                active -= 1

        await asyncio.gather(*(job() for _ in range(6)))

        self.assertEqual(peak, 2)
