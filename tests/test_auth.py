from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest import TestCase
from unittest.mock import MagicMock

import bcrypt

from hermes_ui.auth import (
    AuthenticationConfigurationError,
    LoginAttemptLimiter,
    verify_credentials,
)
from hermes_ui.app import AuthenticatedHermesSession


class CredentialVerificationTest(TestCase):
    def test_accepts_matching_bcrypt_htpasswd_entry(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / ".htpasswd"
            password_hash = bcrypt.hashpw(b"secret", bcrypt.gensalt(rounds=4))
            path.write_bytes(b"carmel:" + password_hash + b"\n")

            self.assertTrue(verify_credentials(path, "carmel", "secret"))
            self.assertFalse(verify_credentials(path, "carmel", "incorrect"))
            self.assertFalse(verify_credentials(path, "unknown", "secret"))

    def test_rejects_missing_credential_file(self):
        with self.assertRaises(AuthenticationConfigurationError):
            verify_credentials(Path("missing.htpasswd"), "carmel", "secret")


class LoginAttemptLimiterTest(TestCase):
    def test_blocks_after_limit_and_releases_after_window(self):
        current_time = [100.0]
        limiter = LoginAttemptLimiter(
            max_attempts=3,
            window_seconds=60,
            clock=lambda: current_time[0],
        )

        self.assertEqual(limiter.record_failure("client"), 0)
        self.assertEqual(limiter.record_failure("client"), 0)
        self.assertEqual(limiter.record_failure("client"), 60)
        self.assertEqual(limiter.retry_after("client"), 60)

        current_time[0] = 161.0
        self.assertEqual(limiter.retry_after("client"), 0)

    def test_success_clears_previous_failures(self):
        limiter = LoginAttemptLimiter(max_attempts=2)
        limiter.record_failure("client")
        limiter.record_success("client")

        self.assertEqual(limiter.retry_after("client"), 0)


class SessionCleanupTest(TestCase):
    def test_repeated_disconnect_clears_authentication_only_once(self):
        store = MagicMock()
        store.contains_key.side_effect = [True, False]
        session = object.__new__(AuthenticatedHermesSession)
        session.page = SimpleNamespace(session=SimpleNamespace(store=store))

        session._clear_authentication()
        session._clear_authentication()

        store.remove.assert_called_once()
