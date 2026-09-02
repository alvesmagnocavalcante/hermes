import os
from unittest import TestCase
from unittest.mock import patch

from hermes_ui import telemetry


class TelemetryConfigurationTest(TestCase):
    def tearDown(self):
        telemetry._configured = False

    @patch.dict(os.environ, {}, clear=True)
    @patch.object(telemetry.logfire, "configure")
    @patch.object(telemetry.logfire, "instrument_system_metrics")
    @patch.object(telemetry.logfire, "log")
    def test_uses_project_credentials_without_environment_token(
        self, log, instrument_system_metrics, configure
    ):
        self.assertTrue(telemetry.configure_telemetry())
        configure.assert_called_once()
        instrument_system_metrics.assert_called_once_with()
        log.assert_called_once()

    @patch.dict(
        os.environ,
        {"HERMES_TELEMETRY_ENABLED": "false"},
        clear=True,
    )
    @patch.object(telemetry.logfire, "configure")
    def test_remains_disabled_when_explicitly_disabled(self, configure):
        self.assertFalse(telemetry.configure_telemetry())
        configure.assert_not_called()

    @patch.dict(
        os.environ,
        {
            "LOGFIRE_TOKEN": "test-token",
            "LOGFIRE_ENVIRONMENT": "test",
            "HERMES_LOGFIRE_SYSTEM_METRICS": "true",
        },
        clear=True,
    )
    @patch.object(telemetry.logfire, "log")
    @patch.object(telemetry.logfire, "instrument_system_metrics")
    @patch.object(telemetry.logfire, "configure")
    def test_configures_project_export_and_basic_system_metrics(
        self, configure, instrument_system_metrics, log
    ):
        self.assertTrue(telemetry.configure_telemetry())

        configure.assert_called_once_with(
            send_to_logfire="if-token-present",
            service_name="hermes",
            service_version="0.1.0",
            environment="test",
            console=False,
        )
        instrument_system_metrics.assert_called_once_with()
        self.assertEqual(log.call_count, 1)

    @patch.dict(os.environ, {"LOGFIRE_TOKEN": "test-token"}, clear=True)
    @patch.object(telemetry.logfire, "configure", side_effect=RuntimeError("offline"))
    def test_configuration_failure_does_not_block_application(self, _configure):
        self.assertFalse(telemetry.configure_telemetry())
        self.assertFalse(telemetry._configured)


class TelemetryEventTest(TestCase):
    def tearDown(self):
        telemetry._configured = False

    @patch.object(telemetry.logfire, "log")
    def test_automation_event_contains_only_operational_metadata(self, log):
        telemetry._configured = True

        telemetry.automation_event(
            automation="cupons",
            automation_name="Conferência dos Cupons",
            hotel="Cumbuco",
            success=True,
            file_count=3,
            total_bytes=1024,
            record_count=25,
            reconciled_count=20,
            pending_count=4,
            informational_count=1,
            duration_seconds=1.23456,
            execution_mode="web",
        )

        level, message, attributes = log.call_args.args
        self.assertEqual(level, "info")
        self.assertIn("{automation_name}", message)
        self.assertEqual(attributes["duration_seconds"], 1.235)
        self.assertEqual(attributes["automation_name"], "Conferência dos Cupons")
        self.assertEqual(attributes["quality_percent"], 80.0)
        self.assertEqual(attributes["pending_count"], 4)
        self.assertNotIn("error_detail", attributes)

    def test_sanitizes_sensitive_error_details(self):
        error = ValueError(
            r"Falha em C:\Users\usuario\Downloads\arquivo.xlsx "
            "token=segredo 23260727708448000110650010000000000000000000"
        )

        message = telemetry.safe_error_message(error)

        self.assertIn("[caminho omitido]", message)
        self.assertIn("token=[omitido]", message)
        self.assertIn("[identificador omitido]", message)
        self.assertNotIn("segredo", message)

    @patch.object(telemetry.logfire, "log")
    def test_failure_contains_readable_and_sanitized_reason(self, log):
        telemetry._configured = True

        telemetry.automation_event(
            automation="folha",
            automation_name="Lançamento da Folha de Pagamento",
            hotel="Não se aplica",
            success=False,
            file_count=6,
            duration_seconds=0.078,
            execution_mode="web",
            error_type="ValueError",
            error=ValueError("Relatório de IRRF não encontrado."),
        )

        level, message, attributes = log.call_args.args
        self.assertEqual(level, "error")
        self.assertIn("{error_detail}", message)
        self.assertEqual(attributes["result"], "falha")
        self.assertEqual(
            attributes["error_detail"], "Relatório de IRRF não encontrado."
        )
