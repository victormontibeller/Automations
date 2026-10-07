"""Operational contracts: synthetic inputs and fake transports, never Google."""
from contextlib import redirect_stderr, redirect_stdout
from dataclasses import replace
from decimal import Decimal
from io import StringIO
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from uuid import UUID

from resumos_cartoes import service
from resumos_cartoes.config import Settings
from resumos_cartoes.models import Recipient, Report


class RunResultTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.settings = Settings(self.root, "local", "Synthetic", self.root / "output",
                                 self.root / "state", self.root / "fake-token.json",
                                 "Synthetic", (), "Synthetic Owner", personal_copy_email="copy@example.com")
        self.reports = [Report("black", "2026-09", Recipient("Person01", "Participant01", "one@example.com"),
                               (), (Decimal("-1.00"),), "Synthetic Owner")]
        self.reports.append(replace(self.reports[0], recipient=Recipient("Person02", "Participant02", "two@example.com"), payments=()))
        self.stdout, self.stderr = StringIO(), StringIO()
        for context in (
            patch.object(service.config, "load_settings", return_value=self.settings),
            patch.object(service, "input_workbook_bytes", return_value=b"synthetic"),
            patch.object(service.workbook, "read_reports", return_value=self.reports),
            redirect_stdout(self.stdout), redirect_stderr(self.stderr),
        ):
            context.__enter__()
            self.addCleanup(context.__exit__, None, None, None)

    def run_result(self, **kwargs):
        self.assertTrue(callable(getattr(service, "run_result", None)), "Structured run_result API is missing")
        return service.run_result(self.root / "fake-config.json", card="black", requested_month="2026-09", **kwargs)

    def test_preview_result_and_legacy_wrapper_keep_human_contract(self):
        with patch.object(service.delivery, "send_reports") as send, patch.object(service.alerts, "notify_owner") as alert:
            result = self.run_result()
            self.assertEqual(result.status, "succeeded")
            self.assertEqual(result.mode, "preview")
            self.assertEqual((result.card, result.month), ("black", "2026-09"))
            self.assertEqual((result.sent, result.skipped, result.no_movement, result.previews), (None, None, 1, 1))
            self.assertEqual(result.exit_code, 0)
            self.assertEqual(UUID(result.run_id).version, 4)
            self.assertGreaterEqual(result.elapsed_seconds, 0)
            self.assertEqual(result.alert_status, "not_requested")
            self.assertIsNone(result.error_category)
            text = self.stdout.getvalue()
            self.stdout.seek(0)
            self.stdout.truncate()
            self.assertEqual(service.run(self.root / "fake-config.json", card="black", requested_month="2026-09"), 0)
            self.assertEqual(self.stdout.getvalue(), text)
            self.assertEqual(self.stderr.getvalue(), "")
            send.assert_not_called()
            alert.assert_not_called()
            self.assertFalse(self.settings.state_dir.exists())

    def test_opt_in_json_correlates_runs_without_changing_stdout_or_handlers(self):
        import json
        import logging
        from resumos_cartoes import cli
        handlers = tuple(logging.getLogger().handlers)
        argv = ["--config", str(self.root / "fake-config.json"), "--card", "black", "--month", "2026-09"]
        self.assertEqual(cli.main(argv), 0)
        expected = self.stdout.getvalue()
        self.stdout.seek(0)
        self.stdout.truncate()
        self.assertEqual(self.stderr.getvalue(), "")
        for _ in range(2):
            try:
                code = cli.main([*argv, "--log-json"])
            except SystemExit as exc:
                self.fail(f"--log-json is not supported: exit {exc.code}")
            self.assertEqual(code, 0)
        self.assertEqual(self.stdout.getvalue(), expected * 2)
        records = [json.loads(line) for line in self.stderr.getvalue().splitlines()]
        self.assertEqual([r["event"] for r in records], ["run_started", "run_finished"] * 2)
        self.assertEqual(records[0]["run_id"], records[1]["run_id"])
        self.assertEqual(records[2]["run_id"], records[3]["run_id"])
        self.assertNotEqual(records[0]["run_id"], records[2]["run_id"])
        self.assertEqual(records[1]["status"], "succeeded")
        self.assertEqual(records[1]["previews"], 1)
        self.assertEqual(records[1]["mode"], "preview")
        self.assertEqual(tuple(logging.getLogger().handlers), handlers)

    def test_valid_period_survives_config_failure_without_exposing_payload(self):
        from resumos_cartoes.errors import ReportError
        with patch.object(service.config, "load_settings", side_effect=ReportError("PRIVATE_TOKEN private@example.com")):
            result = self.run_result(log_json=True)
        self.assertEqual((result.card, result.month), ("black", "2026-09"))
        self.assertEqual(result.status, "failed")
        self.assertEqual(result.error_category, "report_error")
        self.assertIsNone(result.sent)
        self.assertIsNone(result.previews)
        self.assertNotIn("PRIVATE_TOKEN", str(result))
        self.assertNotIn("PRIVATE_TOKEN", "\n".join(line for line in self.stderr.getvalue().splitlines() if line.startswith("{")))

    def test_invalid_api_card_fails_before_input_or_delivery_and_is_not_logged(self):
        import json
        with patch.object(service, "input_workbook_bytes") as read, patch.object(service.delivery, "send_reports") as send:
            result = service.run_result(self.root / "SECRET_CONFIG", card="SECRET_CARD", requested_month="2026-09", log_json=True)
        self.assertEqual(result.status, "failed")
        self.assertIsNone(result.card)
        read.assert_not_called()
        send.assert_not_called()
        records = [json.loads(line) for line in self.stderr.getvalue().splitlines() if line.startswith("{")]
        self.assertEqual(len(records), 2)
        self.assertNotIn("SECRET", json.dumps(records))

    def records(self):
        import json
        return [json.loads(line) for line in self.stderr.getvalue().splitlines() if line.startswith("{")]

    def fake_provider(self):
        from unittest.mock import MagicMock
        api = MagicMock()
        api.users.return_value.getProfile.return_value.execute.return_value = {"emailAddress": "owner@example.com"}
        api.users.return_value.messages.return_value.send.return_value.execute.return_value = {"id": "PRIVATE_PROVIDER_ID"}
        return api

    def test_successful_sends_and_repeated_skips_are_counts_not_delivery_receipts(self):
        api = self.fake_provider()
        with patch("automation_core.google_auth.build_service", return_value=api) as build:
            first = self.run_result(send=True, log_json=True)
            second = self.run_result(send=True, log_json=True)
        self.assertEqual((first.status, first.mode, first.sent, first.skipped, first.no_movement), ("succeeded", "send", 1, 0, 1))
        self.assertEqual((second.sent, second.skipped), (0, 1))
        self.assertIsNone(first.previews)
        build.assert_called_once()
        api.users.return_value.messages.return_value.send.return_value.execute.assert_called_once_with(num_retries=0)
        self.assertNotIn("PRIVATE_PROVIDER_ID", str(self.records()))
        self.assertEqual([r["sent"] for r in self.records() if r["event"] == "run_finished"], [1, 0])

    def test_test_mode_is_labelled_and_never_creates_production_ledger(self):
        api = self.fake_provider()
        with patch("automation_core.google_auth.build_service", return_value=api):
            result = self.run_result(send=True, test_to="test@example.com", log_json=True)
        self.assertEqual((result.mode, result.sent, result.skipped), ("test", 1, 0))
        self.assertTrue(all(r["mode"] == "test" for r in self.records()))
        self.assertFalse((self.settings.state_dir / "deliveries.sqlite3").exists())
        self.assertNotIn("test@example.com", str(self.records()))

    def test_partial_failure_counts_are_unknown_and_alert_outcomes_are_safe(self):
        import sqlite3
        self.reports[1] = replace(self.reports[1], payments=(Decimal("-2.00"),))
        for alert_fails in (False, True):
            with self.subTest(alert_fails=alert_fails):
                api = self.fake_provider()
                private = "SECRET_TOKEN private@example.com Name Subject 123.45 /private/config"
                hostile_error = type("SECRET_EXCEPTION_CLASS", (RuntimeError,), {})(private)
                request = api.users.return_value.messages.return_value.send.return_value
                request.execute.side_effect = [{"id": "PRIVATE_PROVIDER_ID"}, hostile_error]
                settings = replace(self.settings, state_dir=self.root / str(alert_fails))
                self.stderr.seek(0)
                self.stderr.truncate()
                with patch.object(service.config, "load_settings", return_value=settings), patch("automation_core.google_auth.build_service", return_value=api), patch.object(service.alerts, "notify_owner", side_effect=hostile_error if alert_fails else None) as alert:
                    result = self.run_result(send=True, log_json=True)
                self.assertEqual((result.status, result.exit_code, result.error_category), ("failed", 1, "report_error"))
                self.assertEqual((result.sent, result.skipped, result.no_movement), (None, None, 0))
                self.assertEqual(result.alert_status, "failed" if alert_fails else "accepted")
                alert.assert_called_once()
                self.assertEqual(request.execute.call_count, 2)
                self.assertEqual(len(self.records()), 2)
                for token in ("SECRET", "PRIVATE", "private@example.com", "123.45", "/private/config", "Person01", "Person02"):
                    self.assertNotIn(token, str(self.records()) + str(result))
                with sqlite3.connect(settings.state_dir / "deliveries.sqlite3") as db:
                    self.assertEqual(db.execute("SELECT status FROM deliveries ORDER BY sheet").fetchall(), [("sent",), ("unknown",)])
                db.close()

    def test_logging_write_flush_and_serialization_failures_do_not_change_acceptance(self):
        from resumos_cartoes import operations
        from contextlib import nullcontext
        class BrokenSink(StringIO):
            def fileno(self):
                raise AssertionError("Custom sinks must keep their stream interface")
            def write(self, text):
                if fault == "write":
                    raise OSError("SECRET_SINK")
                return super().write(text)
            def flush(self):
                if fault == "flush":
                    raise OSError("SECRET_SINK")
        for fault in ("write", "flush", "closed", "serialization", "fd-write", "fd-serialization"):
            with self.subTest(fault=fault):
                api = self.fake_provider()
                settings = replace(self.settings, state_dir=self.root / fault)
                sink = tempfile.TemporaryFile(mode="w+", encoding="utf-8") if fault.startswith("fd-") else BrokenSink()
                self.addCleanup(sink.close)
                if fault == "closed":
                    sink.close()
                failure = (patch.object(operations.json, "dumps", side_effect=ValueError("SECRET_JSON"))
                           if fault.endswith("serialization") else
                           patch.object(operations.os, "write", side_effect=OSError("SECRET_FD"))
                           if fault == "fd-write" else nullcontext())
                with patch.object(service.config, "load_settings", return_value=settings), patch("automation_core.google_auth.build_service", return_value=api), patch.object(service.alerts, "notify_owner") as alert, redirect_stderr(sink), failure:
                    result = self.run_result(send=True, log_json=True)
                self.assertEqual((result.status, result.sent, result.exit_code), ("succeeded", 1, 0))
                alert.assert_not_called()
                api.users.return_value.messages.return_value.send.assert_called_once()
                with patch.object(service.config, "load_settings", return_value=settings), patch("automation_core.google_auth.build_service") as connect:
                    repeated = self.run_result(send=True)
                self.assertEqual((repeated.sent, repeated.skipped), (0, 1))
                connect.assert_not_called()

    def test_acknowledged_send_then_ledger_failure_is_not_reported_as_zero(self):
        import sqlite3
        api = self.fake_provider()
        with patch("automation_core.google_auth.build_service", return_value=api), patch.object(service.ledger.Ledger, "finish", side_effect=sqlite3.OperationalError("SECRET_DB")), patch.object(service.alerts, "notify_owner"):
            result = self.run_result(send=True, log_json=True)
        self.assertEqual(result.status, "failed")
        self.assertIsNone(result.sent)
        self.assertIsNone(result.skipped)
        self.assertEqual(result.error_category, "operational_error")
        self.assertNotIn("SECRET", str(self.records()))
        api.users.return_value.messages.return_value.send.assert_called_once()

    def test_maintenance_counts_are_not_applicable_and_no_alert_is_sent(self):
        with patch.object(service.ledger, "maintain_ledger") as maintain, patch.object(service.alerts, "notify_owner") as alert:
            for options, mode in (({"history": True}, "history"), ({"resolve": "sent", "recipient": "Person01"}, "resolve")):
                result = self.run_result(log_json=True, **options)
                self.assertEqual((result.status, result.mode), ("succeeded", mode))
                self.assertEqual((result.sent, result.skipped, result.previews, result.no_movement), (None,) * 4)
            self.assertEqual(maintain.call_count, 2)
            alert.assert_not_called()

    def test_bad_month_and_hostile_report_errors_have_only_fixed_categories(self):
        from resumos_cartoes.errors import ReportError
        for value in ("SECRET_MONTH", "2026-13", "0000-01", "2026-09\nSECRET"):
            self.stderr.seek(0)
            self.stderr.truncate()
            result = service.run_result(self.root / "SECRET_CONFIG", card="black", requested_month=value, log_json=True)
            self.assertEqual(result.status, "failed")
            self.assertIsNone(result.month)
            self.assertNotIn("SECRET", str(self.records()))
        self.stderr.seek(0)
        self.stderr.truncate()
        with patch.object(service.workbook, "read_reports", side_effect=ReportError("SECRET_NAME SECRET_EMAIL SECRET_AMOUNT")):
            result = self.run_result(log_json=True)
        self.assertEqual(result.error_category, "report_error")
        self.assertNotIn("SECRET", str(self.records()))

    def test_empty_send_reports_known_zero_without_a_provider_connection(self):
        self.reports[0] = replace(self.reports[0], payments=())
        with patch("automation_core.google_auth.build_service") as connect:
            result = self.run_result(send=True, log_json=True)
        self.assertEqual((result.status, result.sent, result.skipped, result.no_movement), ("succeeded", 0, 0, 2))
        connect.assert_not_called()

    def test_preview_failure_does_not_claim_zero_or_complete_output(self):
        with patch.object(service.previews, "write_previews", side_effect=OSError("SECRET_PATH")), patch.object(service.alerts, "notify_owner") as alert:
            result = self.run_result(log_json=True)
        self.assertEqual((result.status, result.previews, result.no_movement), ("failed", None, 1))
        alert.assert_not_called()
        self.assertNotIn("SECRET_PATH", str(self.records()))

    def test_scheduled_result_logs_valid_computed_period(self):
        with patch.object(service.scheduling, "scheduled_month", return_value="2026-09"):
            result = service.run_result(self.root / "fake-config.json", card="black", scheduled=True, log_json=True)
        self.assertEqual((result.status, result.month), ("succeeded", "2026-09"))
        self.assertIsNone(self.records()[0]["month"])
        self.assertEqual(self.records()[1]["month"], "2026-09")


class LoggingSchemaTests(unittest.TestCase):
    def test_real_stderr_closed_pipe_preserves_process_exit_and_stdout(self):
        import os
        import subprocess
        import sys
        code = '''
from dataclasses import asdict
import io
import sys
from uuid import uuid4
from resumos_cartoes.operations import RunResult, emit_json
original = sys.stderr
assert type(original) is io.TextIOWrapper and original.fileno() == 2
result = RunResult(str(uuid4()), 'succeeded', 'send', 'black', '2026-09', 1, 0, 0, None, 0.1)
emit_json(sys.stderr, 'run_started', run_id=result.run_id, mode=result.mode, card=result.card, month=result.month)
emit_json(sys.stderr, 'run_finished', **asdict(result))
assert sys.stderr is original and not original.closed
print('human stdout unchanged')
'''
        # Close the sole reader before starting the child: no scheduling race,
        # and no replacement StringIO/wrapper can hide buffered shutdown errors.
        for ending, expected in (("", 0), ("sys.exit(result.exit_code)", 0), ("sys.exit(7)", 7)):
            with self.subTest(ending=ending):
                reader, writer = os.pipe()
                os.close(reader)
                try:
                    result = subprocess.run([sys.executable, "-I", "-c", code + ending],
                                            stdout=subprocess.PIPE, stderr=writer, text=True, timeout=30)
                finally:
                    os.close(writer)
                self.assertEqual(result.returncode, expected, "Optional logging poisoned interpreter shutdown")
                self.assertEqual(result.stdout, "human stdout unchanged\n")

    def test_native_and_memory_wrapped_sinks_keep_schema_privacy(self):
        import io
        import json
        from uuid import uuid4
        from resumos_cartoes.operations import emit_json
        safe = dict(run_id=str(uuid4()), mode="preview", card="black", month="2026-09")
        for factory in (lambda: tempfile.TemporaryFile(mode="w+", encoding="utf-8"),
                        lambda: io.TextIOWrapper(io.BytesIO(), encoding="utf-8")):
            with self.subTest(factory=factory), factory() as sink:
                emit_json(sink, "run_started", **(safe | {"token": "SYNTHETIC_PRIVATE"}))
                emit_json(sink, "run_started", **(safe | {"month": "SYNTHETIC_PRIVATE"}))
                emit_json(sink, "SYNTHETIC_PRIVATE", **safe)
                sink.seek(0)
                self.assertEqual(sink.read(), "")
                emit_json(sink, "run_started", **safe)
                sink.seek(0)
                self.assertEqual(json.loads(sink.read()), {"event": "run_started", **safe})

    def test_native_short_writes_complete_without_catching_interrupts(self):
        import json
        from uuid import uuid4
        from resumos_cartoes import operations
        safe = dict(run_id=str(uuid4()), mode="preview", card="black", month="2026-09")
        write = operations.os.write
        with tempfile.TemporaryFile(mode="w+", encoding="utf-8") as sink:
            with patch.object(operations.os, "write", side_effect=lambda fd, data: write(fd, data[:7])) as short:
                operations.emit_json(sink, "run_started", **safe)
            self.assertGreater(short.call_count, 1)
            sink.seek(0)
            self.assertEqual(json.loads(sink.read()), {"event": "run_started", **safe})
            with patch.object(operations.os, "write", side_effect=KeyboardInterrupt), self.assertRaises(KeyboardInterrupt):
                operations.emit_json(sink, "run_started", **safe)

    def test_imports_and_reloads_do_not_configure_logging(self):
        import subprocess
        import sys
        code = '''
import importlib
import logging
from unittest.mock import patch
root = logging.getLogger()
initial = (tuple(root.handlers), root.level)
with patch('logging.basicConfig', side_effect=AssertionError('Global config')), patch.object(logging.Logger, 'addHandler', side_effect=AssertionError('Global handler')):
    from resumos_cartoes import cli, service, operations
    for module in (operations, service, cli):
        importlib.reload(module)
assert (tuple(root.handlers), root.level) == initial
'''
        result = subprocess.run([sys.executable, "-I", "-c", code], capture_output=True, text=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr)
        self.assertEqual(result.stdout + result.stderr, "")

    def test_emitter_rejects_unknown_fields_and_non_allowlisted_values(self):
        from resumos_cartoes.operations import emit_json
        from uuid import uuid4
        safe = dict(run_id=str(uuid4()), mode="preview", card="black", month="2026-09")
        class HostileString(str):
            def __eq__(self, other):
                return True
        for update in ({"mode": HostileString("PRIVATE_SENTINEL")}, {"recipient": "PRIVATE_SENTINEL"}, {"mode": "PRIVATE_SENTINEL"},
                       {"card": "PRIVATE_SENTINEL"}, {"month": "PRIVATE_SENTINEL"},
                       {"month": "2026-13"}, {"month": "0000-01"},
                       {"run_id": "PRIVATE_SENTINEL"}, {"card": ["black"]}):
            with self.subTest(update=update):
                output = StringIO()
                emit_json(output, "run_started", **(safe | update))
                self.assertEqual(output.getvalue(), "", "Invalid fields must never reach the log sink")
        output = StringIO()
        emit_json(output, "PRIVATE_SENTINEL", **safe)
        emit_json(output, "run_started", run_id=safe["run_id"])
        self.assertEqual(output.getvalue(), "")

    def test_finished_schema_rejects_private_error_types_and_invalid_counts(self):
        from dataclasses import asdict
        from uuid import uuid4
        from resumos_cartoes.operations import RunResult, emit_json
        safe = asdict(RunResult(str(uuid4()), "failed", "send", "black", "2026-09", None, None, 0, None, 0.1, "operational_error", "failed"))
        for update in ({"error_category": "SECRET_CLASS"}, {"alert_status": "SECRET"}, {"status": "SECRET"},
                       {"sent": "SECRET"}, {"skipped": -1}, {"no_movement": True},
                       {"previews": {}}, {"elapsed_seconds": float("nan")}, {"elapsed_seconds": -1},
                       {"subject": "SECRET"}):
            with self.subTest(update=update):
                output = StringIO()
                emit_json(output, "run_finished", **(safe | update))
                self.assertEqual(output.getvalue(), "")
