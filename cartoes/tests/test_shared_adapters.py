"""Application adapter contracts with synthetic settings/services only."""
import ast
from pathlib import Path
import sqlite3
import tempfile
import unittest
from unittest.mock import MagicMock, patch

from automation_core.gmail import GmailClient
from resumos_cartoes import delivery, gmail, drive
from resumos_cartoes.config import Settings
from resumos_cartoes.errors import ReportError
from resumos_cartoes.ledger import Ledger
from resumos_cartoes.models import Recipient, Report


class SharedAdapterTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        root = Path(temp.name)
        self.settings = Settings(root, "drive", "Cartão", root / "output", root / "state", root / "fake.json", "Synthetic", (), "Synthetic Owner", personal_copy_email="copy@example.com")
        self.report = Report("black", "2026-09", Recipient("Pessoa01", "Participante01", "person@example.com"), (), (), "Synthetic Owner")

    def test_invalid_card_is_rejected_before_workbook_access(self):
        from contextlib import redirect_stderr
        from io import StringIO
        from resumos_cartoes import service
        with patch.object(service.config, "load_settings", return_value=self.settings), \
                patch.object(service, "input_workbook_bytes") as read, redirect_stderr(StringIO()):
            self.assertEqual(service.run(Path("synthetic.json"), card="unknown", requested_month="2026-09"), 1)
        read.assert_not_called()

    def test_app_adapters_delegate_to_shared_integrations(self):
        root = Path(__file__).resolve().parents[1] / "src/resumos_cartoes"
        for name in ("gmail", "drive"):
            tree = ast.parse((root / f"{name}.py").read_text())
            imports = [n.module for n in ast.walk(tree) if isinstance(n, ast.ImportFrom) and n.module]
            self.assertTrue(any(n.startswith("automation_core") for n in imports), name)
            self.assertFalse(any(n.startswith(("google.", "googleapiclient")) for n in imports), name)

    def test_missing_provider_acknowledgement_blocks_another_connection(self):
        service = MagicMock()
        service.users.return_value.getProfile.return_value.execute.return_value = {"emailAddress": "owner@example.com"}
        service.users.return_value.messages.return_value.send.return_value.execute.return_value = {}
        # Patch service construction, not the real adapter or send/ledger paths.
        with patch("automation_core.google_auth.build_service", return_value=service) as build:
            with self.assertRaisesRegex(ReportError, "incerto"):
                delivery.send_reports([self.report], self.settings, "synthetic-hash")
            with self.assertRaisesRegex(ReportError, "incerto"):
                delivery.send_reports([self.report], self.settings, "synthetic-hash")
        build.assert_called_once_with("gmail", "v1", self.settings.google_token_file)
        service.users.return_value.messages.return_value.send.assert_called_once()
        service._http.close.assert_called_once()
        with sqlite3.connect(self.settings.state_dir / "deliveries.sqlite3") as db:
            self.assertEqual(db.execute("SELECT status FROM deliveries").fetchall(), [("unknown",)])
        db.close()

    def test_local_invalid_message_keeps_existing_conservative_unknown_policy(self):
        from email.message import EmailMessage
        service = MagicMock()
        service.users.return_value.getProfile.return_value.execute.return_value = {"emailAddress": "owner@example.com"}
        invalid = EmailMessage()
        invalid["From"] = "owner@example.com"
        invalid["To"] = "not-a-mailbox"
        invalid.set_content("Synthetic")
        with patch("automation_core.google_auth.build_service", return_value=service), patch.object(delivery, "make_message", return_value=invalid):
            with self.assertRaisesRegex(ReportError, "incerto"):
                delivery.send_reports([self.report], self.settings, "hash")
        service.users.return_value.messages.return_value.send.assert_not_called()
        with sqlite3.connect(self.settings.state_dir / "deliveries.sqlite3") as db:
            self.assertEqual(db.execute("SELECT status FROM deliveries").fetchall(), [("unknown",)])
        db.close()
        with patch("automation_core.google_auth.build_service") as connect:
            with self.assertRaisesRegex(ReportError, "incerto"):
                delivery.send_reports([self.report], self.settings, "hash")
        connect.assert_not_called()

    def test_close_failure_does_not_hide_acceptance_or_skip_ledger_close(self):
        connection = MagicMock()
        connection.email_address = "owner@example.com"
        connection.close.side_effect = RuntimeError("PRIVATE_CLOSE")
        original_close = Ledger.close
        closed = []
        def close(ledger):
            closed.append(True)
            original_close(ledger)
        with patch.object(Ledger, "close", close):
            self.assertEqual(delivery.send_reports([self.report], self.settings, "hash", connector=lambda *_: connection), (1, 0))
        self.assertEqual(closed, [True])
        self.assertEqual(delivery.send_reports([self.report], self.settings, "hash", connector=lambda *_: self.fail("must not reconnect")), (0, 1))
        connection.send_message.assert_called_once()

    def test_reusable_send_failures_keep_existing_ledger_classification(self):
        from dataclasses import replace
        for status in (400, 403, 429, 500, 503, None):
            with self.subTest(status=status):
                settings = replace(self.settings, state_dir=self.settings.state_dir / str(status))
                service = MagicMock()
                service.users.return_value.getProfile.return_value.execute.return_value = {"emailAddress": "owner@example.com"}
                error = TimeoutError("PRIVATE_PROVIDER")
                if status is not None:
                    error.resp = type("Response", (), {"status": status})()
                request = service.users.return_value.messages.return_value.send.return_value
                request.execute.side_effect = error
                with patch("automation_core.google_auth.build_service", return_value=service) as build:
                    with self.assertRaises(ReportError) as raised:
                        delivery.send_reports([self.report], settings, "hash")
                self.assertNotIn("PRIVATE", str(raised.exception))
                expected = "failed" if status is not None and status < 500 else "unknown"
                with sqlite3.connect(settings.state_dir / "deliveries.sqlite3") as db:
                    self.assertEqual(db.execute("SELECT status FROM deliveries").fetchall(), [(expected,)])
                db.close()
                build.assert_called_once()
                request.execute.assert_called_once_with(num_retries=0)
                service._http.close.assert_called_once()

    def test_drive_adapter_preserves_missing_ambiguous_mime_and_no_fallback_policy(self):
        folder = {"id": "folder", "mimeType": "application/vnd.google-apps.folder"}
        year = {"id": "year", "mimeType": "application/vnd.google-apps.folder"}
        for pages, diagnostic in (
            ([[]], "Meu Drive"),
            ([[folder], []], "2026 não encontrada"),
            ([[folder], [year], []], "2026-09.xlsx não encontrado"),
            ([[folder, {"id": "duplicate"}]], "mais de um item chamado Cartão"),
            ([[folder], [year, {"id": "duplicate"}]], "mais de um item chamado Drive/Cartão/2026"),
            ([[folder], [year], [{"id": "one"}, {"id": "two"}]], "mais de um item chamado Drive/Cartão/2026/2026-09.xlsx"),
            ([[folder], [year], [{"id": "file", "mimeType": "text/plain"}]], "não é um Excel"),
        ):
            with self.subTest(diagnostic=diagnostic):
                service = MagicMock()
                service.files.return_value.get.return_value.execute.return_value = {"id": "root"}
                service.files.return_value.list.return_value.execute.side_effect = [{"files": page} for page in pages]
                with patch("automation_core.google_auth.build_service", return_value=service) as build, patch.object(Path, "read_bytes", side_effect=AssertionError("No local fallback")):
                    with self.assertRaisesRegex(ReportError, diagnostic):
                        drive.drive_workbook_bytes(self.settings, "2026-09")
                build.assert_called_once_with("drive", "v3", self.settings.google_token_file)
                self.assertEqual(service.files.return_value.list.call_count, len(pages))
                service.files.return_value.get_media.assert_not_called()
                service._http.close.assert_called_once()

    def test_profile_validation_and_auth_failures_are_app_diagnostics_before_claim(self):
        from automation_core.google_auth import GoogleAuthError
        with patch("automation_core.google_auth.build_service", side_effect=GoogleAuthError("Safe auth error")):
            with self.assertRaisesRegex(ReportError, "token OAuth"):
                delivery.send_reports([self.report], self.settings, "hash")
        ledger = Ledger(self.settings.state_dir / "deliveries.sqlite3")
        try:
            self.assertIsNone(ledger.get(self.report))
        finally:
            ledger.close()
        service = MagicMock()
        service.users.return_value.getProfile.return_value.execute.return_value = {"emailAddress": "not-a-mailbox"}
        service._http.close.side_effect = RuntimeError("PRIVATE_CLOSE")
        with patch("automation_core.google_auth.build_service", return_value=service):
            with self.assertRaisesRegex(ReportError, "endereço Gmail válido"):
                gmail.gmail_connect(self.settings.google_token_file)
        service._http.close.assert_called_once()


if __name__ == "__main__":
    unittest.main()
