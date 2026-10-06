from contextlib import redirect_stderr, redirect_stdout
from datetime import date, datetime, timezone
from dataclasses import replace
from decimal import Decimal
from email import policy
from email.parser import BytesParser
from io import StringIO
import base64
import json
from pathlib import Path
import re
import runpy
import sqlite3
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch
from zipfile import ZipFile

import openpyxl

import cartoes as app


class FakeGmail:
    def __init__(self, error_at=None, error=None):
        self.messages = []
        self.calls = 0
        self.error_at = error_at
        self.error = error
        self.closed = False
        self.email_address = "owner@gmail.com"

    def send_message(self, message):
        self.calls += 1
        if self.calls == self.error_at:
            raise self.error
        self.messages.append(message)
        return {}

    def close(self):
        self.closed = True


class ReportTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.config = self.root / "config.json"
        self.config.write_text(json.dumps({
            "owner_name": "Titular de teste",
            "sender_name": "Hermes de teste",
            "google_token_file": "auth/google_token.json",
            "input_dir": ".",
            "input_source": "local",
            "payment_footer": "Pix: (00) 00000-0000",
            "personal_copy_email": "copy@example.com",
            "recipients": [
                {"sheet": "Pessoa01", "participant": "Participante01", "email": "pessoa01@example.com"},
                {"sheet": "Pessoa02", "participant": "Participante02", "email": "pessoa02@example.com"},
            ]
        }))
        self.settings = app.load_settings(self.config)
        self.path = self.root / "2026-09.xlsx"
        self.make_workbook()

    def make_workbook(self):
        book = openpyxl.Workbook()
        book.remove(book.active)
        for name in ("Black", "Latam"):
            s = book.create_sheet(name)
            s["A2"], s["F2"], s["G2"] = "PARTICIPANTES", "Participante01", "Participante02"
            s["A4"] = "CARTÃO AZUL"
            for r in (6, 7):
                s.cell(r, 1, date(2025, 1, 3))
                s.cell(r, 2, "Loja <script> & teste")
                s.cell(r, 3, "Parcela 10 de 12" if r == 6 else None)
                s.cell(r, 4, 20)
                s.cell(r, 6, 10)
                s.cell(r, 7, 10)
            s["A8"] = "CARTÃO BLACK"
            s["A9"], s["B9"], s["D9"], s["F9"] = date(2026, 9, 1), "Estorno", -2.35, -2.35
            s["A10"], s["D10"], s["F10"], s["G10"] = "TOTAL", "=SUM(D6:D9)", "=SUM(F6:F9)", "=SUM(G6:G9)"
            s["A12"], s["F12"], s["G12"] = "PAGAMENTOS", -5.65, -20
            s["A13"], s["F13"] = "PAGAMENTOS", -2
            s["A15"], s["F15"], s["G15"] = "TOTAL GERAL", "=SUM(F10:F14)", "=SUM(G10:G14)"
        book.create_sheet("Pessoa01")
        s = book.create_sheet("Pessoa02")
        s["A6"] = "#VALUE!"  # Cached summaries must not be needed.
        book.save(self.path)
        book.close()

    def change(self, fn):
        b = openpyxl.load_workbook(self.path)
        fn(b)
        b.save(self.path)
        b.close()

    def reports(self):
        return app.read_reports(self.path, self.settings, "black", "2026-09")

    def test_shared_purchases_credits_payments_and_old_installments(self):
        before = self.path.read_bytes()
        person1, person2 = self.reports()
        self.assertEqual(len(person1.purchases), 3)
        self.assertEqual(person1.spending, Decimal("17.65"))
        self.assertEqual(person1.balance, Decimal("10.00"))
        self.assertEqual(person1.payments, (Decimal("-5.65"), Decimal("-2.00")))
        self.assertEqual(person1.purchases[0].date.year, 2025)
        self.assertEqual(person1.purchases[0].total, Decimal("20.00"))
        self.assertEqual(person1.purchases[1].installment, "—")
        self.assertTrue(person2.has_movement)
        self.assertEqual(person2.balance, Decimal("0.00"))
        self.assertEqual(before, self.path.read_bytes())

    def test_no_movement_is_not_an_error(self):
        self.change(lambda b: [setattr(b["Black"][f"G{r}"], "value", None) for r in (6, 7, 12)])
        self.assertFalse(self.reports()[1].has_movement)

    def test_payment_only_and_credit_balance(self):
        self.change(lambda b: [setattr(b["Black"][f"F{r}"], "value", None) for r in (6, 7, 9)])
        report = self.reports()[0]
        self.assertTrue(report.has_movement)
        self.assertEqual(report.balance, Decimal("-7.65"))
        self.assertIn("Crédito a seu favor", app.render_html(report))

    def test_no_payments_renders_zero_without_creating_movement(self):
        self.change(lambda b: [setattr(b["Black"][f"F{r}"], "value", None) for r in (12, 13)])
        report = self.reports()[0]
        self.assertEqual(report.payments, ())
        self.assertEqual(report.balance, report.spending)
        html = app.render_html(report)
        self.assertRegex(html, r"PAGAMENTOS</td><td[^>]*></td><td[^>]*>R\$ 0,00</td>")
        self.assertIn("PAGAMENTOS: R$ 0,00", app.render_text(report))
        self.change(lambda b: [setattr(b["Black"][f"F{r}"], "value", None) for r in (6, 7, 9)])
        self.assertFalse(self.reports()[0].has_movement)

    def test_numeric_rounding_and_bad_values(self):
        self.assertEqual(app._money(0.1 + 0.2, "F6"), Decimal("0.30"))
        self.assertEqual(app._money(1.005, "F6"), Decimal("1.01"))
        for value in ("10", "=SUM(A1:A3)", True, float("nan"), float("inf")):
            with self.subTest(value=str(value)), self.assertRaises(app.ReportError):
                app._money(value, "F6")

    def test_formula_in_source_is_rejected(self):
        self.change(lambda b: setattr(b["Black"]["F6"], "value", "=10"))
        with self.assertRaisesRegex(app.ReportError, "Black!F6"):
            self.reports()

    def test_invalid_date_and_purchase_total(self):
        self.change(lambda b: setattr(b["Black"]["A6"], "value", None))
        with self.assertRaisesRegex(app.ReportError, "data"):
            self.reports()
        self.make_workbook()
        self.change(lambda b: setattr(b["Black"]["D6"], "value", None))
        with self.assertRaisesRegex(app.ReportError, "Black!D6"):
            self.reports()

    def test_unknown_or_missing_sheets_and_headers_fail(self):
        self.change(lambda b: b.create_sheet("Novo amigo"))
        with self.assertRaisesRegex(app.ReportError, "sem cadastro"):
            self.reports()
        self.make_workbook()
        self.change(lambda b: b.remove(b["Pessoa02"]))
        with self.assertRaisesRegex(app.ReportError, "ausentes"):
            self.reports()
        self.make_workbook()
        self.change(lambda b: setattr(b["Black"]["G2"], "value", "Outro"))
        with self.assertRaisesRegex(app.ReportError, "Pessoa02"):
            self.reports()

    def test_duplicate_headers_and_bad_footer_fail(self):
        self.change(lambda b: setattr(b["Black"]["G2"], "value", "Participante01"))
        with self.assertRaisesRegex(app.ReportError, "duplicado"):
            self.reports()
        self.make_workbook()
        self.change(lambda b: setattr(b["Black"]["A15"], "value", "Saldo"))
        with self.assertRaisesRegex(app.ReportError, "TOTAL GERAL"):
            self.reports()

    def test_header_columns_can_move(self):
        self.change(lambda b: b["Black"].insert_cols(6))
        self.assertEqual(self.reports()[0].balance, Decimal("10.00"))

    def test_html_escapes_content_and_keeps_five_columns(self):
        report = self.reports()[0]
        html = app.render_html(report)
        self.assertNotIn("<script>", html)
        self.assertIn("&lt;script&gt; &amp;", html)
        self.assertIn(app.BLUE, html)
        for label in ("Data", "Lançamento", "Parcelas", "Total", "Rateio", "TOTAL GERAL"):
            self.assertIn(label, html)
        self.assertIn("03/01/2025", html)
        self.assertIn("R$ 10,00", html)
        self.assertNotIn("2026-09.xlsx", html)
        self.assertNotIn("pessoa02@example.com", html)

    def test_multipart_message_is_individual(self):
        report = self.reports()[0]
        msg = app.make_message(report, "owner@gmail.com", self.settings.sender_name, "pessoa01@example.com", cc=self.settings.personal_copy_email)
        self.assertEqual(msg["From"], "Hermes de teste <owner@gmail.com>")
        self.assertEqual(msg["Cc"], "copy@example.com")
        self.assertIsNone(msg["Bcc"])
        self.assertEqual(msg.get_body(preferencelist=("plain",)).get_content_type(), "text/plain")
        self.assertEqual(msg.get_body(preferencelist=("html",)).get_content_type(), "text/html")
        plain_part = msg.get_body(preferencelist=("plain",))
        html_part = msg.get_body(preferencelist=("html",))
        assert plain_part is not None and html_part is not None
        plain = plain_part.get_content()
        html = html_part.get_content()
        for content in (plain, html):
            self.assertIn("Olá, Pessoa01! Tudo bem?", content)
            self.assertIn("Sou o Hermes, assistente pessoal do Titular de teste.", content)
            self.assertIn("Segue abaixo o seu resumo, referente a 09/2026", content)
            self.assertIn("fale diretamente com Titular de teste.", content)
            self.assertIn("Pix: (00) 00000-0000", content)
            self.assertNotIn("O valor individual está na coluna Rateio.", content)
            self.assertNotIn("Saldo restante a pagar.", content)
        self.assertTrue(plain.rstrip().endswith("Pix: (00) 00000-0000"))

    def test_missing_owner_name_allows_preview_but_blocks_send(self):
        from dataclasses import replace
        report = replace(self.reports()[0], owner_name=None)
        self.assertIn("assistente pessoal do [seu nome]", app.render_html(report))
        with self.assertRaisesRegex(app.ReportError, "owner_name"):
            app.send_reports([report], self.settings, "hash", connector=lambda *_: self.fail("must not connect"))
        self.assertFalse(self.settings.state_dir.exists())

    def test_config_defaults_to_drive_cartoes_and_year_folder(self):
        data = json.loads(self.config.read_text())
        data.pop("input_source")
        data.pop("drive_folder_name", None)
        self.config.write_text(json.dumps(data))
        settings = app.load_settings(self.config)
        self.assertEqual(settings.input_source, "drive")
        self.assertEqual(settings.drive_folder_name, "Cartão")

    def test_drive_workbook_resolves_root_year_and_exact_month_file(self):
        data = json.loads(self.config.read_text())
        data["input_source"] = "drive"
        data["drive_folder_name"] = "Cartão"
        self.config.write_text(json.dumps(data))
        settings = app.load_settings(self.config)
        service = MagicMock()
        service.files.return_value.get.return_value.execute.return_value = {"id": "my-drive-root"}
        service.files.return_value.list.return_value.execute.side_effect = [
            {"files": [{"id": "cards-folder", "name": "Cartão", "mimeType": "application/vnd.google-apps.folder"}]},
            {"files": [{"id": "year-folder", "name": "2026", "mimeType": "application/vnd.google-apps.folder"}]},
            {"files": [{"id": "month-file", "name": "2026-09.xlsx", "mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}]},
        ]
        with patch.object(app, "_build_drive_service", return_value=service), patch.object(app, "_download_drive_file", return_value=b"xlsx-bytes") as download:
            content = app.drive_workbook_bytes(settings, "2026-09")
        self.assertEqual(content, b"xlsx-bytes")
        download.assert_called_once_with(service, "month-file")
        queries = [call.kwargs["q"] for call in service.files.return_value.list.call_args_list]
        self.assertIn("'my-drive-root' in parents", queries[0])
        self.assertIn("name = 'Cartão'", queries[0])
        self.assertIn("'cards-folder' in parents", queries[1])
        self.assertIn("name = '2026'", queries[1])
        self.assertIn("'year-folder' in parents", queries[2])
        self.assertIn("name = '2026-09.xlsx'", queries[2])

    def test_download_drive_file_uses_media_endpoint_without_retry(self):
        retry_counts = []

        class FakeDownloader:
            def __init__(self, target, request):
                self.target = target
                self.done = False

            def next_chunk(self, num_retries):
                retry_counts.append(num_retries)
                if not self.done:
                    self.target.write(b"xlsx-bytes")
                    self.done = True
                return None, self.done

        service = MagicMock()
        with patch("googleapiclient.http.MediaIoBaseDownload", FakeDownloader):
            content = app._download_drive_file(service, "excel-id")
        self.assertEqual(content, b"xlsx-bytes")
        self.assertEqual(retry_counts, [0])
        service.files.return_value.get_media.assert_called_once_with(fileId="excel-id")

    def test_drive_workbook_reports_missing_month_file(self):
        data = json.loads(self.config.read_text())
        data["input_source"] = "drive"
        self.config.write_text(json.dumps(data))
        settings = app.load_settings(self.config)
        service = MagicMock()
        service.files.return_value.get.return_value.execute.return_value = {"id": "my-drive-root"}
        service.files.return_value.list.return_value.execute.side_effect = [
            {"files": [{"id": "cards-folder", "name": "Cartão", "mimeType": "application/vnd.google-apps.folder"}]},
            {"files": [{"id": "year-folder", "name": "2026", "mimeType": "application/vnd.google-apps.folder"}]},
            {"files": []},
        ]
        with patch.object(app, "_build_drive_service", return_value=service):
            with self.assertRaisesRegex(app.ReportError, "2026-09.xlsx"):
                app.drive_workbook_bytes(settings, "2026-09")

    def test_cli_help_explains_scheduled_competencies(self):
        output = StringIO()
        with redirect_stdout(output), self.assertRaises(SystemExit):
            app.main(["--help"])
        help_text = " ".join(output.getvalue().split())
        self.assertIn("Black dia 5 = mês anterior", help_text)
        self.assertIn("Latam dia 20 = mês atual", help_text)

    def test_config_example_uses_nonpersonal_placeholders(self):
        example_path = Path(app.__file__).with_name("config.example.json")
        data = json.loads(example_path.read_text(encoding="utf-8"))
        self.assertIsNone(data["owner_name"])
        self.assertIsNone(data["payment_footer"])
        self.assertIsNone(data["personal_copy_email"])
        self.assertEqual(data["sender_name"], "Hermes")
        recipients = data["recipients"]
        self.assertEqual(len(recipients), 12)
        for index, recipient in enumerate(recipients, 1):
            number = f"{index:02d}"
            self.assertEqual(recipient["sheet"], f"Pessoa{number}")
            self.assertEqual(recipient["participant"], f"Participante{number}")
            self.assertIsNone(recipient["email"])

    def test_default_sender_name_does_not_expose_owner_identity(self):
        data = json.loads(self.config.read_text())
        data.pop("sender_name", None)
        self.config.write_text(json.dumps(data))
        settings = app.load_settings(self.config)
        self.assertEqual(settings.sender_name, "Hermes")

    def test_config_validation_and_relative_paths(self):
        self.assertEqual(self.settings.input_dir, self.root.resolve())
        self.assertEqual(self.settings.google_token_file, (self.root / "auth/google_token.json").resolve())
        self.assertEqual(self.settings.sender_name, "Hermes de teste")
        self.assertEqual(self.settings.payment_footer, "Pix: (00) 00000-0000")
        self.assertEqual(self.settings.personal_copy_email, "copy@example.com")
        data = json.loads(self.config.read_text())
        data["recipients"][1]["sheet"] = "Pessoa01"
        self.config.write_text(json.dumps(data))
        with self.assertRaisesRegex(app.ReportError, "duplicado"):
            app.load_settings(self.config)
        self.assertFalse(app.valid_email("a@example.com\nBcc: b@example.com"))
        self.assertFalse(app.valid_email("a@example.com,b@example.com"))
        for email in ("a:b@example.com", "a@domain..com", ".a@example.com", "a@example.com\n", "a@-domain.com"):
            self.assertFalse(app.valid_email(email), email)
        self.assertTrue(app.valid_email("a.b+cartoes@example.com"))

    def test_scheduled_month_uses_brasilia_and_year_boundary(self):
        self.assertEqual(app.scheduled_month("black", datetime(2027, 1, 5, 12, tzinfo=timezone.utc)), "2026-12")
        self.assertEqual(app.scheduled_month("latam", datetime(2026, 10, 20, 12, tzinfo=timezone.utc)), "2026-10")
        self.assertEqual(app.scheduled_month("latam", datetime(2027, 1, 20, 12, tzinfo=timezone.utc)), "2027-01")
        with self.assertRaises(app.ReportError):
            app.scheduled_month("black", datetime(2026, 10, 5, 1, tzinfo=timezone.utc))
        for month in ("2026-13", "2026-9", "../2026-09", "0000-01"):
            with self.assertRaises(app.ReportError):
                app.validate_month(month)

    def test_dry_run_uses_drive_source_when_configured(self):
        data = json.loads(self.config.read_text())
        data["input_source"] = "drive"
        data["drive_folder_name"] = "Cartão"
        self.config.write_text(json.dumps(data))
        with patch.object(app, "drive_workbook_bytes", return_value=self.path.read_bytes()) as fetch, redirect_stdout(StringIO()):
            code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09"])
        self.assertEqual(code, 0)
        fetch.assert_called_once_with(app.load_settings(self.config), "2026-09")
        self.assertTrue((self.root / "outputs/2026-09/black/index.html").exists())

    def test_dry_run_does_not_connect_or_write_ledger(self):
        with patch.object(app, "gmail_connect", side_effect=AssertionError("must not connect")), redirect_stdout(StringIO()):
            code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09"])
        self.assertEqual(code, 0)
        self.assertEqual(len(list((self.root / "outputs/2026-09/black").glob("*.html"))), 3)
        self.assertFalse(self.settings.state_dir.exists())

    def test_preview_refresh_removes_only_previous_generated_files(self):
        folder = self.root / "preview"
        reports = self.reports()
        app.write_previews(reports, folder)
        custom = folder / "minha-nota.txt"
        custom.write_text("preservar")
        app.write_previews([reports[0]], folder)
        self.assertEqual(len(list(folder.glob("*.html"))), 2)  # Pessoa01 + index
        self.assertTrue(custom.exists())
        self.assertNotIn("Pessoa02", (folder / "index.html").read_text())

    def test_send_requires_personal_copy_email_before_connecting(self):
        settings = replace(self.settings, personal_copy_email=None)
        connector = MagicMock()
        with self.assertRaisesRegex(app.ReportError, "personal_copy_email"):
            app.send_reports(self.reports(), settings, "hash", connector=connector)
        connector.assert_not_called()
        self.assertFalse(settings.state_dir.exists())

    def test_send_reports_uses_hermes_oauth_profile_without_app_password(self):
        reports = [self.reports()[0]]
        connection = MagicMock()
        connection.email_address = "owner@gmail.com"
        with patch.object(app, "gmail_connect", return_value=connection) as connect:
            sent, skipped = app.send_reports(reports, self.settings, "hash")
        connect.assert_called_once_with(self.settings.google_token_file)
        self.assertEqual((sent, skipped), (1, 0))
        message = connection.send_message.call_args.args[0]
        self.assertEqual(message["From"], "Hermes de teste <owner@gmail.com>")
        self.assertEqual(message["Subject"], app.subject(reports[0]))
        self.assertEqual(message["To"], "pessoa01@example.com")
        self.assertEqual(message["Cc"], "copy@example.com")
        connection.close.assert_called_once()

    def test_gmail_connect_loads_configured_hermes_oauth_token(self):
        token_file = self.root / "google_token.json"
        service = unittest.mock.MagicMock()
        service.users.return_value.getProfile.return_value.execute.return_value = {
            "emailAddress": "owner@gmail.com"
        }
        with patch.object(app, "_build_gmail_service", return_value=service) as build_service:
            connection = app.gmail_connect(token_file)
        build_service.assert_called_once_with(token_file)
        self.assertEqual(connection.email_address, "owner@gmail.com")

    def test_gmail_api_client_uses_authenticated_profile_and_sends_raw_mime(self):
        service = unittest.mock.MagicMock()
        service.users.return_value.getProfile.return_value.execute.return_value = {
            "emailAddress": "owner@gmail.com"
        }
        service.users.return_value.messages.return_value.send.return_value.execute.return_value = {
            "id": "gmail-internal-id"
        }
        client = app.GmailAPIConnection(service)
        self.assertEqual(client.email_address, "owner@gmail.com")

        message = app.make_message(self.reports()[0], client.email_address, self.settings.sender_name, "pessoa01@example.com", cc=self.settings.personal_copy_email)
        result = client.send_message(message)

        self.assertEqual(result["id"], "gmail-internal-id")
        service.users.return_value.messages.return_value.send.assert_called_once()
        args, kwargs = service.users.return_value.messages.return_value.send.call_args
        self.assertEqual(kwargs["userId"], "me")
        raw = kwargs["body"]["raw"]
        decoded = base64.urlsafe_b64decode(raw)
        parsed = BytesParser(policy=policy.default).parsebytes(decoded)
        self.assertEqual(parsed["To"], "pessoa01@example.com")
        self.assertEqual(parsed["Cc"], "copy@example.com")
        self.assertEqual(parsed["From"], "Hermes de teste <owner@gmail.com>")
        self.assertEqual(parsed["Message-ID"], message["Message-ID"])
        service.users.return_value.messages.return_value.send.return_value.execute.assert_called_once_with(num_retries=0)

    def test_gmail_connection_closes_transport_when_profile_lookup_fails(self):
        service = MagicMock()
        service.users.return_value.getProfile.return_value.execute.side_effect = RuntimeError("offline")
        with self.assertRaisesRegex(RuntimeError, "offline"):
            app.GmailAPIConnection(service)
        service._http.close.assert_called_once()

    def test_gmail_api_connection_marks_http_4xx_as_definitive_rejection(self):
        service = MagicMock()
        service.users.return_value.getProfile.return_value.execute.return_value = {
            "emailAddress": "owner@gmail.com"
        }
        error = RuntimeError("private API response")
        error.resp = type("Response", (), {"status": 403})()
        service.users.return_value.messages.return_value.send.return_value.execute.side_effect = error
        connection = app.GmailAPIConnection(service)
        with self.assertRaises(app.GmailSendRejected):
            connection.send_message(app.make_message(self.reports()[0], "owner@gmail.com", "Cartões", "pessoa01@example.com"))

    def test_hermes_launchers_use_config_and_scheduled_production_mode(self):
        root = Path(__file__).resolve().parents[1]
        for card in ("black", "latam"):
            default = Path("~/Projetos/Automations/automations/cartoes/config.json").expanduser()
            for configured, expected in ((None, default), (str(self.config), self.config)):
                with self.subTest(card=card, configured=configured), patch.dict(app.os.environ), patch.object(app, "main", return_value=0) as main:
                    app.os.environ.pop("CARTOES_CONFIG", None)
                    if configured is not None:
                        app.os.environ["CARTOES_CONFIG"] = configured
                    with self.assertRaises(SystemExit) as exited:
                        runpy.run_path(str(root / f"hermes/cartao_{card}.py"), run_name="__main__")
                    self.assertEqual(exited.exception.code, 0)
                    main.assert_called_once_with(["--config", str(expected), "--card", card, "--scheduled", "--send"])

    def test_missing_file_sends_owner_alert_with_copy_without_creating_ledger(self):
        self.path.unlink()
        gmail = FakeGmail()
        with patch.object(app, "gmail_connect", return_value=gmail) as connect, redirect_stderr(StringIO()):
            self.assertEqual(app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", "--send"]), 1)
        connect.assert_called_once()
        self.assertEqual([m["To"] for m in gmail.messages], ["owner@gmail.com"])
        self.assertEqual(gmail.messages[0]["From"], "Hermes de teste <owner@gmail.com>")
        self.assertEqual(gmail.messages[0]["Cc"], "copy@example.com")
        self.assertIn("[ERRO]", gmail.messages[0]["Subject"])
        self.assertIn("Nenhum resumo foi enviado", gmail.messages[0].get_content())
        self.assertIn("Este aviso é destinado ao proprietário.", gmail.messages[0].get_content())
        self.assertNotIn("somente para você", gmail.messages[0].get_content())
        self.assertFalse(self.settings.state_dir.exists())

    def test_corrupt_excel_at_open_or_lazy_read_alerts_owner_and_copy(self):
        for failure in ("archive", "cell", "xml"):
            with self.subTest(failure=failure):
                self.make_workbook()
                if failure == "archive":
                    self.path.write_bytes(b"not an Excel archive")
                else:
                    with ZipFile(self.path) as source:
                        files = [(name, source.read(name)) for name in source.namelist()]
                    with ZipFile(self.path, "w") as target:
                        for name, data in files:
                            if name == "xl/worksheets/sheet1.xml":
                                if failure == "cell":
                                    data, count = re.subn(rb'(<c r="F6"[^>]*><v>)10(</v>)', rb'\g<1>PRIVATE_CELL_CONTENT\g<2>', data)
                                    self.assertEqual(count, 1)
                                else:
                                    data = data.replace(b"</sheetData>", b"</bad-sheetData>")
                            target.writestr(name, data)
                gmail, stderr = FakeGmail(), StringIO()
                with patch.object(app, "gmail_connect", return_value=gmail) as connect, redirect_stderr(stderr):
                    code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", "--send"])
                self.assertEqual(code, 1)
                connect.assert_called_once()
                self.assertEqual(len(gmail.messages), 1)
                message = gmail.messages[0]
                self.assertEqual(message["To"], "owner@gmail.com")
                self.assertEqual(message["Cc"], "copy@example.com")
                self.assertIsNone(message["Bcc"])
                self.assertEqual(list(message.iter_attachments()), [])
                self.assertIn("Excel", message.get_content())
                self.assertIn("Nenhum resumo foi enviado", message.get_content())
                for text in (stderr.getvalue(), message.as_string()):
                    self.assertNotIn("PRIVATE_CELL_CONTENT", text)
                    self.assertNotIn("Traceback", text)
                self.assertFalse(self.settings.state_dir.exists())
                self.assertTrue(gmail.closed)

    def test_invalid_values_or_recipients_alert_owner_with_copy(self):
        for failure in ("value", "recipient"):
            with self.subTest(failure=failure):
                if failure == "value":
                    self.change(lambda b: setattr(b["Black"]["F6"], "value", "=10"))
                else:
                    self.make_workbook()
                    data = json.loads(self.config.read_text())
                    data["recipients"][1]["email"] = None
                    self.config.write_text(json.dumps(data))
                gmail = FakeGmail()
                with patch.object(app, "gmail_connect", return_value=gmail), redirect_stderr(StringIO()):
                    code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", "--send"])
                self.assertEqual(code, 1)
                self.assertEqual([m["To"] for m in gmail.messages], ["owner@gmail.com"])
                self.assertFalse((self.settings.state_dir / "deliveries.sqlite3").exists())

    def test_invalid_config_uses_hermes_oauth_for_owner_alert(self):
        self.config.write_text("{invalid JSON")
        gmail = FakeGmail()
        with patch.dict(app.os.environ, {}, clear=True), patch.object(app, "gmail_connect", return_value=gmail) as connect, redirect_stderr(StringIO()):
            code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", "--send"])
        self.assertEqual(code, 1)
        connect.assert_called_once_with(Path.home() / ".hermes/google_token.json")
        self.assertEqual([m["To"] for m in gmail.messages], ["owner@gmail.com"])
        self.assertIn("configuração", gmail.messages[0].get_content())

    def test_dry_run_errors_never_connect_to_gmail(self):
        self.path.write_bytes(b"not an Excel archive")
        for mode in ([], ["--dry-run"]):
            with self.subTest(mode=mode), patch.object(app, "gmail_connect") as connect, redirect_stderr(StringIO()):
                code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", *mode])
                self.assertEqual(code, 1)
                connect.assert_not_called()

    def test_alert_failure_does_not_retry_or_hide_original_error(self):
        self.path.unlink()
        for error in (RuntimeError("SECRET_SERVER_RESPONSE"), TimeoutError("SECRET_SERVER_RESPONSE")):
            with self.subTest(error=type(error).__name__):
                stderr = StringIO()
                with patch.object(app, "gmail_connect", side_effect=error) as connect, redirect_stderr(stderr):
                    code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", "--send"])
                self.assertEqual(code, 1)
                connect.assert_called_once()
                self.assertIn("ausente", stderr.getvalue())
                self.assertIn("Não foi possível confirmar o aviso", stderr.getvalue())
                self.assertNotIn("SECRET_SERVER_RESPONSE", stderr.getvalue())

    def test_alert_is_not_redirected_to_test_recipient(self):
        self.path.unlink()
        gmail = FakeGmail()
        with patch.object(app, "gmail_connect", return_value=gmail), redirect_stderr(StringIO()):
            code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", "--send", "--test-to", "test@example.com"])
        self.assertEqual(code, 1)
        self.assertEqual([m["To"] for m in gmail.messages], ["owner@gmail.com"])

    def test_alert_rejection_does_not_retry(self):
        self.path.unlink()
        gmail = FakeGmail(1, app.GmailSendRejected("rejected"))
        stderr = StringIO()
        with patch.object(app, "gmail_connect", return_value=gmail) as connect, redirect_stderr(stderr):
            code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", "--send"])
        self.assertEqual(code, 1)
        connect.assert_called_once()
        self.assertEqual(gmail.calls, 1)
        self.assertTrue(gmail.closed)
        self.assertIn("Não foi possível confirmar o aviso", stderr.getvalue())

    def test_missing_oauth_token_reports_setup_requirement(self):
        with patch.object(app, "_build_gmail_service", side_effect=FileNotFoundError):
            with self.assertRaisesRegex(app.ReportError, "token OAuth"):
                app.gmail_connect(self.settings.google_token_file)

    def test_unexpected_failure_alert_does_not_expose_exception_payload(self):
        gmail, stderr = FakeGmail(), StringIO()
        with patch.object(app, "read_reports", side_effect=RuntimeError("PRIVATE_PAYLOAD")), patch.object(app, "gmail_connect", return_value=gmail), redirect_stderr(stderr):
            code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", "--send"])
        self.assertEqual(code, 1)
        self.assertEqual([m["To"] for m in gmail.messages], ["owner@gmail.com"])
        self.assertIn("RuntimeError", gmail.messages[0].get_content())
        self.assertNotIn("PRIVATE_PAYLOAD", stderr.getvalue() + gmail.messages[0].get_content())

    def test_partial_send_failure_sends_owner_alert_and_keeps_history(self):
        delivery = FakeGmail(2, TimeoutError("PRIVATE_SERVER_RESPONSE"))
        alert = FakeGmail()
        with patch.object(app, "gmail_connect", side_effect=[delivery, alert]), redirect_stderr(StringIO()):
            code = app.main(["--config", str(self.config), "--card", "black", "--month", "2026-09", "--send"])
        self.assertEqual(code, 1)
        self.assertEqual([m["To"] for m in delivery.messages], ["pessoa01@example.com"])
        self.assertEqual([m["To"] for m in alert.messages], ["owner@gmail.com"])
        self.assertIn("Alguns resumos podem já ter sido aceitos", alert.messages[0].get_content())
        self.assertNotIn("Nenhum resumo foi enviado", alert.messages[0].get_content())
        self.assertNotIn("PRIVATE_SERVER_RESPONSE", alert.messages[0].get_content())
        with sqlite3.connect(self.settings.state_dir / "deliveries.sqlite3") as db:
            self.assertEqual(db.execute("SELECT sheet, status FROM deliveries ORDER BY sheet").fetchall(), [("Pessoa01", "sent"), ("Pessoa02", "unknown")])

    def test_history_shows_original_address_and_update_time(self):
        app.send_reports([self.reports()[0]], self.settings, "hash", connector=lambda *_: FakeGmail())
        data = json.loads(self.config.read_text())
        data["recipients"][0]["email"] = "changed@example.com"
        self.config.write_text(json.dumps(data))
        stdout = StringIO()
        with redirect_stdout(stdout):
            code = app.main(["--config", str(self.config), "--history", "--month", "2026-09"])
        self.assertEqual(code, 0)
        self.assertIn("destinatário | atualização (UTC) | Message-ID", stdout.getvalue())
        self.assertIn("pessoa01@example.com", stdout.getvalue())
        self.assertNotIn("changed@example.com", stdout.getvalue())
        self.assertRegex(stdout.getvalue(), r"\d{4}-\d{2}-\d{2}T.*\+00:00")

    def test_partial_failure_and_retry_skip_success(self):
        reports = self.reports()
        gmail = FakeGmail(error_at=2, error=app.GmailSendRejected("Rejected"))
        with self.assertRaises(app.ReportError):
            app.send_reports(reports, self.settings, "hash", connector=lambda *_: gmail)
        self.assertTrue(gmail.closed)
        second = FakeGmail()
        self.assertEqual(app.send_reports(reports, self.settings, "newhash", connector=lambda *_: second), (1, 1))
        self.assertEqual(second.messages[0]["To"], "pessoa02@example.com")
        self.assertEqual(app.send_reports(reports, self.settings, "newhash", connector=lambda *_: self.fail("duplicate")), (0, 2))

    def test_timeout_is_unknown_and_not_retried(self):
        reports = self.reports()
        with self.assertRaisesRegex(app.ReportError, "incerto"):
            app.send_reports(reports, self.settings, "hash", connector=lambda *_: FakeGmail(1, TimeoutError()))
        with self.assertRaisesRegex(app.ReportError, "incerto"):
            app.send_reports(reports, self.settings, "hash", connector=lambda *_: self.fail("must not connect"))
        with redirect_stdout(StringIO()):
            app.maintain_ledger(self.settings, "2026-09", "black", "Pessoa01", "sent")
        gmail = FakeGmail()
        self.assertEqual(app.send_reports(reports, self.settings, "hash", connector=lambda *_: gmail), (1, 1))

    def test_interruption_is_recovered_as_unknown(self):
        report = self.reports()[0]
        with app.run_lock(self.settings.state_dir):
            ledger = app.Ledger(self.settings.state_dir / "deliveries.sqlite3")
            ledger.claim(report, "<attempt@example.com>", "hash")
            ledger.close()
        with self.assertRaisesRegex(app.ReportError, "incerto"):
            app.send_reports([report], self.settings, "hash", connector=lambda *_: self.fail("must not connect"))

    def test_database_failure_after_gmail_acceptance_blocks_retry(self):
        report, gmail = self.reports()[0], FakeGmail()
        with patch.object(app.Ledger, "finish", side_effect=sqlite3.OperationalError("disk error")):
            with self.assertRaises(sqlite3.OperationalError):
                app.send_reports([report], self.settings, "hash", connector=lambda *_: gmail)
        self.assertEqual(len(gmail.messages), 1)
        with self.assertRaisesRegex(app.ReportError, "incerto"):
            app.send_reports([report], self.settings, "hash", connector=lambda *_: self.fail("must not reconnect"))

    def test_test_delivery_redirects_without_production_history(self):
        gmail = FakeGmail()
        reports = self.reports()
        app.send_reports(reports, self.settings, "hash", test_to="owner@gmail.com", connector=lambda *_: gmail)
        self.assertEqual(len(gmail.messages), 2)
        for report, msg in zip(reports, gmail.messages, strict=True):
            self.assertEqual(msg["To"], "owner@gmail.com")
            self.assertEqual(msg["Cc"], "copy@example.com")
            self.assertEqual(msg["Subject"], "[TESTE] " + app.subject(report))
        self.assertFalse((self.settings.state_dir / "deliveries.sqlite3").exists())

    def test_missing_recipient_blocks_whole_batch_before_gmail(self):
        report = self.reports()[1]
        bad = app.Report(report.card, report.month, app.Recipient("Pessoa02", "Participante02", None), report.purchases, report.payments)
        with self.assertRaisesRegex(app.ReportError, "Pessoa02"):
            app.send_reports([self.reports()[0], bad], self.settings, "hash", connector=lambda *_: self.fail("must not connect"))
        self.assertFalse(self.settings.state_dir.exists())

    def test_lock_blocks_parallel_execution(self):
        with app.run_lock(self.settings.state_dir):
            with self.assertRaisesRegex(app.ReportError, "andamento"):
                with app.run_lock(self.settings.state_dir):
                    self.fail("second sender obtained lock")

    def test_lock_blocks_a_separate_process(self):
        script = """import sys
from pathlib import Path
import cartoes
try:
    with cartoes.run_lock(Path(sys.argv[1])):
        sys.exit(2)
except cartoes.ReportError:
    sys.exit(0)
"""
        with app.run_lock(self.settings.state_dir):
            result = subprocess.run([sys.executable, "-c", script, str(self.settings.state_dir)],
                                    cwd=Path(app.__file__).resolve().parent, capture_output=True, timeout=10)
        self.assertEqual(result.returncode, 0, result.stderr.decode())

    def test_oauth_auth_error_occurs_before_claim(self):
        def fail_connect(*_):
            raise app.ReportError("OAuth authentication failed")

        with self.assertRaises(app.ReportError):
            app.send_reports(self.reports(), self.settings, "hash", connector=fail_connect)
        with app.run_lock(self.settings.state_dir):
            ledger = app.Ledger(self.settings.state_dir / "deliveries.sqlite3")
            self.assertIsNone(ledger.get(self.reports()[0]))
            ledger.close()

    def test_hermes_home_env_controls_default_oauth_token_path(self):
        data = json.loads(self.config.read_text())
        del data["google_token_file"]
        self.config.write_text(json.dumps(data))
        hermes_home = self.root / "profile-hermes"
        with patch.dict(app.os.environ, {"HERMES_HOME": str(hermes_home)}):
            settings = app.load_settings(self.config)
        self.assertEqual(settings.google_token_file, hermes_home / "google_token.json")


class AnonymousWorkbookTests(unittest.TestCase):
    def test_24_fictitious_blocks_without_private_workbook(self):
        fixture = json.loads((Path(__file__).parent / "fixtures/anonymous_reports.json").read_text(encoding="utf-8"))
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            config = root / "config.json"
            config.write_text(json.dumps({"owner_name": "Proprietário fictício", "recipients": [
                {"sheet": p["sheet"], "participant": p["participant"], "email": None} for p in fixture["people"]
            ]}), encoding="utf-8")
            settings = app.load_settings(config)
            workbook = openpyxl.Workbook()
            workbook.remove(workbook.active)
            expected_purchases = {}
            for card, title in (("black", "Black"), ("latam", "Latam")):
                sheet = workbook.create_sheet(title)
                sheet["A2"] = "PARTICIPANTES"
                row = 4
                for column, person in enumerate(fixture["people"], 6):
                    sheet.cell(2, column, person["participant"])
                    expected_purchases[card, person["sheet"]] = []
                    sheet.cell(row, 1, "SEÇÃO DE COMPRAS")
                    row += 1
                    for value in person[card]["purchases"]:
                        amount = Decimal(value)
                        purchase = app.Purchase(date(2025, 1, 3), ("Descrição fictícia extensa com <símbolos> & detalhes " * 5).strip(),
                                                "2 de 12", amount * 2, amount)
                        sheet.cell(row, 1, purchase.date)
                        sheet.cell(row, 2, purchase.description)
                        sheet.cell(row, 3, purchase.installment)
                        sheet.cell(row, 4, float(purchase.total))
                        sheet.cell(row, column, float(amount))
                        expected_purchases[card, person["sheet"]].append(purchase)
                        row += 1
                sheet.cell(row, 1, "TOTAL")
                row += 1
                for column, person in enumerate(fixture["people"], 6):
                    for value in person[card]["payments"]:
                        sheet.cell(row, 1, "PAGAMENTOS")
                        sheet.cell(row, column, float(Decimal(value)))
                        row += 1
                sheet.cell(row, 1, "TOTAL GERAL")
            for person in fixture["people"]:
                workbook.create_sheet(person["sheet"])
            path = root / f"{fixture['month']}.xlsx"
            workbook.save(path)
            workbook.close()
            for card in ("black", "latam"):
                reports = app.read_reports(path, settings, card, fixture["month"])
                self.assertEqual(len(reports), 12)
                self.assertEqual(sum(r.has_movement for r in reports), fixture["expected_messages"][card])
                for report, person in zip(reports, fixture["people"], strict=True):
                    with self.subTest(card=card, sheet=person["sheet"]):
                        expected = person[card]
                        self.assertEqual(report.purchases, tuple(expected_purchases[card, person["sheet"]]))
                        self.assertEqual(report.payments, tuple(Decimal(v) for v in expected["payments"]))
                        self.assertEqual(report.spending, Decimal(expected["expected_spending"]))
                        self.assertEqual(report.balance, Decimal(expected["expected_balance"]))


class ProvidedWorkbookTests(unittest.TestCase):
    """Optional private fixture: never committed, included in local validation."""
    root = Path(__file__).resolve().parents[1]

    @unittest.skipUnless((root / "inputs/2026-09.xlsx").exists(), "Exemplo pessoal não está no checkout")
    def test_all_24_blocks_match_independent_cached_personal_summaries(self):
        settings = app.load_settings(self.root / "config.example.json")
        wb = openpyxl.load_workbook(self.root / "inputs/2026-09.xlsx", data_only=True)
        self.addCleanup(wb.close)
        for card, start, share_column, expected_count in (("black", 1, 5, 11), ("latam", 7, 11, 3)):
            reports = app.read_reports(self.root / "inputs/2026-09.xlsx", settings, card, "2026-09")
            self.assertEqual(len(reports), 12)
            self.assertEqual(sum(r.has_movement for r in reports), expected_count)
            for report in reports:
                with self.subTest(card=card, person=report.recipient.sheet):
                    self.assertLess(len(app.render_html(report).encode("utf-8")), 80_000)
                    source = wb[report.recipient.sheet]
                    purchases, payments = [], []
                    for row in range(6, source.max_row + 1):
                        when = source.cell(row, start).value
                        value = source.cell(row, share_column).value
                        if isinstance(when, datetime):
                            installment = source.cell(row, start + 2).value
                            purchases.append(app.Purchase(when.date(), source.cell(row, start + 1).value.strip(),
                                "—" if installment is None or installment == 0 else str(installment),
                                app._money(source.cell(row, start + 3).value, "fixture"), app._money(value, "fixture")))
                        elif when == "PAGAMENTOS":
                            payments.append(app._money(value, "fixture"))
                        elif when in ("TOTAL", "TOTAL GERAL"):
                            expected = Decimal(str(value or 0)).quantize(Decimal("0.01"), rounding="ROUND_HALF_UP")
                            self.assertEqual(report.spending if when == "TOTAL" else report.balance, expected)
                    self.assertEqual(report.purchases, tuple(purchases))
                    self.assertEqual(report.payments, tuple(payments))


if __name__ == "__main__":
    unittest.main()
