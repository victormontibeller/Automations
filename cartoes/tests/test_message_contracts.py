"""Delivery envelope characterization; no provider or credential access."""
from email import policy
from email.parser import BytesParser
from pathlib import Path
import unittest
from unittest.mock import MagicMock, patch

from resumos_cartoes import alerts, gmail
from resumos_cartoes.config import Settings
from resumos_cartoes.delivery import make_message
from resumos_cartoes.models import Recipient, Report


class MessageContractTests(unittest.TestCase):
    def test_serialized_summary_deduplicates_cc_and_never_adds_bcc(self):
        report = Report("black", "2026-09", Recipient("Pessoa01", "Participante01", "person@example.com"), (), ())
        for copy, expected in (("copy@example.com", "copy@example.com"), ("PERSON@example.com", None)):
            with self.subTest(copy=copy):
                message = make_message(report, "owner@example.com", "Hermes", "person@example.com", cc=copy)
                parsed = BytesParser(policy=policy.default).parsebytes(message.as_bytes())
                self.assertEqual(parsed["To"], "person@example.com")
                self.assertEqual(parsed["Cc"], expected)
                self.assertIsNone(parsed["Bcc"])
                self.assertEqual(list(parsed.iter_attachments()), [])
                self.assertEqual(parsed.get_body(preferencelist=("plain",)).get_content_type(), "text/plain")
                self.assertEqual(parsed.get_body(preferencelist=("html",)).get_content_type(), "text/html")

    def test_owner_alert_deduplicates_copy_against_authenticated_owner(self):
        settings = Settings(Path("input"), "local", "Cartão", Path("output"), Path("state"),
                            Path("fake-token.json"), "Hermes", (), personal_copy_email="OWNER@example.com")
        connection = MagicMock()
        connection.email_address = "owner@example.com"
        with patch.object(gmail, "gmail_connect", return_value=connection):
            alerts.notify_owner(settings, "black", "2026-09", "configuração", "Synthetic error")
        connection.send_message.assert_called_once()
        connection.close.assert_called_once()
        message = connection.send_message.call_args.args[0]
        parsed = BytesParser(policy=policy.default).parsebytes(message.as_bytes())
        self.assertEqual(parsed["To"], "owner@example.com")
        self.assertIsNone(parsed["Cc"])
        self.assertIsNone(parsed["Bcc"])


if __name__ == "__main__":
    unittest.main()
