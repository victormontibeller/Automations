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

    def test_monthly_copy_and_signature_survive_mime_round_trip(self):
        for card in ("black", "latam"):
            with self.subTest(card=card):
                report = Report(card, "2026-09", Recipient("Pessoa01", "Participante01", "person@example.com"),
                                (), (), "Titular de teste", "Pagamento: referência de teste")
                message = make_message(report, "owner@example.com", "Hermes", "person@example.com")
                parsed = BytesParser(policy=policy.default).parsebytes(message.as_bytes())
                introduction = (
                    "Olá, Pessoa01!",
                    f"Segue o seu resumo do cartão {card.upper()}, referente a 09/2026, "
                    "com os lançamentos, pagamentos registrados e saldo.",
                    "Se tiver alguma dúvida ou identificar alguma diferença, fale com o Titular de teste.",
                )
                for subtype in ("plain", "html"):
                    content = parsed.get_body(preferencelist=(subtype,)).get_content()
                    for paragraph in introduction:
                        self.assertIn(paragraph, content)
                    self.assertNotIn("Sou o Hermes", content)
                    self.assertNotIn("assistente pessoal", content)
                    self.assertNotIn("Tudo bem?", content)
                    self.assertLess(content.index(introduction[0]), content.index(introduction[1]))
                    self.assertLess(content.index(introduction[1]), content.index(introduction[2]))
                    self.assertLess(content.index(introduction[2]), content.index("TOTAL GERAL"))
                    self.assertLess(content.index("TOTAL GERAL"), content.index("Abraço,"))
                    self.assertIn("Pagamento: referência de teste", content)
                    self.assertLess(content.index("TOTAL GERAL"), content.index("Pagamento: referência de teste"))
                    self.assertLess(content.index("Pagamento: referência de teste"), content.index("Abraço,"))
                    self.assertIn("Abraço,\nHermes" if subtype == "plain" else "Abraço,<br>Hermes", content)
                    if subtype == "plain":
                        self.assertTrue(content.rstrip().endswith("Abraço,\nHermes"))

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
