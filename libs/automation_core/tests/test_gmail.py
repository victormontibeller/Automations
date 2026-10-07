"""Caller-owned messages and fake Gmail API responses only."""
import base64
from email import policy
from email.message import EmailMessage
from email.parser import BytesParser
import importlib.util
import unittest
from unittest.mock import MagicMock


def fake_service():
    service = MagicMock()
    service.users.return_value.getProfile.return_value.execute.return_value = {"emailAddress": "robot@example.com"}
    service.users.return_value.messages.return_value.send.return_value.execute.return_value = {"id": "message-123", "threadId": "thread-456"}
    return service


def message():
    mail = EmailMessage()
    mail["From"] = "Release Robot <robot@example.com>"
    mail["To"] = "Team One <team@example.com>, Team Two <other@example.com>"
    mail["Cc"] = "observer@example.com"
    mail["Bcc"] = "audit@example.com"
    mail["Subject"] = "Synthetic release notice"
    mail.set_content("Release ready.")
    mail.add_alternative("<p>Release ready.</p>", subtype="html")
    return mail


class GmailTests(unittest.TestCase):
    def test_non_card_message_preserves_caller_mime_and_returns_provider_ids(self):
        self.assertIsNotNone(importlib.util.find_spec("automation_core.gmail"), "Reusable Gmail module missing")
        from automation_core.gmail import GmailClient, SendResult
        service = fake_service()
        client = GmailClient(service)
        self.assertEqual(client.email_address, "robot@example.com")
        mail = message()
        original = mail.as_bytes()
        result = client.send_message(mail)
        self.assertEqual(result, SendResult(message_id="message-123", thread_id="thread-456"))
        send = service.users.return_value.messages.return_value.send
        send.assert_called_once()
        self.assertEqual(send.call_args.kwargs["userId"], "me")
        raw = base64.urlsafe_b64decode(send.call_args.kwargs["body"]["raw"])
        self.assertEqual(raw, original)
        parsed = BytesParser(policy=policy.default).parsebytes(raw)
        for header in ("From", "To", "Cc", "Bcc", "Subject"):
            self.assertEqual(str(parsed[header]), str(mail[header]))
        self.assertEqual(mail.as_bytes(), original)
        send.return_value.execute.assert_called_once_with(num_retries=0)
        service.users.return_value.getProfile.return_value.execute.assert_called_once_with(num_retries=0)
        client.close()
        service._http.close.assert_called_once()

    def test_rejected_uncertain_and_missing_acknowledgement_are_distinct_no_retries(self):
        from automation_core import gmail
        self.assertTrue(hasattr(gmail, "GmailSendUncertain"), "Uncertain sends require an explicit outcome")
        for status in (400, 403, 429, 500, 503, None):
            with self.subTest(status=status):
                service = fake_service()
                error = TimeoutError("PRIVATE_PROVIDER_PAYLOAD")
                if status is not None:
                    error.resp = type("Response", (), {"status": status})()
                request = service.users.return_value.messages.return_value.send.return_value
                request.execute.side_effect = error
                client = gmail.GmailClient(service)
                expected = gmail.GmailSendRejected if status is not None and status < 500 else gmail.GmailSendUncertain
                with self.assertRaises(expected) as raised:
                    client.send_message(message())
                self.assertNotIn("PRIVATE", str(raised.exception))
                request.execute.assert_called_once_with(num_retries=0)
                service.users.return_value.messages.return_value.send.assert_called_once()
        for response in ({}, None, {"id": ""}, {"id": "  "}, {"id": 7}, {"threadId": "thread-only"}):
            with self.subTest(response=response):
                service = fake_service()
                request = service.users.return_value.messages.return_value.send.return_value
                request.execute.return_value = response
                with self.assertRaises(gmail.GmailSendUncertain):
                    gmail.GmailClient(service).send_message(message())
                request.execute.assert_called_once_with(num_retries=0)
        service = fake_service()
        service.users.return_value.messages.return_value.send.return_value.execute.return_value = {"id": "ack"}
        self.assertEqual(gmail.GmailClient(service).send_message(message()), gmail.SendResult("ack"))

    def test_malformed_optional_thread_id_does_not_invalidate_acceptance(self):
        from automation_core.gmail import GmailClient, SendResult
        for thread in (None, "", "  ", 123, [], {}, False):
            with self.subTest(thread=thread):
                service = fake_service()
                request = service.users.return_value.messages.return_value.send.return_value
                request.execute.return_value = {"id": "ack", "threadId": thread}
                self.assertEqual(GmailClient(service).send_message(message()), SendResult("ack", None))
                request.execute.assert_called_once_with(num_retries=0)

    def test_profile_failure_and_close_failure_never_leak_or_mask_outcomes(self):
        from automation_core import gmail
        self.assertTrue(hasattr(gmail, "GmailProfileError"), "Profile errors need a safe diagnostic")
        for response in ({}, {"emailAddress": ""}, {"emailAddress": None}):
            with self.subTest(response=response):
                service = fake_service()
                service.users.return_value.getProfile.return_value.execute.return_value = response
                service._http.close.side_effect = RuntimeError("PRIVATE_CLOSE")
                with self.assertRaises(gmail.GmailProfileError) as raised:
                    gmail.GmailClient(service)
                self.assertNotIn("PRIVATE", str(raised.exception))
                service._http.close.assert_called_once()
        service = fake_service()
        service.users.return_value.getProfile.return_value.execute.side_effect = RuntimeError("PRIVATE_PROFILE")
        service._http.close.side_effect = RuntimeError("PRIVATE_CLOSE")
        with self.assertRaises(gmail.GmailProfileError):
            gmail.GmailClient(service)
        service._http.close.assert_called_once()
        service = fake_service()
        service._http.close.side_effect = RuntimeError("PRIVATE_CLOSE")
        client = gmail.GmailClient(service)
        self.assertEqual(client.send_message(message()).message_id, "message-123")
        client.close()
        client.close()
        service._http.close.assert_called_once()

    def test_message_validation_is_not_cards_recipient_policy(self):
        from automation_core import gmail
        self.assertTrue(hasattr(gmail, "InvalidMessage"), "Local validation must be distinguishable from uncertain sends")
        for header in ("To", "Cc", "Bcc"):
            with self.subTest(header=header):
                mail = EmailMessage()
                mail["From"] = "Robot <robot@example.com>"
                mail[header] = "One <one@example.com>, Two <two@example.com>"
                mail.set_content("Synthetic notice")
                service = fake_service()
                self.assertEqual(gmail.GmailClient(service).send_message(mail).message_id, "message-123")
                raw = service.users.return_value.messages.return_value.send.call_args.kwargs["body"]["raw"]
                parsed = BytesParser(policy=policy.default).parsebytes(base64.urlsafe_b64decode(raw))
                for key in ("To", "Cc", "Bcc"):
                    self.assertEqual(parsed[key], mail[key])
        invalid = [None, "not a message", EmailMessage()]
        mail = message()
        del mail["To"]
        del mail["Cc"]
        del mail["Bcc"]
        invalid.append(mail)
        mail = message()
        mail.replace_header("To", "not-an-address")
        invalid.append(mail)
        for mail in invalid:
            with self.subTest(mail=type(mail).__name__):
                service = fake_service()
                with self.assertRaises(gmail.InvalidMessage):
                    gmail.GmailClient(service).send_message(mail)
                service.users.return_value.messages.return_value.send.assert_not_called()


if __name__ == "__main__":
    unittest.main()
