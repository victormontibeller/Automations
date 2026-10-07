"""Gmail transport for caller-composed messages; never retries a send."""
import base64
from dataclasses import dataclass
from email.message import EmailMessage


class GmailError(RuntimeError):
    """Safe Gmail integration diagnostic; never a raw provider payload."""


class InvalidMessage(GmailError):
    """Local MIME validation failed; no send was attempted."""


class GmailProfileError(GmailError):
    """The authenticated profile could not be read."""


class GmailSendRejected(GmailError):
    """Definitive API rejection, not acceptance."""


class GmailSendUncertain(GmailError):
    """Acceptance is unknown. Reconcile externally before any resend."""


@dataclass(frozen=True)
class SendResult:
    """Provider acceptance, not delivery/read confirmation or RFC Message-ID."""
    message_id: str
    thread_id: str | None = None


class GmailClient:
    """Owns the injected Google service; recipient/content policy stays outside."""

    def __init__(self, service):
        self.service = service
        self._closed = False
        try:
            profile = service.users().getProfile(userId="me").execute(num_retries=0)
            self.email_address = profile["emailAddress"]
            if not isinstance(self.email_address, str) or not self.email_address.strip():
                raise ValueError("Missing profile address")
        except Exception as exc:
            self.close()
            raise GmailProfileError("Unable to read the authenticated Gmail profile.") from exc

    def send_message(self, message: EmailMessage) -> SendResult:
        try:
            if not isinstance(message, EmailMessage) or not message.get("From"):
                raise ValueError("Expected a composed EmailMessage with From")
            recipients = []
            for name in ("From", "To", "Cc", "Bcc"):
                for header in message.get_all(name, []):
                    if header.defects or not header.addresses:
                        raise ValueError("Invalid address header")
                    if any(not address.username or not address.domain for address in header.addresses):
                        raise ValueError("Incomplete mailbox")
                    if name != "From":
                        recipients.extend(header.addresses)
            if not recipients:
                raise ValueError("No recipients")
            raw = base64.urlsafe_b64encode(message.as_bytes()).decode("ascii")
        except Exception as exc:
            raise InvalidMessage("Expected a serializable EmailMessage with valid sender and recipients.") from exc
        try:
            response = self.service.users().messages().send(
                userId="me", body={"raw": raw}
            ).execute(num_retries=0)
        except Exception as exc:
            status = getattr(getattr(exc, "resp", None), "status", None)
            if isinstance(status, int) and 400 <= status < 500:
                raise GmailSendRejected("Gmail rejected the message.") from exc
            raise GmailSendUncertain("Gmail acceptance is unknown; do not resend without reconciliation.") from exc
        if not isinstance(response, dict) or not isinstance(response.get("id"), str) or not response["id"].strip():
            raise GmailSendUncertain("Gmail did not return a message acknowledgement; acceptance is unknown.")
        thread_id = response.get("threadId")
        if not isinstance(thread_id, str) or not thread_id.strip():
            thread_id = None
        return SendResult(response["id"], thread_id)

    def close(self) -> None:
        # Best effort: cleanup must never turn accepted mail into a retry signal.
        if self._closed:
            return
        self._closed = True
        try:
            http = getattr(self.service, "_http", None)
            close = getattr(http, "close", None)
            if close:
                close()
        except Exception:
            pass
