"""Build Google services from an explicit, read-only authorized-user token file."""
from pathlib import Path


class GoogleAuthError(RuntimeError):
    """Credentials or service construction failed before an operation began."""


def build_service(api: str, version: str, token_file: Path):
    """Load caller-selected credentials; never discover or persist tokens."""
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    try:
        credentials = Credentials.from_authorized_user_file(str(token_file))
        return build(api, version, credentials=credentials, cache_discovery=False)
    except Exception as exc:
        raise GoogleAuthError("Unable to initialize the Google service with the supplied credentials.") from exc
