"""Local Drive OAuth adapter for exact folder/year/month workbook lookup."""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

from .config import Settings
from .errors import ReportError
from .scheduling import validate_month


def _build_drive_service(token_file: Path):
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    credentials = Credentials.from_authorized_user_file(str(token_file))
    return build("drive", "v3", credentials=credentials, cache_discovery=False)


def _drive_query_value(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def _drive_children(service, parent_id: str, name: str, mime_type: str | None = None) -> list[dict]:
    query = f"'{parent_id}' in parents and name = '{_drive_query_value(name)}' and trashed = false"
    if mime_type:
        query += f" and mimeType = '{mime_type}'"
    files, page_token = [], None
    while True:
        params = {
            "q": query,
            "fields": "nextPageToken,files(id,name,mimeType)",
            "pageSize": 100,
            "supportsAllDrives": True,
            "includeItemsFromAllDrives": True,
        }
        if page_token:
            params["pageToken"] = page_token
        response = service.files().list(**params).execute(num_retries=0)
        files.extend(response.get("files", []))
        page_token = response.get("nextPageToken")
        if not page_token:
            return files


def _unique_drive_item(items: list[dict], description: str) -> dict | None:
    if not items:
        return None
    if len(items) > 1:
        raise ReportError(f"Há mais de um item chamado {description} no Google Drive; deixe apenas um caminho correspondente.")
    return items[0]


def _download_drive_file(service, file_id: str) -> bytes:
    from googleapiclient.http import MediaIoBaseDownload

    buffer = BytesIO()
    request = service.files().get_media(fileId=file_id)
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk(num_retries=0)
    return buffer.getvalue()


def drive_workbook_bytes(settings: Settings, month: str) -> bytes:
    """Fetch <month>.xlsx from My Drive/<folder>/<year> using Hermes OAuth."""
    month = validate_month(month)
    year = month[:4]
    service = None
    try:
        service = _build_drive_service(settings.google_token_file)
        root_id = service.files().get(fileId="root", fields="id").execute(num_retries=0)["id"]
        folder = _unique_drive_item(
            _drive_children(service, root_id, settings.drive_folder_name, "application/vnd.google-apps.folder"),
            settings.drive_folder_name,
        )
        if folder is None:
            raise ReportError(f"Pasta Drive/{settings.drive_folder_name} não encontrada no Meu Drive.")
        year_folder = _unique_drive_item(
            _drive_children(service, folder["id"], year, "application/vnd.google-apps.folder"),
            f"Drive/{settings.drive_folder_name}/{year}",
        )
        if year_folder is None:
            raise ReportError(f"Pasta Drive/{settings.drive_folder_name}/{year} não encontrada.")
        filename = f"{month}.xlsx"
        workbook = _unique_drive_item(
            _drive_children(service, year_folder["id"], filename),
            f"Drive/{settings.drive_folder_name}/{year}/{filename}",
        )
        if workbook is None:
            raise ReportError(f"Arquivo {filename} não encontrado em Drive/{settings.drive_folder_name}/{year}.")
        if workbook.get("mimeType") != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            raise ReportError(f"O arquivo {filename} no Drive não é um Excel .xlsx válido.")
        return _download_drive_file(service, workbook["id"])
    except ReportError:
        raise
    except Exception as exc:
        raise ReportError("Falha ao localizar ou baixar a planilha no Google Drive. Verifique a autorização OAuth do Drive e tente novamente.") from exc
    finally:
        if service is not None:
            http = getattr(service, "_http", None)
            close = getattr(http, "close", None)
            if close:
                close()
