"""Cards workbook path/MIME policy over the reusable Drive client."""
from automation_core import google_auth
from automation_core.drive import AmbiguousItem, DriveClient

from .config import Settings
from .errors import ReportError
from .scheduling import validate_month


def drive_workbook_bytes(settings: Settings, month: str) -> bytes:
    """Fetch exactly My Drive/<folder>/<year>/<month>.xlsx, with no fallback."""
    month = validate_month(month)
    year = month[:4]
    client = None
    description = settings.drive_folder_name
    try:
        client = DriveClient(google_auth.build_service("drive", "v3", settings.google_token_file))
        folder = client.find_folder(client.root_id(), settings.drive_folder_name)
        if folder is None:
            raise ReportError(f"Pasta Drive/{settings.drive_folder_name} não encontrada no Meu Drive.")
        description = f"Drive/{settings.drive_folder_name}/{year}"
        year_folder = client.find_folder(folder["id"], year)
        if year_folder is None:
            raise ReportError(f"Pasta Drive/{settings.drive_folder_name}/{year} não encontrada.")
        filename = f"{month}.xlsx"
        description = f"Drive/{settings.drive_folder_name}/{year}/{filename}"
        workbook = client.find_file(year_folder["id"], filename)
        if workbook is None:
            raise ReportError(f"Arquivo {filename} não encontrado em Drive/{settings.drive_folder_name}/{year}.")
        if workbook.get("mimeType") != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
            raise ReportError(f"O arquivo {filename} no Drive não é um Excel .xlsx válido.")
        return client.download_bytes(workbook["id"])
    except AmbiguousItem as exc:
        raise ReportError(f"Há mais de um item chamado {description} no Google Drive; deixe apenas um caminho correspondente.") from exc
    except ReportError:
        raise
    except Exception as exc:
        raise ReportError("Falha ao localizar ou baixar a planilha no Google Drive. Verifique a autorização OAuth do Drive e tente novamente.") from exc
    finally:
        if client is not None:
            client.close()
