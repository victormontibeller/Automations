"""Exact-parent/name Drive lookup, with no application path conventions."""


class DriveError(RuntimeError):
    """Safe Drive operational diagnostic."""


class AmbiguousItem(DriveError):
    """More than one item matches the exact parent and name."""


def _query_value(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


class DriveClient:
    """Owns an injected Google service. Lookups never search alternate paths."""

    def __init__(self, service):
        self.service = service
        self._closed = False

    def root_id(self) -> str:
        try:
            root_id = self.service.files().get(fileId="root", fields="id").execute(num_retries=0)["id"]
            if not isinstance(root_id, str) or not root_id.strip():
                raise ValueError("Missing root identifier")
            return root_id
        except Exception as exc:
            raise DriveError("Unable to read the Drive root.") from exc

    def find_folder(self, parent_id: str, name: str) -> dict | None:
        return self._find(parent_id, name, "application/vnd.google-apps.folder")

    def find_file(self, parent_id: str, name: str) -> dict | None:
        """Find a unique named item; the caller validates its required MIME type."""
        return self._find(parent_id, name)

    def _find(self, parent_id: str, name: str, mime_type: str | None = None) -> dict | None:
        query = f"'{_query_value(parent_id)}' in parents and name = '{_query_value(name)}' and trashed = false"
        if mime_type:
            query += f" and mimeType = '{_query_value(mime_type)}'"
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
            try:
                response = self.service.files().list(**params).execute(num_retries=0)
                page_files = response.get("files", [])
                if not isinstance(page_files, list) or any(
                    not isinstance(item, dict) or not isinstance(item.get("id"), str) or not item["id"].strip()
                    for item in page_files
                ):
                    raise ValueError("Invalid Drive item response")
                files.extend(page_files)
                page_token = response.get("nextPageToken")
            except Exception as exc:
                raise DriveError("Unable to list Drive items.") from exc
            if not page_token:
                if len(files) > 1:
                    raise AmbiguousItem("Multiple Drive items match the requested parent and name.")
                return files[0] if files else None

    def download_bytes(self, file_id: str) -> bytes:
        """Download media as bytes; no conversion, parsing or disk persistence."""
        from io import BytesIO
        from googleapiclient.http import MediaIoBaseDownload

        try:
            with BytesIO() as buffer:
                request = self.service.files().get_media(fileId=file_id)
                downloader = MediaIoBaseDownload(buffer, request)
                done = False
                while not done:
                    _, done = downloader.next_chunk(num_retries=0)
                return buffer.getvalue()
        except Exception as exc:
            raise DriveError("Unable to download Drive media.") from exc

    def close(self) -> None:
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
