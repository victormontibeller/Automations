"""Drive tests use fake APIs and synthetic non-card assets."""
import importlib.util
import unittest
from unittest.mock import MagicMock, patch


class DriveTests(unittest.TestCase):
    def test_root_exact_folder_and_paginated_file_lookup_escape_query_values(self):
        self.assertIsNotNone(importlib.util.find_spec("automation_core.drive"), "Reusable Drive module missing")
        from automation_core.drive import DriveClient
        service = MagicMock()
        service.files.return_value.get.return_value.execute.return_value = {"id": "root-id"}
        folder = {"id": "folder-id", "name": "Design's\\Assets", "mimeType": "application/vnd.google-apps.folder"}
        asset = {"id": "logo-id", "name": "logo.svg", "mimeType": "image/svg+xml"}
        service.files.return_value.list.return_value.execute.side_effect = [
            {"files": [folder]}, {"files": [], "nextPageToken": "next"}, {"files": [asset]},
        ]
        client = DriveClient(service)
        self.assertEqual(client.root_id(), "root-id")
        self.assertEqual(client.find_folder("root-id", "Design's\\Assets"), folder)
        self.assertEqual(client.find_file("parent'\\id", "logo.svg"), asset)
        service.files.return_value.get.assert_called_once_with(fileId="root", fields="id")
        calls = service.files.return_value.list.call_args_list
        self.assertEqual(calls[0].kwargs["q"], "'root-id' in parents and name = 'Design\\'s\\\\Assets' and trashed = false and mimeType = 'application/vnd.google-apps.folder'")
        self.assertEqual(calls[1].kwargs["q"], "'parent\\'\\\\id' in parents and name = 'logo.svg' and trashed = false")
        self.assertEqual(calls[2].kwargs["q"], calls[1].kwargs["q"])
        self.assertNotIn("pageToken", calls[1].kwargs)
        self.assertEqual(calls[2].kwargs["pageToken"], "next")
        for call in calls:
            self.assertEqual(call.kwargs["fields"], "nextPageToken,files(id,name,mimeType)")
            self.assertTrue(call.kwargs["supportsAllDrives"])
            self.assertTrue(call.kwargs["includeItemsFromAllDrives"])
        for call in service.files.return_value.list.return_value.execute.call_args_list:
            self.assertEqual(call.kwargs, {"num_retries": 0})
        client.close()
        service._http.close.assert_called_once()

    def test_missing_is_none_but_duplicates_across_pages_are_ambiguous(self):
        from automation_core import drive
        self.assertTrue(hasattr(drive, "AmbiguousItem"), "Duplicate names must not silently select an item")
        for method in ("find_folder", "find_file"):
            service = MagicMock()
            service.files.return_value.list.return_value.execute.return_value = {"files": []}
            self.assertIsNone(getattr(drive.DriveClient(service), method)("parent", "asset"))
            service.files.return_value.list.return_value.execute.side_effect = [
                {"files": [{"id": "one"}], "nextPageToken": "next"}, {"files": [{"id": "two"}]},
            ]
            with self.assertRaises(drive.AmbiguousItem):
                getattr(drive.DriveClient(service), method)("parent", "asset")
            self.assertEqual(service.files.return_value.list.call_count, 3)

    def test_provider_failures_and_cleanup_have_safe_diagnostics(self):
        from automation_core import drive
        self.assertTrue(hasattr(drive, "DriveError"), "Drive failures need safe library diagnostics")
        for method, args, endpoint in (("root_id", (), "get"), ("find_file", ("parent", "asset"), "list")):
            service = MagicMock()
            getattr(service.files.return_value, endpoint).return_value.execute.side_effect = RuntimeError("PRIVATE_PROVIDER")
            client = drive.DriveClient(service)
            with self.assertRaises(drive.DriveError) as raised:
                getattr(client, method)(*args)
            self.assertNotIn("PRIVATE", str(raised.exception))
            service._http.close.side_effect = RuntimeError("PRIVATE_CLOSE")
            client.close()
            client.close()
            service._http.close.assert_called_once()

    def test_download_bytes_uses_all_chunks_without_retries_and_wraps_media_errors(self):
        from automation_core.drive import DriveClient, DriveError
        self.assertTrue(hasattr(DriveClient, "download_bytes"), "Drive must return arbitrary file bytes")
        service = MagicMock()
        counts = []
        class Downloader:
            def __init__(self, target, request):
                self.target = target
                self.chunks = iter((b"<svg>", b"</svg>"))
            def next_chunk(self, num_retries):
                counts.append(num_retries)
                self.target.write(next(self.chunks))
                return None, len(counts) == 2
        with patch("googleapiclient.http.MediaIoBaseDownload", Downloader):
            self.assertEqual(DriveClient(service).download_bytes("logo-id"), b"<svg></svg>")
        self.assertEqual(counts, [0, 0])
        service.files.return_value.get_media.assert_called_once_with(fileId="logo-id")
        for stage in ("get_media", "chunk"):
            service = MagicMock()
            with patch("googleapiclient.http.MediaIoBaseDownload") as downloader:
                if stage == "get_media":
                    service.files.return_value.get_media.side_effect = RuntimeError("PRIVATE_MEDIA")
                else:
                    downloader.return_value.next_chunk.side_effect = RuntimeError("PRIVATE_MEDIA")
                with self.assertRaises(DriveError) as raised:
                    DriveClient(service).download_bytes("logo-id")
                self.assertNotIn("PRIVATE", str(raised.exception))
                if stage == "chunk":
                    downloader.return_value.next_chunk.assert_called_once_with(num_retries=0)

    def test_malformed_root_or_listing_responses_are_safe_errors(self):
        from automation_core.drive import DriveClient, DriveError
        for response in ({}, {"id": ""}, {"id": 42}, None):
            with self.subTest(root=response):
                service = MagicMock()
                service.files.return_value.get.return_value.execute.return_value = response
                with self.assertRaises(DriveError):
                    DriveClient(service).root_id()
        for response in (None, {"files": None}, {"files": ["PRIVATE_ITEM"]}, {"files": [{}]}):
            with self.subTest(listing=response):
                service = MagicMock()
                service.files.return_value.list.return_value.execute.return_value = response
                with self.assertRaises(DriveError) as raised:
                    DriveClient(service).find_file("parent", "asset")
                self.assertNotIn("PRIVATE", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
