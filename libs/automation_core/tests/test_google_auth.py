"""Synthetic authentication tests: never read a real credential file."""
import importlib.util
import io
from pathlib import Path
import tempfile
import unittest
from contextlib import redirect_stdout, redirect_stderr
from unittest.mock import patch, sentinel


class GoogleAuthTests(unittest.TestCase):
    def test_explicit_token_path_builds_service_without_writing_or_printing(self):
        self.assertIsNotNone(importlib.util.find_spec("automation_core"), "Shared package must be installed")
        from automation_core.google_auth import build_service
        with tempfile.TemporaryDirectory() as directory:
            token = Path(directory) / "synthetic.json"
            token.write_text("synthetic sentinel, not real credentials", encoding="utf-8")
            before = token.read_bytes()
            output = io.StringIO()
            with patch("google.oauth2.credentials.Credentials.from_authorized_user_file", return_value=sentinel.credentials) as load, patch("googleapiclient.discovery.build", return_value=sentinel.service) as build, redirect_stdout(output), redirect_stderr(output):
                result = build_service("gmail", "v1", token)
            self.assertIs(result, sentinel.service)
            load.assert_called_once_with(str(token))
            build.assert_called_once_with("gmail", "v1", credentials=sentinel.credentials, cache_discovery=False)
            self.assertEqual(token.read_bytes(), before)
            self.assertEqual(output.getvalue(), "")
            self.assertEqual(list(Path(directory).iterdir()), [token])

    def test_auth_failures_are_library_errors_without_provider_payload(self):
        from automation_core import google_auth
        self.assertTrue(hasattr(google_auth, "GoogleAuthError"), "Auth failures need a reusable safe error")
        for failure in (FileNotFoundError("PRIVATE_PATH"), ValueError("PRIVATE_TOKEN"), RuntimeError("PRIVATE_SERVER")):
            with self.subTest(failure=type(failure).__name__), patch("google.oauth2.credentials.Credentials.from_authorized_user_file", side_effect=failure), patch("googleapiclient.discovery.build") as build:
                with self.assertRaises(google_auth.GoogleAuthError) as raised:
                    google_auth.build_service("drive", "v3", Path("synthetic.json"))
                self.assertNotIn("PRIVATE", str(raised.exception))
                build.assert_not_called()
        with patch("google.oauth2.credentials.Credentials.from_authorized_user_file", return_value=sentinel.credentials), patch("googleapiclient.discovery.build", side_effect=RuntimeError("PRIVATE_SERVER")):
            with self.assertRaises(google_auth.GoogleAuthError) as raised:
                google_auth.build_service("drive", "v3", Path("synthetic.json"))
            self.assertNotIn("PRIVATE", str(raised.exception))


if __name__ == "__main__":
    unittest.main()
