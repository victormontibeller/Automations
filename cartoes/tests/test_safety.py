"""Safety boundaries for synthetic test runs; no private workbook is opened."""
import os
from pathlib import Path
import subprocess
import sys
import unittest
from unittest.mock import patch
import runpy

ROOT = Path(__file__).resolve().parents[2]


class TestSafety(unittest.TestCase):
    def test_network_guard_blocks_before_connecting(self):
        guard = ROOT / "tests/offline/sitecustomize.py"
        self.assertTrue(guard.is_file(), "Test-only network guard missing")
        code = """
import runpy, socket, sys
runpy.run_path(sys.argv[1])
try:
    socket.create_connection(('example.com', 443))
except RuntimeError as exc:
    assert str(exc) == 'Network disabled during tests'
else:
    raise AssertionError('Network guard did not block the connection')
"""
        result = subprocess.run([sys.executable, "-I", "-c", code, str(guard)],
                                capture_output=True, text=True, timeout=10)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_private_workbook_requires_explicit_opt_in(self):
        fixture = ROOT / "cartoes/inputs/2026-09.xlsx"
        original_exists = Path.exists

        def exists_with_fixture(path):
            return path == fixture or original_exists(path)

        for enabled in (False, True):
            with self.subTest(enabled=enabled), patch.dict(os.environ, {}, clear=True), \
                    patch.object(Path, "exists", exists_with_fixture):
                if enabled:
                    os.environ["CARTOES_TEST_PRIVATE_WORKBOOK"] = "1"
                namespace = runpy.run_path(str(ROOT / "cartoes/tests/test_cartoes.py"))
                method = namespace["ProvidedWorkbookTests"].test_all_24_blocks_match_independent_cached_personal_summaries
                self.assertEqual(getattr(method, "__unittest_skip__", False), not enabled)
