"""Contracts for the isolated, no-credentials verification gate."""
import importlib.util
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[2]


class VerificationGateTests(unittest.TestCase):
    def load_gate(self):
        path = ROOT / "tools/verify.py"
        self.assertTrue(path.is_file(), "Reproducible verification gate is missing")
        spec = importlib.util.spec_from_file_location("verification_gate", path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module

    def test_snapshot_includes_public_sources_but_never_runtime_inputs(self):
        gate = self.load_gate()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory).resolve() / "repo"
            dest = Path(directory).resolve() / "snapshot"
            public = ("cartoes/src/app.py", "cartoes/tests/test_app.py", "cartoes/pyproject.toml",
                      "cartoes/config.example.json", "cartoes/cartoes.py", "cartoes/hermes/launcher.py",
                      "libs/automation_core/src/automation_core/__init__.py", "libs/automation_core/tests/test_core.py",
                      "libs/automation_core/pyproject.toml", "tools/verify.py", "requirements/verification.txt")
            private = ("cartoes/config.json", "cartoes/.env", "cartoes/inputs/private.xlsx",
                       "cartoes/var/deliveries.sqlite3", "cartoes/outputs/private.html", "cartoes/.venv/token.json",
                       "cartoes/src/__pycache__/app.pyc", "cartoes/src/app.egg-info/PKG-INFO")
            for name in (*public, *private):
                path = root / name
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("synthetic")
            gate.snapshot(root, dest)
            for name in public:
                self.assertTrue((dest / name).is_file(), name)
            for name in private:
                self.assertFalse((dest / name).exists(), name)

    def test_snapshot_copies_only_allowlisted_public_assets_without_links(self):
        gate = self.load_gate()
        with tempfile.TemporaryDirectory() as directory:
            base = Path(directory).resolve()
            root, dest = base / "repo", base / "snapshot"
            public = ("cartoes/src/resumos_cartoes/__init__.py",
                      "cartoes/tests/test_public.py", "cartoes/config.example.json",
                      "cartoes/tests/fixtures/anonymous_reports.json",
                      "cartoes/tests/fixtures/golden_reports.json",
                      "docs/architecture.md", "docs/operations.md",
                      "requirements/verification.txt", "tools/verify.py")
            private = tuple(f"{tree}/{name}" for tree in (
                "cartoes/src", "cartoes/src/resumos_cartoes", "cartoes/tests", "cartoes/hermes",
                "libs/automation_core/src", "libs/automation_core/src/automation_core",
                "libs/automation_core/tests", "tools", "requirements", "docs",
            ) for name in (".env", "config.json", "token.json",
                           "nested/.env", "nested/config.json", "nested/token.json",
                           "secrets/private.py", ".venv/lib/private.py")) + (
                "cartoes/tests/fixtures/token.json",
                "cartoes/tests/fixtures/private.json",
                "cartoes/src/resumos_cartoes/config.json",
                "docs/private.md", "requirements/private.txt",
            )
            for name in (*public, *private):
                path = root / name
                path.parent.mkdir(parents=True, exist_ok=True)
                path.write_text("PUBLIC" if name in public else "SYNTHETIC_PRIVATE")
            # Prefix-sharing sibling and valid-looking names defeat suffix-only filters.
            outside = base / "repo-private"
            outside.mkdir()
            (outside / "secret.py").write_text("SYNTHETIC_EXTERNAL")
            (root / "cartoes/src/leaked.py").symlink_to(outside / "secret.py")
            (root / "cartoes/tests/external").symlink_to(outside, target_is_directory=True)
            (root / "README.md").symlink_to(outside / "secret.py")
            from unittest.mock import patch
            scan = gate.os.scandir
            def public_scan(path):
                self.assertNotIn(".venv", Path(path).parts)
                self.assertNotIn("secrets", Path(path).parts)
                self.assertNotIn("nested", Path(path).parts)
                self.assertTrue(Path(path).is_relative_to(root))
                return scan(path)
            with patch.object(gate.os, "scandir", side_effect=public_scan):
                gate.snapshot(root, dest)
            copied = {str(path.relative_to(dest)) for path in dest.rglob("*") if path.is_file()}
            self.assertEqual(copied, set(public))
            self.assertFalse(any(path.is_symlink() for path in dest.rglob("*")))

    def test_snapshot_never_follows_source_tree_or_ancestor_links(self):
        gate = self.load_gate()
        for linked in ("cartoes/src", "cartoes/src/resumos_cartoes", "cartoes/tests/fixtures",
                       "cartoes", "libs/automation_core/src/automation_core",
                       "libs/automation_core", "libs", "tools", ".github"):
            with self.subTest(linked=linked), tempfile.TemporaryDirectory() as directory:
                base = Path(directory).resolve()
                root, dest, outside = base / "repo", base / "snapshot", base / "repo-private"
                root.mkdir()
                outside.mkdir()
                for name in ("app.py", "src/app.py", "src/automation_core/__init__.py",
                             "automation_core/src/automation_core/__init__.py",
                             "__init__.py", "verify.py", "workflows/tests.yml", "README.md",
                             "anonymous_reports.json", "golden_reports.json"):
                    path = outside / name
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_text("SYNTHETIC_EXTERNAL")
                link = root / linked
                link.parent.mkdir(parents=True, exist_ok=True)
                link.symlink_to(outside, target_is_directory=True)
                gate.snapshot(root, dest)
                self.assertEqual(list(dest.rglob("*")), [])

    def test_snapshot_rejects_linked_repository_root_or_its_ancestors(self):
        gate = self.load_gate()
        for suffix in ("", "repo"):
            with self.subTest(suffix=suffix), tempfile.TemporaryDirectory() as directory:
                base = Path(directory).resolve()
                actual = base / "actual"
                root = actual / suffix
                root.mkdir(parents=True)
                (root / "README.md").write_text("SYNTHETIC_EXTERNAL")
                alias = base / "alias"
                alias.symlink_to(actual, target_is_directory=True)
                dest = base / "snapshot"
                with self.assertRaises(ValueError):
                    gate.snapshot(alias / suffix, dest)
                self.assertFalse(dest.exists())

    def test_offline_guard_blocks_network_before_connect(self):
        gate = self.load_gate()
        code = gate.OFFLINE_GUARD + '''
import socket
for attempt in (lambda: socket.create_connection(('example.com', 443)),
                lambda: socket.socket().connect(('127.0.0.1', 1))):
    try:
        attempt()
    except RuntimeError as exc:
        assert str(exc) == 'Network disabled by verification gate'
    else:
        raise AssertionError('Network was not blocked')
'''
        result = subprocess.run([sys.executable, "-I", "-c", code], capture_output=True, text=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_documented_fake_client_example_runs_without_network(self):
        import re
        gate = self.load_gate()
        document = (ROOT / "docs/architecture.md").read_text()
        examples = re.findall(r"```python\n(.*?)```", document, flags=re.DOTALL)
        self.assertEqual(len(examples), 1)
        result = subprocess.run([sys.executable, "-I", "-c", gate.OFFLINE_GUARD + examples[0]], capture_output=True, text=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr)

    def test_snapshot_pins_are_complete_and_used_by_ci_and_docs(self):
        self.load_gate()
        path = ROOT / "requirements/verification.txt"
        self.assertTrue(path.is_file())
        requirements = [line for line in path.read_text().splitlines() if line and not line.startswith("#")]
        self.assertTrue(requirements)
        for line in requirements:
            self.assertRegex(line, r"^[a-zA-Z0-9_.-]+==[0-9][a-zA-Z0-9.]*$")
        for name in ("openpyxl", "google-api-python-client", "google-auth", "pip", "setuptools", "wheel"):
            self.assertTrue(any(line.startswith(name + "==") for line in requirements), name)
        self.assertIn("tools/verify.py", (ROOT / ".github/workflows/tests.yml").read_text())
        self.assertIn("requirements/verification.txt", (ROOT / "docs/operations.md").read_text())
