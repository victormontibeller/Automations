"""The shared package must be usable without cards, spreadsheets or Hermes."""
import ast
import importlib.metadata
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest


class DependencyDirectionTests(unittest.TestCase):
    def test_library_has_no_application_imports_or_path_hacks(self):
        root = Path(__file__).resolve().parents[1] / "src/automation_core"
        files = list(root.glob("*.py"))
        self.assertTrue(files)
        forbidden = {"resumos_cartoes", "cartoes", "hermes", "openpyxl"}
        for path in files:
            source = path.read_text(encoding="utf-8")
            tree = ast.parse(source)
            for node in ast.walk(tree):
                imports = []
                if isinstance(node, ast.Import):
                    imports = [n.name for n in node.names]
                elif isinstance(node, ast.ImportFrom) and node.module:
                    imports = [node.module]
                self.assertFalse(forbidden.intersection(n.split(".")[0] for n in imports), path.name)
                if isinstance(node, ast.Attribute):
                    self.assertNotEqual(ast.unparse(node), "sys.path", path.name)
            for term in ("HERMES_HOME", "Cartão", ".xlsx", "personal_copy_email"):
                self.assertNotIn(term, source, path.name)
        requirements = importlib.metadata.requires("automation-core") or []
        self.assertFalse(any("cartoes" in r or "openpyxl" in r for r in requirements))

    def test_installed_library_imports_outside_checkout_without_application(self):
        code = '''
import importlib.abc
import sys
class NoApplication(importlib.abc.MetaPathFinder):
    def find_spec(self, fullname, path=None, target=None):
        if fullname.split('.')[0] in {'resumos_cartoes', 'cartoes', 'hermes', 'openpyxl', 'google', 'googleapiclient'}:
            raise AssertionError('Unexpected eager dependency: ' + fullname)
sys.meta_path.insert(0, NoApplication())
from automation_core.google_auth import build_service
from automation_core.gmail import GmailClient, SendResult
from automation_core.drive import DriveClient
assert SendResult('synthetic').message_id == 'synthetic'
'''
        with tempfile.TemporaryDirectory() as outside:
            result = subprocess.run([sys.executable, "-I", "-c", code], cwd=outside, capture_output=True, text=True, timeout=30)
        self.assertEqual(result.returncode, 0, result.stderr)


if __name__ == "__main__":
    unittest.main()
