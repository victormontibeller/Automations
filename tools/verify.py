#!/usr/bin/env python3
"""Build and verify isolated installs of public sources, never runtime config.

Run with --python pointing to a Python >=3.11 with pip >=22.3. Only bootstrap
and dependency installation may use the network. Everything is disposable.
"""
from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile

ROOT = Path(__file__).resolve().parents[1]
PACKAGES = ("libs/automation_core", "cartoes")
OFFLINE_GUARD = '''
import sys
def _deny_network(event, args):
    if event in {'socket.connect', 'socket.getaddrinfo', 'socket.gethostbyname',
                 'socket.gethostbyaddr', 'socket.sendto', 'socket.sendmsg'}:
        raise RuntimeError('Network disabled by verification gate')
sys.addaudithook(_deny_network)
'''


def _unlinked(path: Path) -> bool:
    """Check ancestors before inspecting a child; never traverse a linked tree."""
    return not any(part.is_symlink() for part in (*reversed(path.parents), path))


def snapshot(root: Path, target: Path) -> None:
    """Copy an explicit public allowlist, without recursive directory traversal."""
    root = root.absolute()
    if not _unlinked(root):
        raise ValueError("Snapshot source and its ancestors must not be symlinks")
    root = root.resolve()
    # Patterns apply only to immediate files in these named code directories.
    # New packages/assets must opt in here, never via recursive suffix filtering.
    patterns = ("cartoes/src/*.py", "cartoes/src/resumos_cartoes/*.py",
                "cartoes/tests/test_*.py", "cartoes/hermes/*.py",
                "libs/automation_core/src/automation_core/*.py",
                "libs/automation_core/tests/test_*.py")
    files = ("README.md", ".github/workflows/tests.yml", "cartoes/cartoes.py",
             "cartoes/config.example.json", "tools/verify.py", "requirements/verification.txt",
             "docs/architecture.md", "docs/operations.md",
             "cartoes/tests/fixtures/anonymous_reports.json", "cartoes/tests/fixtures/golden_reports.json",
             *(f"{p}/{f}" for p in PACKAGES for f in ("pyproject.toml", "README.md")))
    candidates = [root / name for name in files]
    for pattern in patterns:
        parent = root / Path(pattern).parent
        if _unlinked(parent) and parent.resolve().is_relative_to(root) and parent.is_dir():
            candidates.extend(parent.glob(Path(pattern).name))
    for source in candidates:
        if (_unlinked(source) and source.resolve().is_relative_to(root)
                and source.is_file()):
            destination = target / source.relative_to(root)
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(source, destination, follow_symlinks=False)


def run(command: list[str], *, cwd: Path, env: dict) -> None:
    print("+ " + " ".join(map(str, command)), flush=True)
    subprocess.run(command, cwd=cwd, env=env, check=True)


def verify(python: str, mode: str, scratch: Path) -> None:
    """Use a fresh venv per mode; never install into the supplied interpreter."""
    env = os.environ.copy()
    env.pop("PYTHONPATH", None)
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    with tempfile.TemporaryDirectory(prefix="automations-verify-", dir=scratch) as temporary:
        work = Path(temporary)
        source = work / "source"
        outside = work / "outside"
        outside.mkdir()
        snapshot(ROOT, source)
        env["TMPDIR"] = str(work)
        env["HERMES_HOME"] = str(work / "no-credentials")
        requirements = source / "requirements/verification.txt"
        for install_mode in (("wheel", "editable") if mode == "both" else (mode,)):
            venv = work / install_mode
            run([python, "-m", "venv", "--without-pip", str(venv)], cwd=outside, env=env)
            executable = str(venv / "bin/python")
            run([python, "-m", "pip", "--python", executable, "install", "-r", str(requirements)], cwd=outside, env=env)
            pip = [executable, "-m", "pip"]
            local = [str(source / package) for package in PACKAGES]
            if install_mode == "wheel":
                wheels = work / "wheels"
                run([*pip, "wheel", "--no-deps", "--no-build-isolation", "--wheel-dir", str(wheels), *local], cwd=outside, env=env)
                built = sorted(str(path) for path in wheels.glob("*.whl"))
                if len(built) != len(PACKAGES):
                    raise RuntimeError("Expected both local wheels")
                run([*pip, "install", "--no-index", "--no-deps", *built], cwd=outside, env=env)
            else:
                run([*pip, "install", "--no-deps", "--no-build-isolation", "-e", local[0], "-e", local[1]], cwd=outside, env=env)
            run([*pip, "check"], cwd=outside, env=env)
            # A site-local audit hook also applies to -I child interpreters used by
            # packaging tests. It is never installed into a development/production venv.
            site = subprocess.check_output([executable, "-I", "-c", "import sysconfig; print(sysconfig.get_path('purelib'))"], text=True, cwd=outside, env=env).strip()
            (Path(site) / "sitecustomize.py").write_text(OFFLINE_GUARD, encoding="utf-8")
            check = '''
import importlib.metadata as metadata
import json
from pathlib import Path
import sys
import automation_core, resumos_cartoes, cartoes
from resumos_cartoes.cli import main
assert cartoes.main is main
source = Path(sys.argv[1]).resolve()
editable = sys.argv[2] == 'editable'
for name, module in (('automation-core', automation_core), ('resumos-cartoes', resumos_cartoes)):
    location = Path(module.__file__).resolve()
    assert location.is_relative_to(source) == editable, (name, location)
    direct = json.loads(metadata.distribution(name).read_text('direct_url.json'))
    assert bool(direct.get('dir_info', {}).get('editable')) == editable
    print(name, location)
pins = dict(line.split('==') for line in Path(sys.argv[3]).read_text().splitlines() if line and not line.startswith('#'))
for name, version in pins.items():
    assert metadata.version(name) == version, name
normalize = lambda name: name.lower().replace('_', '-').replace('.', '-')
expected = {normalize(name) for name in pins} | {'automation-core', 'resumos-cartoes'}
assert {normalize(dist.metadata['Name']) for dist in metadata.distributions()} == expected
import socket
try:
    socket.create_connection(('example.com', 443))
except RuntimeError as exc:
    assert str(exc) == 'Network disabled by verification gate'
else:
    raise AssertionError('Offline guard missing')
'''
            run([executable, "-I", "-c", check, str(source), install_mode, str(requirements)], cwd=outside, env=env)
            run([executable, "-I", "-m", "resumos_cartoes", "--help"], cwd=outside, env=env)
            run([str(venv / "bin/cartoes"), "--help"], cwd=outside, env=env)
            for package in PACKAGES:
                run([executable, "-I", "-m", "unittest", "discover", "-s", str(source / package / "tests"), "-v"], cwd=outside, env=env)
            print(f"PASS: {install_mode} installation, dependency closure, outside-checkout CLI and offline suites", flush=True)
    run(["git", "diff", "--check"], cwd=ROOT, env=env)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--python", required=True, help="Interpreter with venv and pip (never modified)")
    parser.add_argument("--scratch", required=True, type=Path, help="Existing disposable directory outside this checkout")
    parser.add_argument("--mode", choices=("wheel", "editable", "both"), default="both")
    args = parser.parse_args(argv)
    python = shutil.which(args.python)
    if python is None:
        parser.error("Interpreter was not found")
    scratch = args.scratch.expanduser().resolve()
    if not scratch.is_dir() or scratch == ROOT or scratch.is_relative_to(ROOT):
        parser.error("Scratch must be an existing directory outside this checkout")
    verify(os.path.abspath(python), args.mode, scratch)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
