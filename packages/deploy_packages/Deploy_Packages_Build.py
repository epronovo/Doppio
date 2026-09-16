"""
Deploy_Packages_Build - discovers the deployable apps under packages/ and
builds the zip for one of them.

An app is "deployable" if its folder has a *_App.py (the Flask entry point
every packages/ app is built around - see e.g. ERP_Concur_App.py). Folders
without one, like etl_datalake, are plain scripts and are left out of the
picker.

The zip holds the app's own folder as-is - minus the venv, caches, and the
input/output/gnupg_home data folders, which the app recreates itself on
first run - plus a run_mac.command / run_windows.bat pair that sets up a
virtualenv, installs requirements.txt, and starts the app on the chosen
port. Nothing is built or installed here; that all happens the first time
the launcher runs on the target machine.
"""
from __future__ import annotations

import ast
import io
import os
import re
import zipfile
from dataclasses import dataclass
from pathlib import Path

BASE_DIR = Path(__file__).parent.resolve()
PACKAGES_DIR = BASE_DIR.parent
SELF_NAME = BASE_DIR.name

EXCLUDE_DIRS = {"__pycache__", ".venv", ".git", ".vscode", "input", "output",
                "gnupg_home"}
EXCLUDE_FILES = {".DS_Store"}
EXCLUDE_SUFFIXES = {".pyc"}

_PORT_RE = re.compile(r'--port["\']\s*,\s*type\s*=\s*int\s*,\s*default\s*=\s*(\d+)')


@dataclass
class AppInfo:
    key: str
    dir: Path
    entry: Path
    description: str
    default_port: int


def discover_apps() -> list[AppInfo]:
    """Every packages/ folder with a *_App.py, sorted by name."""
    apps = []
    for child in sorted(PACKAGES_DIR.iterdir()):
        if not child.is_dir() or child.name.startswith(".") or child.name == SELF_NAME:
            continue
        entries = sorted(child.glob("*_App.py"))
        if not entries:
            continue
        entry = entries[0]
        source = entry.read_text(encoding="utf-8", errors="ignore")
        apps.append(AppInfo(
            key=child.name, dir=child, entry=entry,
            description=_describe(source), default_port=_default_port(source)))
    return apps


def get_app(key: str) -> AppInfo | None:
    """Look up one app by its packages/ folder name, or None if it is not
    a deployable app - never trust a key from the request without this."""
    for app in discover_apps():
        if app.key == key:
            return app
    return None


def _describe(source: str) -> str:
    try:
        doc = ast.get_docstring(ast.parse(source)) or ""
    except SyntaxError:
        doc = ""
    first_para = doc.split("\n\n", 1)[0].replace("\n", " ").strip()
    if " - " in first_para:
        first_para = first_para.split(" - ", 1)[1]
    sentence = first_para.split(". ", 1)[0].strip()
    if sentence and not sentence.endswith("."):
        sentence += "."
    return sentence or "Flask app."


def _default_port(source: str) -> int:
    m = _PORT_RE.search(source)
    return int(m.group(1)) if m else 5000


def _references_ionapi(pkg_dir: Path) -> bool:
    for py in pkg_dir.glob("*.py"):
        try:
            if "ionapi" in py.read_text(encoding="utf-8", errors="ignore").lower():
                return True
        except OSError:
            pass
    return False


def build_zip(app: AppInfo, port: int) -> io.BytesIO:
    """Zip app.dir (minus the excluded folders) plus the generated
    launchers, rooted at a top-level folder named after the app."""
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, dirs, files in os.walk(app.dir):
            dirs[:] = [d for d in dirs if d not in EXCLUDE_DIRS and not d.startswith(".")]
            root_path = Path(root)
            for name in files:
                if name in EXCLUDE_FILES or Path(name).suffix in EXCLUDE_SUFFIXES:
                    continue
                src = root_path / name
                arc = Path(app.dir.name, src.relative_to(app.dir))
                zf.write(src, arc.as_posix())

        _write_launchers(zf, app, port)

    buffer.seek(0)
    return buffer


def _write_launchers(zf: zipfile.ZipFile, app: AppInfo, port: int) -> None:
    root = app.dir.name
    needs_ionapi = _references_ionapi(app.dir)

    mac_info = zipfile.ZipInfo(f"{root}/run_mac.command")
    mac_info.external_attr = 0o755 << 16
    zf.writestr(mac_info, _mac_script(app, port))

    zf.writestr(f"{root}/run_windows.bat", _windows_script(app, port))
    zf.writestr(f"{root}/DEPLOY_README.txt", _deploy_readme(app, port, needs_ionapi))


def _mac_script(app: AppInfo, port: int) -> str:
    return f"""#!/bin/bash
set -e
cd "$(dirname "$0")"

if [ ! -d .venv ]; then
    echo "Setting up {app.key} for the first time..."
    python3 -m venv .venv
fi

source .venv/bin/activate
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

echo ""
echo "Starting {app.key} on port {port}..."
echo "Open http://localhost:{port} on this machine,"
echo "or http://<this machine's IP address>:{port} from another machine on the network."
echo "Press Ctrl+C to stop."
echo ""

python3 {app.entry.name} --host 0.0.0.0 --port {port}
"""


def _windows_script(app: AppInfo, port: int) -> str:
    return f"""@echo off
cd /d "%~dp0"

if not exist .venv (
    echo Setting up {app.key} for the first time...
    py -3 -m venv .venv
)

call .venv\\Scripts\\activate.bat
pip install --quiet --upgrade pip
pip install --quiet -r requirements.txt

echo.
echo Starting {app.key} on port {port}...
echo Open http://localhost:{port} on this machine,
echo or http://THIS-MACHINE-IP:{port} from another machine on the network.
echo Press Ctrl+C to stop.
echo.

python {app.entry.name} --host 0.0.0.0 --port {port}
pause
"""


def _deploy_readme(app: AppInfo, port: int, needs_ionapi: bool) -> str:
    lines = [
        f"{app.key} - deployment package",
        "=" * (len(app.key) + 20),
        "",
        f"Built for port {port}.",
        "",
        "Requires Python 3.10 or later on the target machine (python.org/downloads,",
        "or the Microsoft Store on Windows). Nothing else needs to be installed",
        "ahead of time - the launcher creates its own virtual environment and",
        "installs the packages in requirements.txt the first time it runs.",
        "",
        "Mac:      double-click run_mac.command",
        "          (first time only: right-click it and choose Open, since it is",
        "          an unsigned script and Gatekeeper will otherwise block it)",
        "Windows:  double-click run_windows.bat",
        "",
        f"Once it is running, open http://localhost:{port} on that machine, or",
        f"http://<its IP address>:{port} from another machine on the same network.",
        "Stop it with Ctrl+C in the terminal window it opened.",
    ]
    if needs_ionapi:
        lines += [
            "",
            "This app talks to Infor M3 / ION API and looks for .ionapi credential",
            "files in an 'ionapi' folder two levels above its own folder by default.",
            "Either recreate that layout, or start it with --ionapi-dir pointing at",
            "wherever the .ionapi file(s) live, e.g.:",
            f"  python3 {app.entry.name} --host 0.0.0.0 --port {port} --ionapi-dir /path/to/ionapi",
        ]
    return "\n".join(lines) + "\n"
