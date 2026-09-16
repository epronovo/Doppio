"""
Launcher_Process - discover the other Doppio apps under packages/ and start,
stop, and track them.

An app is any packages/ folder with a *_App.py in it - the same convention
Deploy_Packages_Build.discover_apps() uses to find what it can zip up. That
module is not imported from here on purpose: every package in this repo
stands alone, run as `python3 X_App.py`, and this one is no exception - it
duplicates the handful of lines that find an app's docstring and its --port
default rather than reach into deploy_packages' folder to borrow them.

Each app is started as its own subprocess, in its own directory, on the port
its own argparse default already claims - never a second guess at what port
is free, because that default is already the one place every app (and the
zips deploy_packages builds) agree on a port with nothing else in this repo.

State lives entirely in run/ - a PID file and a log file per app - not in
this process's own memory, so restarting the launcher itself never loses
track of what it started: the next request re-checks every PID against what
is actually running before trusting it.
"""
from __future__ import annotations

import ast
import os
import re
import signal
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
PACKAGES_DIR = BASE_DIR.parent
SELF_NAME = BASE_DIR.name
RUN_DIR = BASE_DIR / "run"

_PORT_RE = re.compile(r'--port["\']\s*,\s*type\s*=\s*int\s*,\s*default\s*=\s*(\d+)')


@dataclass
class AppInfo:
    key: str
    dir: Path
    entry: Path
    description: str
    default_port: int


def discover_apps() -> list[AppInfo]:
    """Every packages/ folder with a *_App.py, sorted by name, this one left out."""
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
    """Look up one app by its packages/ folder name - never trust a key from
    a request without this, the same rule Deploy_Packages_Build follows."""
    for a in discover_apps():
        if a.key == key:
            return a
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


# -------------------------------------------------------------- run state

# Popen objects for whatever this launcher process itself started, kept only
# so their exit can be reaped with .poll()/.wait() - without that, a child
# this same process spawned and later killed sits as a <defunct> zombie
# until this process exits, since start_new_session=True detaches the child
# from our process GROUP but not from our PID as its parent. Lost on restart,
# which is fine: an app started by an earlier launcher process is already
# reparented to init by the time this one starts, and init reaps for free.
_procs: dict[str, subprocess.Popen] = {}


def _reap(key: str) -> None:
    p = _procs.get(key)
    if p is not None and p.poll() is not None:
        del _procs[key]


def _pid_file(key: str) -> Path:
    return RUN_DIR / f"{key}.pid"


def _log_file(key: str) -> Path:
    return RUN_DIR / f"{key}.log"


def _alive(pid: int, entry_name: str) -> bool:
    """
    True if `pid` exists and looks like it is still running this app.

    os.kill(pid, 0) alone would believe a PID that has since been recycled by
    an unrelated process - unlikely between one status check and the next,
    but cheap to rule out with `ps`, which both machines this launcher runs
    on (macOS and Linux) carry.
    """
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except OSError:
        pass                        # exists but not ours to signal - alive
    try:
        out = subprocess.run(["ps", "-p", str(pid), "-o", "command="],
                             capture_output=True, text=True, timeout=2)
        return entry_name in out.stdout
    except (OSError, subprocess.TimeoutExpired):
        return True                 # can't confirm the command line; os.kill said alive


def status(app: AppInfo) -> dict:
    """
    What is actually running, not what run/ last said - a stale PID file
    (the app crashed, or was killed from outside the launcher) is cleaned up
    here rather than left to lie to the next request.
    """
    _reap(app.key)
    pid_file = _pid_file(app.key)
    pid = None
    running = False
    if pid_file.exists():
        try:
            pid = int(pid_file.read_text().strip())
        except ValueError:
            pid = None
        if pid is not None:
            running = _alive(pid, app.entry.name)
    if not running:
        pid_file.unlink(missing_ok=True)
        pid = None
    return {"key": app.key, "description": app.description,
            "port": app.default_port, "running": running, "pid": pid}


def start(app: AppInfo, port: int | None = None) -> dict:
    st = status(app)
    if st["running"]:
        return {"ok": False, "message": f"{app.key} is already running "
                 f"(pid {st['pid']}).", **st}
    RUN_DIR.mkdir(parents=True, exist_ok=True)
    port = port or app.default_port
    log = open(_log_file(app.key), "ab")
    log.write(f"\n--- launched {time.strftime('%Y-%m-%d %H:%M:%S')} "
             f"on port {port} ---\n".encode())
    log.flush()
    proc = subprocess.Popen(
        [sys.executable, str(app.entry), "--port", str(port)],
        cwd=app.dir, stdout=log, stderr=log, start_new_session=True)
    _procs[app.key] = proc
    _pid_file(app.key).write_text(str(proc.pid))
    # Flask logs its "Open" line almost immediately; a genuine startup
    # failure (bad import, port already taken by something else) exits
    # in well under this, so a short wait is enough to tell the two apart.
    time.sleep(0.8)
    st = status(app)
    if not st["running"]:
        return {"ok": False, "message": f"{app.key} exited immediately - "
                 f"check its log.", **st}
    return {"ok": True, "message": f"Started {app.key} on port {port} "
             f"(pid {st['pid']}).", **st}


def stop(app: AppInfo, grace: float = 3.0) -> dict:
    """SIGTERM, then SIGKILL after `grace` seconds if it is still standing -
    the same two-step restart.py already uses for ADP_Concur_App."""
    st = status(app)
    if not st["running"]:
        return {"ok": True, "message": f"{app.key} was not running.", **st}
    pid = st["pid"]
    try:
        os.kill(pid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    deadline = time.time() + grace
    while time.time() < deadline:
        if not _alive(pid, app.entry.name):
            break
        time.sleep(0.2)
    else:
        try:
            os.kill(pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
    tracked = _procs.get(app.key)
    if tracked is not None:
        try:
            tracked.wait(timeout=2)          # reap it - see _procs' own comment
        except subprocess.TimeoutExpired:
            pass
        _procs.pop(app.key, None)
    _pid_file(app.key).unlink(missing_ok=True)
    return {"ok": True, "message": f"Stopped {app.key}.", "key": app.key,
            "description": app.description, "port": app.default_port,
            "running": False, "pid": None}


def tail_log(app: AppInfo, n: int = 200) -> str:
    p = _log_file(app.key)
    if not p.exists():
        return ""
    lines = p.read_text(encoding="utf-8", errors="ignore").splitlines()
    return "\n".join(lines[-n:])
