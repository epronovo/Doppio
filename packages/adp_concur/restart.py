#!/usr/bin/env python3
"""
Stop any running ADP_Concur_App and start a fresh one.

Written as a file rather than a shell one-liner on purpose. `pkill -f
ADP_Concur_App.py` matches the shell command that contains the pattern as
well as the server, so it kills the caller - twice now. This matches the
*arguments* instead: argv[0] has to be a python, and one of the later
arguments has to be the script itself. Its own process, and whoever started
it, are skipped whatever their command line says.
"""

import os
import signal
import subprocess
import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
SCRIPT = "ADP_Concur_App.py"
MINE = {os.getpid(), os.getppid()}


def running() -> list[int]:
    out = []
    for entry in os.listdir("/proc"):
        if not entry.isdigit() or int(entry) in MINE:
            continue
        try:
            argv = Path(f"/proc/{entry}/cmdline").read_bytes().decode().split("\0")
        except OSError:
            continue
        argv = [a for a in argv if a]
        if len(argv) < 2 or "python" not in Path(argv[0]).name:
            continue
        if any(a.endswith(SCRIPT) for a in argv[1:]):
            out.append(int(entry))
    return out


def main() -> int:
    for pid in running():
        print(f"stopping {pid}")
        try:
            os.kill(pid, signal.SIGTERM)
        except ProcessLookupError:
            pass
    time.sleep(1.5)
    for pid in running():                       # anything that ignored SIGTERM
        os.kill(pid, signal.SIGKILL)

    args = [sys.executable, str(HERE / SCRIPT)] + sys.argv[1:]
    log = open(HERE / "app.log", "ab")
    subprocess.Popen(args, cwd=HERE, stdout=log, stderr=log,
                     start_new_session=True)
    time.sleep(4)
    pids = running()
    print(f"started {pids or 'nothing - check app.log'}")
    return 0 if pids else 1


if __name__ == "__main__":
    raise SystemExit(main())
