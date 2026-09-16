"""
Launcher_App - Flask front end for starting and stopping the other Doppio
apps under packages/ individually, instead of a terminal per app.

Every app here already runs standalone (`python3 X_App.py --port N`, or
double-clicking a deploy_packages zip's run script); this adds nothing to how
any of them work, just one page listing all of them with a Start and a Stop
button each, and what is actually running right now. See Launcher_Process for
how an app is discovered, launched, and tracked.

Every route returns JSON; the page itself is templates/Launcher_Index.html.
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from flask import Flask, jsonify, render_template, request

import Launcher_Process as proc

BASE_DIR = Path(__file__).parent.resolve()
app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
log = logging.getLogger("Launcher_App")


def _body() -> dict:
    return request.get_json(force=True, silent=True) or {}


def _port_from(body: dict) -> tuple[int | None, str | None]:
    """(port, error) - None, None means "use the app's own default"."""
    raw = body.get("port")
    if raw in (None, ""):
        return None, None
    try:
        port = int(raw)
    except (TypeError, ValueError):
        return None, "Port must be a number."
    if not (1 <= port <= 65535):
        return None, "Port must be between 1 and 65535."
    return port, None


@app.route("/")
def index():
    return render_template("Launcher_Index.html")


@app.route("/api/apps")
def api_apps():
    return jsonify(apps=[proc.status(a) for a in proc.discover_apps()])


@app.route("/api/apps/<key>/start", methods=["POST"])
def api_start(key: str):
    a = proc.get_app(key)
    if a is None:
        return jsonify(ok=False, message=f"Unknown app {key!r}."), 404
    port, error = _port_from(_body())
    if error:
        return jsonify(ok=False, message=error), 400
    return jsonify(proc.start(a, port))


@app.route("/api/apps/<key>/stop", methods=["POST"])
def api_stop(key: str):
    a = proc.get_app(key)
    if a is None:
        return jsonify(ok=False, message=f"Unknown app {key!r}."), 404
    return jsonify(proc.stop(a))


@app.route("/api/apps/<key>/restart", methods=["POST"])
def api_restart(key: str):
    a = proc.get_app(key)
    if a is None:
        return jsonify(ok=False, message=f"Unknown app {key!r}."), 404
    port, error = _port_from(_body())
    if error:
        return jsonify(ok=False, message=error), 400
    proc.stop(a)
    return jsonify(proc.start(a, port))


@app.route("/api/apps/<key>/log")
def api_log(key: str):
    a = proc.get_app(key)
    if a is None:
        return jsonify(ok=False, message=f"Unknown app {key!r}."), 404
    return jsonify(ok=True, key=key, log=proc.tail_log(a))


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Start and stop the other packages/ apps individually.")
    ap.add_argument("--host", default="127.0.0.1")
    # 5057-5059 are m3_security, adp_concur and mig_sync; 5060/5061 are the
    # SIP ports Chrome blocks; 5062 is sheet_security; 5063 is ERP_Concur;
    # 5064 is deploy_packages; 8799 is pgp_keys.
    ap.add_argument("--port", type=int, default=5065)
    ap.add_argument("--debug", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(
        level=logging.DEBUG if args.debug else logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout)])

    log.info("Open     : http://%s:%s", args.host, args.port)
    app.run(host=args.host, port=args.port, debug=args.debug, threaded=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
