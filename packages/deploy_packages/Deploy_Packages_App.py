"""
Deploy_Packages_App - Flask front end for building a standalone,
downloadable copy of one of the other Doppio apps under packages/.

Pick an app and a port, and it zips up that app's own folder - source,
templates, requirements.txt - together with a run_mac.command and a
run_windows.bat that create a virtual environment, install requirements.txt,
and start the app on the port you chose. Nothing is installed here; that all
happens the first time the launcher runs on the target machine.

Data folders (input/, output/, gnupg_home/), virtual environments, and
version-control noise (__pycache__, .venv, .git, .vscode) are left out of
the package - the target machine builds and fills those itself.
"""
from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file

from Deploy_Packages_Build import build_zip, discover_apps, get_app

BASE_DIR = Path(__file__).parent.resolve()
app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
log = logging.getLogger("deploy_packages")


@app.route("/")
def index():
    return render_template("Deploy_Packages_Index.html")


@app.route("/api/apps")
def api_apps():
    return jsonify([{
        "key": a.key,
        "description": a.description,
        "default_port": a.default_port,
    } for a in discover_apps()])


@app.route("/api/build", methods=["POST"])
def api_build():
    data = request.get_json(force=True, silent=True) or {}

    app_info = get_app(str(data.get("app") or ""))
    if app_info is None:
        return jsonify(error="Unknown app."), 404

    try:
        port = int(data.get("port"))
    except (TypeError, ValueError):
        return jsonify(error="Port must be a number."), 400
    if not (1 <= port <= 65535):
        return jsonify(error="Port must be between 1 and 65535."), 400

    buffer = build_zip(app_info, port)
    filename = f"{app_info.key}_{port}.zip"
    log.info("Built %s for port %s", app_info.key, port)
    return send_file(buffer, mimetype="application/zip",
                      as_attachment=True, download_name=filename)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        description="Build a deployable package for one of the packages/ apps.")
    ap.add_argument("--host", default="127.0.0.1")
    # 5057-5059 are m3_security, adp_concur and mig_sync; 5062 is
    # sheet_security; 5063 is ERP_Concur; 5065 is launcher; 8799 is pgp_keys.
    ap.add_argument("--port", type=int, default=5064)
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
