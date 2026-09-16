# Deploy Packages

Flask front end for building a standalone, downloadable copy of one of the
other Doppio apps under `packages/`.

## What it does

`Deploy_Packages_Build.discover_apps()` scans every folder in `packages/`
(except itself) for a `*_App.py` - the Flask entry point every app in this
repo is built around. A folder without one (e.g. `etl_datalake`, which is a
plain script) does not show up in the picker.

Pick an app and a port on the page. `POST /api/build` zips up that app's own
folder as-is - its `.py` files, `templates/`, `requirements.txt`, any JSON
config sitting next to them - and adds:

- `run_mac.command` - creates `.venv`, installs `requirements.txt`, then runs
  `python3 <App>.py --host 0.0.0.0 --port <PORT>`
- `run_windows.bat` - the same, for `py`/`Scripts\activate.bat`
- `DEPLOY_README.txt` - what to double-click, and a note about `--ionapi-dir`
  for apps that talk to Infor M3 / ION API

Left out of the zip: `.venv/`, `__pycache__/`, `.git/`, `.vscode/`,
`input/`, `output/`, `gnupg_home/`, and `.DS_Store`/`*.pyc`. Those are either
version-control noise or data the app rebuilds itself on first run (its
`input`/`output` folders, its SQLite database under `~/sqlite/`) - nothing
in them is needed to run the app somewhere else, and some of it (dropped
import files, GnuPG keyrings) is exactly what should not leave this machine.

The build happens in memory (`io.BytesIO` + `zipfile`) and is streamed back
as the response; nothing is written to disk on this end.

## Routes

- `GET /` - the picker page
- `GET /api/apps` - `[{key, description, default_port}]` for every
  deployable app
- `POST /api/build` - `{app, port}` in, the zip out as an attachment

## Running it

```
python Deploy_Packages_App.py --port 5064
```

Port 5064; see the comment above `main()` for the other packages/ apps'
ports.
