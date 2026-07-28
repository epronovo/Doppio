import json
from pathlib import Path

DEFAULTS_FILE = Path(__file__).parent / "user_defaults.json"

def load_user_defaults():
    if DEFAULTS_FILE.exists():
        with open(DEFAULTS_FILE, "r") as f:
            return json.load(f)
    return {}

def save_user_defaults(data):
    with open(DEFAULTS_FILE, "w") as f:
        json.dump(data, f, indent=2)