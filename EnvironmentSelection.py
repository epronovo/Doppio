import os
import json
from pathlib import Path
from datetime import datetime

# ---------- Paths ----------
DEFAULTS_FILE = Path(__file__).parent / "user_defaults.json"
IONAPI_FOLDER = "./ionapi"

# ---------- Load / Save user defaults ----------
def load_user_defaults():
    if DEFAULTS_FILE.exists():
        with open(DEFAULTS_FILE, "r") as f:
            return json.load(f)
    return {}

def save_user_defaults(data):
    with open(DEFAULTS_FILE, "w") as f:
        json.dump(data, f, indent=2)
    print("✅ user_defaults.json updated successfully.")

# ---------- IONAPI Selection ----------
def list_ionapi_files(folder_path):
    """Return a sorted list of .ionapi files in the given folder."""
    return sorted(
        (f for f in os.listdir(folder_path) if f.endswith(".ionapi")),
        key=str.lower
    )

def select_ionapi_file(folder_path):
    """Prompt user to select an .ionapi file from the folder."""
    files = list_ionapi_files(folder_path)

    if not files:
        print("⚠️ No .ionapi files found.")
        return None

    print("🔹 Select an ION API file to use:\n")
    for i, file in enumerate(files, start=1):
        print(f"{i}. {file}")

    while True:
        try:
            choice = int(input(f"\nEnter your choice (1-{len(files)}): "))
            if 1 <= choice <= len(files):
                selected_file = os.path.join(folder_path, files[choice - 1])
                print(f"\n✅ You selected: {files[choice - 1]}")
                return selected_file
            print("❌ Invalid choice. Try again.")
        except ValueError:
            print("❌ Please enter a number.")

# ---------- Update user defaults from IONAPI ----------
def update_user_defaults_from_ionapi(ionapi_path):
    with open(ionapi_path, "r") as f:
        ionapi_config = json.load(f)

    # Load existing defaults
    defaults = load_user_defaults()

    # Update values
    defaults["company"] = ionapi_config.get("company", defaults.get("company", ""))
    defaults["division"] = ionapi_config.get("division", defaults.get("division", ""))
    defaults["local_db_name"] = ionapi_config.get("database", defaults.get("local_db_name", ""))
    defaults["ionapi_file"] = os.path.basename(ionapi_path)
    defaults["last_ionapi_prompt_time"] = datetime.now().isoformat()

    # Save updated defaults
    save_user_defaults(defaults)

# ---------- Main ----------
if __name__ == "__main__":
    selected_ionapi_path = select_ionapi_file(IONAPI_FOLDER)
    if selected_ionapi_path:
        update_user_defaults_from_ionapi(selected_ionapi_path)