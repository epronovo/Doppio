from pathlib import Path
from InforMI import CONFIG, prompt_for_company_division, select_ionapi_file
from UserDefaults import load_user_defaults, save_user_defaults

def initialize_config_from_defaults():
    user_defaults = load_user_defaults()

    # IONAPI file
    ionapi_dir = Path(__file__).parent / "ionapi"
    CONFIG["tenant"] = user_defaults.get("tenant") or select_ionapi_file(ionapi_dir)
    user_defaults["tenant"] = CONFIG["tenant"]

    # Company / Division
    CONFIG["company"] = user_defaults.get("company", "")
    CONFIG["division"] = user_defaults.get("division", "")
    if not CONFIG["company"] or not CONFIG["division"]:
        prompt_for_company_division()
        user_defaults["company"] = CONFIG["company"]
        user_defaults["division"] = CONFIG["division"]

    # Batch size
    CONFIG["batch_size"] = int(user_defaults.get("batch_size", CONFIG.get("batch_size", 500)))

    # Ensure it's saved as an int, not a string
    user_defaults["batch_size"] = CONFIG["batch_size"]

    save_user_defaults(user_defaults)