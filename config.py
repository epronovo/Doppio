# config.py
import os
from pathlib import Path
from dotenv import load_dotenv
from UserDefaults import load_user_defaults

# Load environment variables
load_dotenv(dotenv_path=Path(__file__).parent / ".env")

# Paths
BASE_DIR = Path(__file__).parent.resolve()
EVS100_DIR = BASE_DIR / "evs100"

def get_sqlite_db_path():
    defaults = load_user_defaults()
    db_path = Path.home() / "sqlite" / defaults.get("local_db_name", "asr_uat2.db")
    db_path.parent.mkdir(parents=True, exist_ok=True)
    return db_path

# Credentials and connection string
DB2_UID = os.getenv("DB2_UID")
DB2_PWD = os.getenv("DB2_PWD")

DB2_CONN_STR = (
    f"DSN=Raymond DB;"
    f"UID={DB2_UID};"
    f"PWD={DB2_PWD};"
)