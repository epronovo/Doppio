import gspread
from google.oauth2.service_account import Credentials
import sys

# --- CONFIGURATION ---
SERVICE_ACCOUNT_FILE = 'creds.json'
SPREADSHEET_ID = '1b-SajPMbrncuhlgRHI-9GuWEdsAGlgjsJSUZqagXEpk' 
SOURCE_TAB_NAME = 'Distribution Config' 
START_ROW = 5               
# ---------------------

def create_program_tabs():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    
    try:
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
        client = gspread.authorize(creds)
        
        # Open the spreadsheet
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        source_worksheet = spreadsheet.worksheet(SOURCE_TAB_NAME)
        print(f"✅ Successfully connected to: {spreadsheet.title}")
        
    except FileNotFoundError:
        print(f"❌ Error: The file '{SERVICE_ACCOUNT_FILE}' was not found in this folder.")
        return
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ Error: Spreadsheet not found. Check the ID and ensure you SHARED the sheet with the service account email.")
        return
    except gspread.exceptions.WorksheetNotFound:
        print(f"❌ Error: The tab '{SOURCE_TAB_NAME}' does not exist in this spreadsheet.")
        return
    except Exception as e:
        print(f"❌ Detailed Connection Error: {e}")
        return

    # Process the tabs
    existing_tabs = {ws.title for ws in spreadsheet.worksheets()}
    codes_column = source_worksheet.col_values(3) # Column C

    print("Checking program codes...")
    for code in codes_column[START_ROW-1:]:
        code = code.strip()
        
        if not code:
            print("🏁 Reached an empty cell. Stopping.")
            break
            
        if code in existing_tabs:
            print(f"🟡 Skipping: '{code}' (already exists)")
        else:
            try:
                spreadsheet.add_worksheet(title=code, rows="100", cols="20")
                print(f"🟢 Created: {code}")
                existing_tabs.add(code) 
            except Exception as e:
                print(f"🔴 Failed to create '{code}': {e}")

if __name__ == "__main__":
    create_program_tabs()