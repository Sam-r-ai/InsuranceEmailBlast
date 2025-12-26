import os
import re
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ----------------------------
# Config
# ----------------------------
load_dotenv()

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SERVICE_ACCOUNT_FILE = "sheet_service_account.json"
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# 👇 CHANGE THIS AND PRESS ▶ RUN
TARGET_SHEET_NAME = "NEW TTC"
TARGET_RANGE = "A1:ZZ"

# State header aliases (add more if needed)
STATE_ALIASES = [
    "state", "st", "province", "region"
]

# ----------------------------
# Sheets API
# ----------------------------
def sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SHEETS_SCOPES
    )
    return build("sheets", "v4", credentials=creds)

def get_values(svc, sheet_name, a1_range):
    rng = f"'{sheet_name}'!{a1_range}"
    resp = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=rng
    ).execute()
    return resp.get("values", [])

def clear_range(svc, sheet_name, a1_range):
    rng = f"'{sheet_name}'!{a1_range}"
    svc.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=rng,
        body={}
    ).execute()

def update_values(svc, sheet_name, start_cell, values):
    rng = f"'{sheet_name}'!{start_cell}"
    svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=rng,
        valueInputOption="RAW",
        body={"values": values}
    ).execute()

# ----------------------------
# Helpers
# ----------------------------
def normalize_header(h):
    h = (h or "").strip().lower()
    h = re.sub(r"[\s\-_]+", " ", h)
    h = re.sub(r"[^a-z ]+", "", h)
    return h

def find_state_column(header_row):
    headers = [normalize_header(h) for h in header_row]
    for i, h in enumerate(headers):
        if h in STATE_ALIASES:
            return i
        for alias in STATE_ALIASES:
            if alias in h:
                return i
    return None

def get_cell(row, idx):
    return row[idx] if idx is not None and idx < len(row) else ""

# ----------------------------
# Core logic
# ----------------------------
def sort_sheet_by_state():
    if not SPREADSHEET_ID:
        raise RuntimeError("Missing SPREADSHEET_ID in .env")

    svc = sheets_service()
    rows = get_values(svc, TARGET_SHEET_NAME, TARGET_RANGE)

    if not rows or len(rows) < 2:
        print("Nothing to sort.")
        return

    header = rows[0]
    data_rows = rows[1:]

    state_col = find_state_column(header)

    if state_col is None:
        raise RuntimeError(
            f"❌ Could not find a State column.\n"
            f"Header row was: {header}\n"
            f"Accepted names: {STATE_ALIASES}"
        )

    # Sort rows by state (case-insensitive, blanks last)
    def sort_key(row):
        val = str(get_cell(row, state_col)).strip().upper()
        return (val == "", val)

    sorted_rows = sorted(data_rows, key=sort_key)

    # Rewrite sheet
    clear_range(svc, TARGET_SHEET_NAME, "A1:ZZ")
    update_values(svc, TARGET_SHEET_NAME, "A1", [header])
    update_values(svc, TARGET_SHEET_NAME, "A2", sorted_rows)

    print(
        f"✅ Sheet '{TARGET_SHEET_NAME}' sorted alphabetically by STATE "
        f"(column {state_col + 1})."
    )

# ----------------------------
# Run
# ----------------------------
if __name__ == "__main__":
    sort_sheet_by_state()
