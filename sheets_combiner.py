# sheets_combiner.py
# Combines specific tabs from ONE Google Spreadsheet into a single "MASTER" tab.
# - You hardcode which tabs to combine + output column order at the top
# - Each source tab must have a header row (row 1)
# - Missing columns in a source tab are ignored (left blank in output)
# - Extra columns in a source tab are ignored unless you include them in OUTPUT_HEADERS

import os
import re
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

# =========================
# ✅ EDIT THESE SETTINGS
# =========================
SOURCE_SHEETS = [
    "TTC 12/30",
    "NEW TTC",
    "TTC HIGH INTENT",
]

OUTPUT_SHEET_NAME = "Combined TTC 12/30"

# ✅ Output column order (edit this anytime)
OUTPUT_HEADERS = [
    "first_name",
    "last_name",
    "number",
    "email",
    "notes",
    "beneficiary",
    "state",
    "hobby",
    "coverage",
    "fe"
]
# =========================

load_dotenv()
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SERVICE_ACCOUNT_FILE = "sheet_service_account.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]  # write

def sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)

def norm_header(h: str) -> str:
    h = (h or "").strip().lower()
    h = re.sub(r"[\s\-_]+", "_", h)           # spaces/dashes -> _
    h = re.sub(r"[^a-z0-9_]+", "", h)         # remove weird chars
    return h

def read_sheet_values(svc, sheet_name: str, a1_range="A1:ZZ"):
    rng = f"'{sheet_name}'!{a1_range}"
    resp = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=rng
    ).execute()
    return resp.get("values", [])

def clear_sheet(svc, sheet_name: str):
    svc.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!A1:ZZ",
        body={}
    ).execute()

def write_values(svc, sheet_name: str, start_cell: str, values):
    svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{sheet_name}'!{start_cell}",
        valueInputOption="RAW",
        body={"values": values}
    ).execute()

def get_tab_names(svc):
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return [s["properties"]["title"] for s in meta.get("sheets", [])]

def ensure_output_tab_exists(svc, tab_name: str):
    tabs = get_tab_names(svc)
    if tab_name in tabs:
        return
    svc.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": [{"addSheet": {"properties": {"title": tab_name}}}]}
    ).execute()

def build_header_index_map(header_row):
    """
    Returns: {normalized_header: index}
    """
    return {norm_header(h): i for i, h in enumerate(header_row)}

def get_cell(row, idx):
    return row[idx] if idx is not None and idx < len(row) else ""

def combine():
    if not SPREADSHEET_ID:
        raise RuntimeError("Missing SPREADSHEET_ID in .env")

    svc = sheets_service()

    all_rows_out = []
    output_headers_norm = [norm_header(h) for h in OUTPUT_HEADERS]

    for tab in SOURCE_SHEETS:
        rows = read_sheet_values(svc, tab, "A1:ZZ")
        if not rows or len(rows) < 2:
            print(f"⚠️ Skipping '{tab}' (empty or no data).")
            continue

        header = rows[0]
        header_map = build_header_index_map(header)

        # Build each output row in OUTPUT_HEADERS order
        for r in rows[1:]:
            out_row = []
            for out_h_norm in output_headers_norm:
                idx = header_map.get(out_h_norm)
                out_row.append(get_cell(r, idx))
            all_rows_out.append(out_row)

        print(f"✅ Added {len(rows)-1} rows from '{tab}'")

    ensure_output_tab_exists(svc, OUTPUT_SHEET_NAME)

    # Write output
    clear_sheet(svc, OUTPUT_SHEET_NAME)
    write_values(svc, OUTPUT_SHEET_NAME, "A1", [OUTPUT_HEADERS])
    if all_rows_out:
        write_values(svc, OUTPUT_SHEET_NAME, "A2", all_rows_out)

    print(f"\n✅ Done. Wrote {len(all_rows_out)} combined rows into '{OUTPUT_SHEET_NAME}'.")

if __name__ == "__main__":
    combine()
