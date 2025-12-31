# lead_finder.py
# Search a Google Spreadsheet for a lead by FIRST + LAST name
# Prints:
#   - which sheet/tab the lead is in
#   - the full row data (as key:value using the header row)
#
# ✅ Hardcode sheets to ignore at the top
# ✅ Works even if columns are different per sheet (uses each tab's header row)

import os
import re
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

# =========================
# ✅ EDIT THESE SETTINGS
# =========================
SHEETS_TO_IGNORE = [
    "Clients",
    "Potential Clients",
]

TARGET_RANGE = "A1:ZZ"

# You can also hardcode the name here if you want:
SEARCH_FIRST_NAME = ""   # e.g. "John"
SEARCH_LAST_NAME  = ""   # e.g. "Smith"
# =========================

load_dotenv()
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SERVICE_ACCOUNT_FILE = "sheet_service_account.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# Header aliases
ALIASES = {
    "first_name": ["first", "first name", "firstname", "fname", "given name"],
    "last_name": ["last", "last name", "lastname", "lname", "surname", "family name"]
#    "full_name": ["name", "full name", "fullname", "client name", "prospect name"],
}

def sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)

def list_sheet_titles(svc):
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return [s["properties"]["title"] for s in meta.get("sheets", [])]

def get_values(svc, sheet_name, a1_range=TARGET_RANGE):
    rng = f"'{sheet_name}'!{a1_range}"
    resp = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=rng
    ).execute()
    return resp.get("values", [])

def normalize_header(h: str) -> str:
    h = (h or "").strip().lower()
    h = re.sub(r"[\s\-_]+", " ", h)
    h = re.sub(r"[^a-z0-9 ]+", "", h)
    return h

def normalize_name(s: str) -> str:
    s = (s or "").strip().lower()
    s = re.sub(r"[^a-z0-9 ]+", "", s)
    s = re.sub(r"\s+", " ", s)
    return s

def build_header_map(header_row):
    headers = [normalize_header(h) for h in header_row]
    idx = {}

    # exact match
    for field, aliases in ALIASES.items():
        aliases_norm = [normalize_header(a) for a in aliases]
        for i, h in enumerate(headers):
            if h in aliases_norm:
                idx[field] = i
                break

    # contains / phrase match
    for field, aliases in ALIASES.items():
        if field in idx:
            continue
        aliases_norm = [normalize_header(a) for a in aliases]
        for i, h in enumerate(headers):
            for a in aliases_norm:
                if " " in a and a in h:
                    idx[field] = i
                    break
            if field in idx:
                break

    return idx

def get_cell(row, idx):
    return row[idx] if idx is not None and idx < len(row) else ""

def row_as_dict(header, row):
    out = {}
    for i, h in enumerate(header):
        key = (h or f"col_{i+1}").strip()
        out[key] = row[i] if i < len(row) else ""
    return out

def split_full_name(full: str):
    full = normalize_name(full)
    if not full:
        return "", ""
    parts = full.split()
    if len(parts) == 1:
        return parts[0], ""
    return parts[0], " ".join(parts[1:])

def matches_name(row, header_map, first_query, last_query):
    # Prefer first/last columns if present
    f_idx = header_map.get("first_name")
    l_idx = header_map.get("last_name")
    n_idx = header_map.get("full_name")

    first_val = normalize_name(get_cell(row, f_idx)) if f_idx is not None else ""
    last_val  = normalize_name(get_cell(row, l_idx)) if l_idx is not None else ""

    # If we don't have both, try full name
    if (not first_val or not last_val) and n_idx is not None:
        f2, l2 = split_full_name(get_cell(row, n_idx))
        first_val = first_val or f2
        last_val = last_val or l2

    return first_val == first_query and last_val == last_query

def find_lead(first_name, last_name):
    if not SPREADSHEET_ID:
        raise RuntimeError("Missing SPREADSHEET_ID in .env")

    first_q = normalize_name(first_name)
    last_q = normalize_name(last_name)

    if not first_q or not last_q:
        raise ValueError("Provide BOTH first and last name.")

    svc = sheets_service()
    ignore = {s.strip().lower() for s in SHEETS_TO_IGNORE}
    sheets = list_sheet_titles(svc)

    hits = []

    for sh in sheets:
        if sh.strip().lower() in ignore:
            continue

        rows = get_values(svc, sh)
        if not rows or len(rows) < 2:
            continue

        header = rows[0]
        data = rows[1:]
        header_map = build_header_map(header)

        # If this sheet has no relevant name fields, skip it
        if not any(k in header_map for k in ("first_name", "last_name", "full_name")):
            continue

        for row_number_1based, row in enumerate(data, start=2):
            if matches_name(row, header_map, first_q, last_q):
                hits.append((sh, row_number_1based, header, row))

    return hits

def print_hits(hits, first_name, last_name):
    if not hits:
        print(f"❌ No matches found for: {first_name} {last_name}")
        return

    print(f"✅ Found {len(hits)} match(es) for: {first_name} {last_name}\n")

    for i, (sheet, rownum, header, row) in enumerate(hits, start=1):
        print(f"--- Match #{i} ---")
        print(f"📄 Sheet: {sheet}")
        print(f"📍 Row: {rownum}")

        d = row_as_dict(header, row)
        # Pretty print key/value
        for k, v in d.items():
            if str(v).strip() != "":
                print(f"  - {k}: {v}")
        print()

if __name__ == "__main__":
    # Option A: hardcode at top
    if SEARCH_FIRST_NAME and SEARCH_LAST_NAME:
        f, l = SEARCH_FIRST_NAME, SEARCH_LAST_NAME
    else:
        # Option B: prompt when you run (press ▶)
        f = input("First name: ").strip()
        l = input("Last name: ").strip()

    hits = find_lead(f, l)
    print_hits(hits, f, l)
