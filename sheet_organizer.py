import os
import re
import json
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

load_dotenv()

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SERVICE_ACCOUNT_FILE = "sheet_service_account.json"
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]  # write

# === CONFIGURE TARGET SHEET HERE ===
TARGET_SHEET_NAME = "Old Vets"   # change this
TARGET_RANGE = "A1:ZZ"
# ==================================

EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)
ZIP_RE = re.compile(r"\b\d{5}(-\d{4})?\b")

TARGET_HEADERS = [
    "first_name", "last_name", "phone", "email",
    "age", "address", "city", "state", "zip",
    "emailed", "emailed_date",
    "status", "notes", "extras_json"
]

# Header aliases (add your own if needed)
ALIASES = {
    "first_name": [
        "first", "first name", "firstname", "fname", "given name"
    ],
    "last_name": [
        "last", "last name", "lastname", "lname", "surname", "family name"
    ],
    "full_name": [
        "name", "full name", "fullname", "client name", "prospect name"
    ],
    "email": [
        "email", "e-mail", "email address", "mail", "gmail"
    ],
    "phone": [
        "phone", "phone number", "phonenumber", "mobile", "cell", "cell phone", "telephone", "tel"
    ],
    "age": [
        "age"
    ],
    "address": [
        "address", "street", "street address", "address1", "address 1"
    ],
    "city": [
        "city"
    ],
    "state": [
        "state", "st", "province"
    ],
    "zip": [
        "zip", "zipcode", "zip code", "postal", "postal code"
    ],
}

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
        spreadsheetId=SPREADSHEET_ID, range=rng, body={}
    ).execute()

def update_values(svc, sheet_name, start_cell, values):
    rng = f"'{sheet_name}'!{start_cell}"
    svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=rng,
        valueInputOption="RAW",
        body={"values": values}
    ).execute()

def norm_header(h: str) -> str:
    h = (h or "").strip().lower()
    h = re.sub(r"[\s\-_]+", " ", h)
    h = re.sub(r"[^a-z0-9 ]+", "", h)
    return h

def build_header_map(header_row):
    """
    Returns dict: canonical_field -> column_index
    Uses ALIASES to match.
    """
    headers = [norm_header(h) for h in header_row]
    index = {}

    # exact/alias matches
    for field, aliases in ALIASES.items():
        for i, h in enumerate(headers):
            if h in aliases:
                index[field] = i
                break

    # fallback: contains match (ex: "primary email")
    for field, aliases in ALIASES.items():
        if field in index:
            continue
        for i, h in enumerate(headers):
            for a in aliases:
                if a in h:
                    index[field] = i
                    break
            if field in index:
                break

    return index

def normalize_email(x):
    if not x:
        return ""
    m = EMAIL_RE.search(str(x).strip())
    return m.group(0).lower() if m else ""

def normalize_phone(x):
    if not x:
        return ""
    digits = re.sub(r"\D", "", str(x))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else ""

def split_name(full):
    full = (full or "").strip()
    if not full:
        return "", ""
    parts = full.split()
    if len(parts) == 1:
        return parts[0].title(), ""
    return parts[0].title(), " ".join(parts[1:]).title()

def get_cell(row, idx):
    if idx is None:
        return ""
    return row[idx] if idx < len(row) else ""

def extract_zip_anywhere(row):
    for c in row:
        m = ZIP_RE.search(str(c or ""))
        if m:
            return m.group(0)
    return ""

def organize_one_sheet_by_headers(sheet_name: str, range_a1="A1:ZZ"):
    if not SPREADSHEET_ID:
        raise RuntimeError("Missing SPREADSHEET_ID in .env")

    svc = sheets_service()
    rows = get_values(svc, sheet_name, range_a1)
    if not rows:
        print(f"Nothing found in {sheet_name}.")
        return

    header = rows[0]
    header_map = build_header_map(header)

    # Must have at least email or phone to be useful
    if "email" not in header_map and "phone" not in header_map:
        raise RuntimeError(
            f"Couldn't find an Email or Phone column from header row in '{sheet_name}'.\n"
            f"Header row was: {header}\n"
            f"Add a header like 'Email' or 'Phone', or add its name to ALIASES."
        )

    organized = []
    for r_i, row in enumerate(rows[1:], start=2):  # 1-based row numbers (row 2 = first data row)
        extras = {
            "source_sheet": sheet_name,
            "source_row": r_i,
            "original_headers": header,
            "raw_row": row
        }

        first = ""
        last = ""

        # Prefer explicit first/last if present
        if "first_name" in header_map:
            first = str(get_cell(row, header_map.get("first_name"))).strip().title()
        if "last_name" in header_map:
            last = str(get_cell(row, header_map.get("last_name"))).strip().title()

        # If no first/last, try full name column
        if (not first and not last) and "full_name" in header_map:
            full = str(get_cell(row, header_map.get("full_name"))).strip()
            first, last = split_name(full)

        email = normalize_email(get_cell(row, header_map.get("email")))
        phone = normalize_phone(get_cell(row, header_map.get("phone")))

        age = str(get_cell(row, header_map.get("age"))).strip() if "age" in header_map else ""
        address = str(get_cell(row, header_map.get("address"))).strip() if "address" in header_map else ""
        city = str(get_cell(row, header_map.get("city"))).strip() if "city" in header_map else ""
        state = str(get_cell(row, header_map.get("state"))).strip() if "state" in header_map else ""
        zipc = str(get_cell(row, header_map.get("zip"))).strip() if "zip" in header_map else ""

        # If zip wasn't mapped but exists, find it anywhere
        if not zipc:
            zipc = extract_zip_anywhere(row)

        # Tracking columns (email program fills later)
        emailed = ""
        emailed_date = ""
        status = ""
        notes = ""

        organized.append([
            first, last, phone, email,
            age, address, city, state, zipc,
            emailed, emailed_date,
            status, notes,
            json.dumps(extras, ensure_ascii=False)
        ])

    # Rewrite the same sheet cleanly
    clear_range(svc, sheet_name, "A1:ZZ")
    update_values(svc, sheet_name, "A1", [TARGET_HEADERS])
    if organized:
        update_values(svc, sheet_name, "A2", organized)

    print(f"✅ Organized '{sheet_name}' using header mapping. Rows written: {len(organized)}")
    print(f"Detected columns: {header_map}")

if __name__ == "__main__":
    organize_one_sheet_by_headers(TARGET_SHEET_NAME, TARGET_RANGE)
