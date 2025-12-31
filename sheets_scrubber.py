# sheets_scrubber.py
# Scrub leads OUT of non-client tabs by matching phone numbers found in "clients" tabs.
#
# ✅ You hardcode:
#   - SOURCE_OF_TRUTH_SHEETS: sheets whose phone numbers should be protected (never remove from these)
#   - IGNORE_SHEETS: sheets to skip entirely (including MASTER, logs, etc.)
#
# ✅ What it does:
#   1) Collects all phone numbers from SOURCE_OF_TRUTH_SHEETS
#   2) Iterates every other sheet in the spreadsheet (except ignored + source sheets)
#   3) Removes any row whose phone matches a protected phone
#   4) Prints exactly who got removed (name + phone), per sheet
#   5) Writes the cleaned sheet back (header preserved)

import os
import re
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

# =========================
# ✅ EDIT THESE SETTINGS
# =========================
SOURCE_OF_TRUTH_SHEETS = [
    "clients",
    "potential clients",
]

IGNORE_SHEETS = [
    "MASTER",
    "Archive",
    "Do Not Touch",
]

RANGE_READ = "A1:ZZ"   # big enough for all columns
# =========================

load_dotenv()
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SERVICE_ACCOUNT_FILE = "sheet_service_account.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]  # write

PHONE_HEADER_ALIASES = [
    "phone", "phone number", "number", "mobile", "cell", "cell phone",
    "telephone", "tel", "contact number", "contact"
]

NAME_HEADER_ALIASES = {
    "first_name": ["first", "first name", "firstname", "fname", "given name"],
    "last_name": ["last", "last name", "lastname", "lname", "surname", "family name"],
}

def sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds)

def list_sheet_titles(svc):
    meta = svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute()
    return [s["properties"]["title"] for s in meta.get("sheets", [])]

def get_values(svc, sheet_name, a1_range=RANGE_READ):
    rng = f"'{sheet_name}'!{a1_range}"
    resp = svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=rng
    ).execute()
    return resp.get("values", [])

def clear_values(svc, sheet_name, a1_range=RANGE_READ):
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

def normalize_header(h: str) -> str:
    h = (h or "").strip().lower()
    h = re.sub(r"[\s\-_]+", " ", h)
    h = re.sub(r"[^a-z0-9 ]+", "", h)
    return h

def normalize_phone(x: str) -> str:
    if not x:
        return ""
    digits = re.sub(r"\D", "", str(x))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else ""

def format_phone_us(digits: str) -> str:
    if not digits:
        return ""
    digits = re.sub(r"\D", "", digits)
    if len(digits) != 10:
        return digits
    return f"({digits[0:3]}) {digits[3:6]}-{digits[6:10]}"

def find_phone_col(header_row):
    headers = [normalize_header(h) for h in header_row]
    aliases = [normalize_header(a) for a in PHONE_HEADER_ALIASES]

    # exact match
    for i, h in enumerate(headers):
        if h in aliases:
            return i

    # phrase/word match (safe)
    for i, h in enumerate(headers):
        words = set(h.split())
        for a in aliases:
            if " " in a and a in h:
                return i
            if " " not in a and a in words:
                return i

    return None

def find_name_cols(header_row):
    headers = [normalize_header(h) for h in header_row]
    first_aliases = set(normalize_header(a) for a in NAME_HEADER_ALIASES["first_name"])
    last_aliases  = set(normalize_header(a) for a in NAME_HEADER_ALIASES["last_name"])

    first_idx = None
    last_idx = None

    for i, h in enumerate(headers):
        if h in first_aliases and first_idx is None:
            first_idx = i
        if h in last_aliases and last_idx is None:
            last_idx = i

    return first_idx, last_idx

def get_cell(row, idx):
    return row[idx] if idx is not None and idx < len(row) else ""

def collect_protected_phones(svc):
    protected = set()

    for sheet in SOURCE_OF_TRUTH_SHEETS:
        rows = get_values(svc, sheet)
        if not rows or len(rows) < 2:
            print(f"⚠️ Source sheet '{sheet}' empty or missing rows. Skipping.")
            continue

        header = rows[0]
        phone_col = find_phone_col(header)
        if phone_col is None:
            print(f"⚠️ Source sheet '{sheet}' has no phone column. Skipping.")
            continue

        added = 0
        for r in rows[1:]:
            p = normalize_phone(get_cell(r, phone_col))
            if p:
                protected.add(p)
                added += 1

        print(f"✅ Collected {added} phone numbers from '{sheet}'")

    return protected

def scrub_sheet(svc, sheet_name, protected_phones):
    rows = get_values(svc, sheet_name)
    if not rows or len(rows) < 2:
        return 0, 0, []

    header = rows[0]
    data = rows[1:]

    phone_col = find_phone_col(header)
    if phone_col is None:
        print(f"↪ Skipping '{sheet_name}' (no phone column found).")
        return 0, len(data), []

    first_idx, last_idx = find_name_cols(header)

    kept_rows = []
    removed_people = []

    for r in data:
        p = normalize_phone(get_cell(r, phone_col))

        if p and p in protected_phones:
            first = str(get_cell(r, first_idx)).strip() if first_idx is not None else ""
            last  = str(get_cell(r, last_idx)).strip() if last_idx is not None else ""
            name = f"{first} {last}".strip() or "(no name)"
            removed_people.append((name, format_phone_us(p)))
            continue

        kept_rows.append(r)

    removed = len(removed_people)

    if removed > 0:
        clear_values(svc, sheet_name)
        update_values(svc, sheet_name, "A1", [header])
        if kept_rows:
            update_values(svc, sheet_name, "A2", kept_rows)

    return removed, len(kept_rows), removed_people

def main():
    if not SPREADSHEET_ID:
        raise RuntimeError("Missing SPREADSHEET_ID in .env")

    svc = sheets_service()

    # normalize sets for comparisons
    source_lower = {s.strip().lower() for s in SOURCE_OF_TRUTH_SHEETS}
    ignore_lower = {s.strip().lower() for s in IGNORE_SHEETS}

    print("🔎 Building protected list from source sheets...")
    protected = collect_protected_phones(svc)
    print(f"\n✅ Total protected phones: {len(protected)}\n")

    if not protected:
        print("No protected phones found. Nothing to scrub.")
        return

    all_sheets = list_sheet_titles(svc)

    removed_total = 0
    for sh in all_sheets:
        sh_lower = sh.strip().lower()

        if sh_lower in source_lower:
            print(f"🔒 Not touching source sheet '{sh}'")
            continue

        if sh_lower in ignore_lower:
            print(f"⏭️ Ignoring '{sh}'")
            continue

        removed, kept, removed_people = scrub_sheet(svc, sh, protected)
        removed_total += removed

        if removed > 0:
            print(f"🧹 Scrubbed '{sh}':")
            for name, phone_fmt in removed_people:
                print(f"   ❌ {name} | {phone_fmt}")
            print(f"   → removed {removed}, kept {kept}")
        else:
            print(f"✅ No matches in '{sh}' (kept {kept})")

    print(f"\n✅ Done. Total removed across sheets: {removed_total}")

if __name__ == "__main__":
    main()
