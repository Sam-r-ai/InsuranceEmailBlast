import os
import re
from datetime import datetime, timezone
from dotenv import load_dotenv
from google.oauth2 import service_account
from googleapiclient.discovery import build

# ----------------------------
# Config
# ----------------------------
load_dotenv()

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SERVICE_ACCOUNT_FILE = "sheet_service_account.json"

# WRITE ENABLED (because we are reorganizing)
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

MASTER_SHEET = "Master"

MASTER_HEADERS = [
    "first_name", "last_name", "email", "phone",
    "age", "address", "city", "state", "zip",
    "source_sheet", "source_row",
    "status", "sent_at", "notes"
]

# Add every raw tab you want normalized here:
# format: (sheet_name, a1_range)
SOURCE_SHEETS = [
    ("Sheet4", "A1:Z"),
    # ("Old Vets", "A1:Z"),
    # ("Bronze_Silver", "A1:Z"),
    # ("NEW TTC", "A1:Z"),
]

# ----------------------------
# Helpers
# ----------------------------
EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)

def normalize_email(x: str) -> str:
    return str(x).strip().lower() if x else ""

def normalize_phone(x: str) -> str:
    if not x:
        return ""
    digits = re.sub(r"\D", "", str(x))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else ""

def split_name(full: str):
    if not full:
        return "", ""
    parts = str(full).strip().split()
    if len(parts) == 1:
        return parts[0].title(), ""
    return parts[0].title(), " ".join(parts[1:]).title()

def extract_email(row):
    for cell in row:
        m = EMAIL_RE.search(str(cell or ""))
        if m:
            return normalize_email(m.group(0))
    return ""

def extract_phone(row):
    for cell in row:
        p = normalize_phone(cell)
        if p:
            return p
    return ""

def now_iso():
    return datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")

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

def ensure_master_headers(svc):
    existing = get_values(svc, MASTER_SHEET, "A1:Z1")
    if not existing or existing[0] != MASTER_HEADERS:
        update_values(svc, MASTER_SHEET, "A1", [MASTER_HEADERS])

def looks_like_header_row(row):
    # If the row contains words like "email" "phone" etc, treat as header row
    joined = " ".join(str(c).lower() for c in row)
    return any(k in joined for k in ["email", "e-mail", "phone", "first", "last", "name", "address", "zip"])

# ----------------------------
# Core: Build/Rewrite Master
# ----------------------------
def load_existing_master_index(svc):
    """Return existing master rows and index maps so we preserve SENT/DNC."""
    ensure_master_headers(svc)
    rows = get_values(svc, MASTER_SHEET, "A2:Z")

    existing_rows = []
    by_email = {}
    by_phone = {}

    for r in rows:
        rr = (r + [""] * len(MASTER_HEADERS))[:len(MASTER_HEADERS)]
        existing_rows.append(rr)

        email = normalize_email(rr[2])
        phone = normalize_phone(rr[3])

        if email:
            by_email[email] = rr
        if phone:
            by_phone[phone] = rr

    return existing_rows, by_email, by_phone

def merge_row(existing, incoming):
    """
    Preserve existing status/sent_at/notes.
    Fill empty core fields from incoming.
    """
    merged = existing[:]
    # core fields 0..10 (through source_row)
    for idx in range(0, 11):
        if not merged[idx] and incoming[idx]:
            merged[idx] = incoming[idx]
    # preserve status/sent_at/notes (11..13) from existing
    return merged

def build_incoming_from_source(sheet_name, row_number_1based, row):
    email = extract_email(row)
    phone = extract_phone(row)

    first, last = split_name(row[0] if len(row) > 0 else "")

    # (Optional) You can improve these later by mapping exact columns
    age = ""
    address = ""
    city = ""
    state = ""
    zipc = ""

    rr = [""] * len(MASTER_HEADERS)
    rr[0] = first
    rr[1] = last
    rr[2] = email
    rr[3] = phone
    rr[4] = age
    rr[5] = address
    rr[6] = city
    rr[7] = state
    rr[8] = zipc
    rr[9] = sheet_name
    rr[10] = str(row_number_1based)
    rr[11] = ""      # status
    rr[12] = ""      # sent_at
    rr[13] = ""      # notes
    return rr

def rewrite_master(svc, rows):
    # Clear master data rows
    clear_range(svc, MASTER_SHEET, "A2:Z")
    if rows:
        update_values(svc, MASTER_SHEET, "A2", rows)

def normalize_all_sources_to_master():
    if not SPREADSHEET_ID:
        raise RuntimeError("Missing SPREADSHEET_ID in .env")

    svc = sheets_service()

    existing_rows, by_email, by_phone = load_existing_master_index(svc)

    # Start with existing master rows so SENT/DNC stays
    # We'll rebuild a final_map keyed by email/phone/name fallback for uniqueness
    final_by_key = {}

    def key_for(rr):
        email = normalize_email(rr[2])
        phone = normalize_phone(rr[3])
        if email:
            return f"email:{email}"
        if phone:
            return f"phone:{phone}"
        return f"name:{rr[0].strip().lower()}_{rr[1].strip().lower()}_{rr[8]}"

    # seed with existing
    for rr in existing_rows:
        final_by_key[key_for(rr)] = rr

    # ingest sources
    for sheet_name, a1 in SOURCE_SHEETS:
        rows = get_values(svc, sheet_name, a1)
        if not rows:
            continue

        start_idx = 1 if looks_like_header_row(rows[0]) else 0

        for i, row in enumerate(rows[start_idx:], start=start_idx + 1):
            row_number_1based = i + 1
            incoming = build_incoming_from_source(sheet_name, row_number_1based, row)

            inc_email = normalize_email(incoming[2])
            inc_phone = normalize_phone(incoming[3])

            existing = None
            if inc_email and inc_email in by_email:
                existing = by_email[inc_email]
            elif inc_phone and inc_phone in by_phone:
                existing = by_phone[inc_phone]

            if existing:
                merged = merge_row(existing, incoming)
                final_by_key[key_for(merged)] = merged
                # update indexes with merged version
                if merged[2]:
                    by_email[normalize_email(merged[2])] = merged
                if merged[3]:
                    by_phone[normalize_phone(merged[3])] = merged
            else:
                final_by_key[key_for(incoming)] = incoming
                if inc_email:
                    by_email[inc_email] = incoming
                if inc_phone:
                    by_phone[inc_phone] = incoming

    final_rows = list(final_by_key.values())

    # Optional: stable sort (status first, then source)
    def sort_key(rr):
        status = (rr[11] or "").upper()
        # keep SENT and DNC grouped
        priority = 0
        if status == "SENT":
            priority = 2
        elif status == "DO_NOT_CONTACT":
            priority = 3
        else:
            priority = 1
        return (priority, rr[0], rr[1], rr[2], rr[3])

    final_rows.sort(key=sort_key)

    ensure_master_headers(svc)
    rewrite_master(svc, final_rows)

    print(f"✅ Master rebuilt: {len(final_rows)} unique leads at {now_iso()}")

if __name__ == "__main__":
    normalize_all_sources_to_master()
