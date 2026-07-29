# common.py
# Shared helpers used by every script in this toolkit.
#
# Before this module existed, each script carried its own slightly different
# copy of normalize_phone / normalize_email / header matching / auth, and the
# differences caused real bugs (columns matched inconsistently, phones blanked
# in one script but kept in another, one script refreshed OAuth tokens and the
# other didn't). Everything lives here now, once.

import json
import os
import random
import re
import time
from datetime import datetime

import httplib2
from dotenv import load_dotenv
from google.auth.exceptions import RefreshError
from google.auth.transport.requests import Request
from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv()

SPREADSHEET_ID = os.getenv("SPREADSHEET_ID", "")
SERVICE_ACCOUNT_FILE = os.getenv("SHEETS_SERVICE_ACCOUNT_FILE", "sheet_service_account.json")
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# Tabs that hold suppression / bookkeeping data rather than leads.
INVALID_EMAIL_TAB = os.getenv("INVALID_EMAIL_TAB", "Invalid_Email")
UNSUBSCRIBES_TAB = os.getenv("UNSUBSCRIBES_TAB", "Unsubscribes")
SOFT_BOUNCES_TAB = os.getenv("SOFT_BOUNCES_TAB", "Soft_Bounces")
BOUNCE_REVIEW_TAB = os.getenv("BOUNCE_REVIEW_TAB", "Bounce_Review")
BOOKKEEPING_TABS = [INVALID_EMAIL_TAB, UNSUBSCRIBES_TAB, SOFT_BOUNCES_TAB, BOUNCE_REVIEW_TAB]

EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)
ZIP_RE = re.compile(r"\b\d{5}(-\d{4})?\b")

# Canonical header aliases. Exact matches are tried across ALL columns first;
# only then does a word-boundary fallback run, and only on columns no other
# field has already claimed (so an "email_sent" column can never steal the
# "email" mapping, and "mailing address" can never be mistaken for an email).
ALIASES = {
    "first_name": ["first", "first name", "firstname", "fname", "given name"],
    "last_name": ["last", "last name", "lastname", "lname", "surname", "family name"],
    "full_name": ["name", "full name", "fullname", "client name", "prospect name"],
    "email": ["email", "e-mail", "email address", "mail", "gmail"],
    "phone": ["phone", "phone number", "phonenumber", "mobile", "cell", "cell phone",
              "telephone", "tel", "contact number", "number", "contact"],
    "email_sent": ["email_sent", "email sent", "emailed", "emailed_date", "email date",
                   "sent at", "sent_on", "sent date"],
    "emailed_date": ["emailed date", "date emailed"],
    "followup_sent": ["followup_sent", "followup sent", "follow up sent", "followed up"],
    "replied": ["replied", "replied_at", "reply", "responded"],
    "status": ["status"],
    "notes": ["notes", "note", "comments", "comment"],
    "age": ["age"],
    "address": ["address", "street", "street address", "address1", "address 1"],
    "city": ["city"],
    "state": ["state", "st", "province"],
    "zip": ["zip", "zipcode", "zip code", "postal", "postal code"],
}

# Aliases this short are only safe as exact matches ("st" would otherwise
# match "first name", "mail" would match "mailing address").
MIN_FUZZY_ALIAS_LEN = 4


# ---------------------------------------------------------------------------
# Environment helpers
# ---------------------------------------------------------------------------

def require_env(*names):
    """Return the values of required .env vars, aborting with a clear message
    if any are missing — instead of sending emails with 'None' in them."""
    values, missing = [], []
    for name in names:
        v = (os.getenv(name) or "").strip()
        values.append(v)
        if not v:
            missing.append(name)
    if missing:
        raise SystemExit(
            "❌ Missing required .env value(s): " + ", ".join(missing)
            + "\nAdd them to your .env file (see .env.example) and re-run."
        )
    return values if len(values) > 1 else values[0]


def env_flag(name, default=False):
    v = (os.getenv(name) or "").strip().lower()
    if not v:
        return default
    return v in ("1", "true", "yes", "y", "on")


def env_int(name, default):
    v = (os.getenv(name) or "").strip()
    try:
        return int(v) if v else default
    except ValueError:
        return default


# ---------------------------------------------------------------------------
# Auth
# ---------------------------------------------------------------------------

def sheets_service():
    if not SPREADSHEET_ID:
        raise SystemExit("❌ Missing SPREADSHEET_ID in .env")
    if not os.path.exists(SERVICE_ACCOUNT_FILE):
        raise SystemExit(
            f"❌ Service account key not found: {SERVICE_ACCOUNT_FILE}\n"
            "Download it from Google Cloud (IAM & Admin > Service Accounts > Keys) "
            f"and save it as {SERVICE_ACCOUNT_FILE} in this folder."
        )
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SHEETS_SCOPES
    )
    return build("sheets", "v4", credentials=creds)


def gmail_service(scopes, token_file, credentials_file="credentials.json"):
    """OAuth flow that refreshes silently when possible and falls back to the
    browser flow when the token is missing, revoked, or lacks the scopes."""
    creds = None
    if os.path.exists(token_file):
        try:
            with open(token_file) as fh:
                info = json.load(fh)
            # Compare against the scopes STORED in the token: constructing
            # Credentials with a scopes argument would report the requested
            # scopes back, making the check always pass.
            stored = set(info.get("scopes") or [])
            if set(scopes) <= stored:
                creds = Credentials.from_authorized_user_info(info, scopes)
        except (ValueError, KeyError):
            creds = None

    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
        except RefreshError:
            creds = None

    if not creds or not creds.valid:
        if not os.path.exists(credentials_file):
            raise SystemExit(
                f"❌ OAuth client file not found: {credentials_file}\n"
                "Create a Desktop-app OAuth client in Google Cloud Console and "
                "save the downloaded JSON as credentials.json in this folder."
            )
        flow = InstalledAppFlow.from_client_secrets_file(credentials_file, scopes)
        creds = flow.run_local_server(port=0)

    with open(token_file, "w") as f:
        f.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


# ---------------------------------------------------------------------------
# Retry / pacing
# ---------------------------------------------------------------------------

RETRYABLE_STATUS = {429, 500, 502, 503}
TRANSIENT_ERRORS = (ConnectionError, TimeoutError, httplib2.HttpLib2Error)


def with_backoff(fn, *, retries=5, base_delay=2.0, what="API call"):
    """Run fn(); on a retryable HttpError (429/5xx) or a transient transport
    error wait 2s, 4s, 8s... and retry. Other errors raise immediately."""
    for attempt in range(retries + 1):
        try:
            return fn()
        except HttpError as e:
            status = getattr(getattr(e, "resp", None), "status", None)
            if status not in RETRYABLE_STATUS or attempt == retries:
                raise
            reason = f"HTTP {status}"
        except TRANSIENT_ERRORS as e:
            if attempt == retries:
                raise
            reason = type(e).__name__
        delay = base_delay * (2 ** attempt)
        print(f"⏳ {what} got {reason}; retrying in {delay:.0f}s "
              f"(attempt {attempt + 1}/{retries})...")
        time.sleep(delay)


def jitter_sleep(min_seconds, max_seconds):
    time.sleep(random.uniform(min_seconds, max_seconds))


# ---------------------------------------------------------------------------
# Normalization
# ---------------------------------------------------------------------------

def normalize_header(h):
    h = (h or "").strip().lower()
    h = re.sub(r"[\s\-_]+", " ", h)
    h = re.sub(r"[^a-z0-9 ]+", "", h)
    return h.strip()


def normalize_email(x):
    if not x:
        return ""
    m = EMAIL_RE.search(str(x).strip())
    return m.group(0).lower() if m else ""


def normalize_phone(x, keep_original=True):
    """Return a bare 10-digit US number when possible. When the value doesn't
    normalize, return the original string (keep_original=True, so data is
    never silently blanked) or '' (keep_original=False, for building match
    sets where only valid numbers are useful)."""
    if not x:
        return ""
    digits = re.sub(r"\D", "", str(x))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    if len(digits) == 10:
        return digits
    return str(x).strip() if keep_original else ""


def format_phone_us(digits):
    if not digits:
        return ""
    digits = re.sub(r"\D", "", str(digits))
    if len(digits) != 10:
        return str(digits)
    return f"({digits[0:3]}) {digits[3:6]}-{digits[6:10]}"


def titlecase_name(x):
    """Title-cased name with ALL whitespace collapsed to single spaces — a
    newline inside a name cell would otherwise end up in the Subject header
    and crash message serialization."""
    if not x:
        return ""
    return " ".join(str(x).split()).title()


def split_name(full):
    full = (full or "").strip()
    if not full:
        return "", ""
    parts = full.split()
    if len(parts) == 1:
        return parts[0].title(), ""
    return parts[0].title(), " ".join(parts[1:]).title()


def now_timestamp():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------------------------------------------------------------------------
# Header mapping
# ---------------------------------------------------------------------------

def build_header_map(header_row, aliases=None):
    """Map canonical field names -> column index.

    Pass 1: exact match of the normalized header against normalized aliases,
    for every field across every column (so an exact 'State' column always
    beats an earlier 'Statement Date').
    Pass 2: word-boundary fallback, longest aliases first, skipping columns
    already claimed in pass 1 and aliases too short to be safe.
    """
    aliases = aliases or ALIASES
    headers = [normalize_header(h) for h in header_row]
    norm_aliases = {
        field: [normalize_header(a) for a in alist] for field, alist in aliases.items()
    }

    idx = {}
    claimed = set()

    # Aliases are listed strongest-first, so try each alias across ALL
    # columns before moving to a weaker one — a broad alias like "contact"
    # can then never hijack a column when a real "Phone" column exists.
    for field, alist in norm_aliases.items():
        for a in alist:
            hit = next(
                (i for i, h in enumerate(headers)
                 if i not in claimed and h and h == a),
                None,
            )
            if hit is not None:
                idx[field] = hit
                claimed.add(hit)
                break

    for field, alist in norm_aliases.items():
        if field in idx:
            continue
        for a in sorted(alist, key=len, reverse=True):
            if len(a) < MIN_FUZZY_ALIAS_LEN:
                continue
            for i, h in enumerate(headers):
                if i in claimed or not h:
                    continue
                words = h.split()
                if (" " in a and a in h) or (" " not in a and a in words):
                    idx[field] = i
                    claimed.add(i)
                    break
            if field in idx:
                break

    return idx


def get_cell(row, idx):
    if idx is None:
        return ""
    return row[idx] if idx < len(row) else ""


def col_index_to_letter(idx0):
    n = idx0 + 1
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ---------------------------------------------------------------------------
# Sheets I/O
# ---------------------------------------------------------------------------

def quote_tab(sheet_name):
    """A1-notation tab reference; embedded single quotes must be doubled or
    every range for that tab fails with 'Unable to parse range'."""
    return "'" + sheet_name.replace("'", "''") + "'"


def get_values(svc, sheet_name, a1_range="A1:ZZ"):
    rng = f"{quote_tab(sheet_name)}!{a1_range}"
    resp = with_backoff(
        lambda: svc.spreadsheets().values().get(
            spreadsheetId=SPREADSHEET_ID, range=rng
        ).execute(),
        what=f"read {sheet_name}",
    )
    return resp.get("values", [])


def update_values(svc, sheet_name, start_cell, values):
    rng = f"{quote_tab(sheet_name)}!{start_cell}"
    return with_backoff(
        lambda: svc.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=rng,
            valueInputOption="RAW",
            body={"values": values},
        ).execute(),
        what=f"write {sheet_name}",
    )


def append_values(svc, sheet_name, values):
    return with_backoff(
        lambda: svc.spreadsheets().values().append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{quote_tab(sheet_name)}!A:Z",
            valueInputOption="RAW",
            insertDataOption="INSERT_ROWS",
            body={"values": values},
        ).execute(),
        what=f"append {sheet_name}",
    )


def clear_values(svc, sheet_name, a1_range):
    return with_backoff(
        lambda: svc.spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID, range=f"{quote_tab(sheet_name)}!{a1_range}", body={}
        ).execute(),
        what=f"clear {sheet_name}",
    )


def list_sheet_titles(svc):
    meta = with_backoff(
        lambda: svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute(),
        what="read spreadsheet metadata",
    )
    return [s["properties"]["title"] for s in meta.get("sheets", [])]


def get_sheet_props(svc, sheet_name):
    """Properties dict for a tab title (case-insensitive). None if absent."""
    meta = with_backoff(
        lambda: svc.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute(),
        what="read spreadsheet metadata",
    )
    want = sheet_name.strip().lower()
    for s in meta.get("sheets", []):
        if s["properties"]["title"].strip().lower() == want:
            return s["properties"]
    return None


def get_sheet_id(svc, sheet_name):
    """Numeric sheetId for a tab title (case-insensitive). None if absent."""
    props = get_sheet_props(svc, sheet_name)
    return props["sheetId"] if props else None


def find_tab_title(svc, sheet_name):
    """Actual tab title matching sheet_name case-insensitively, or None."""
    want = sheet_name.strip().lower()
    for title in list_sheet_titles(svc):
        if title.strip().lower() == want:
            return title
    return None


def ensure_tab(svc, sheet_name, header=None):
    """Create the tab (with an optional header row) if it doesn't exist.
    Returns the actual tab title."""
    existing = find_tab_title(svc, sheet_name)
    if existing:
        if header:
            rows = get_values(svc, existing, "A1:Z1")
            if not rows:
                update_values(svc, existing, "A1", [header])
        return existing
    with_backoff(
        lambda: svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body={"requests": [{"addSheet": {"properties": {"title": sheet_name}}}]},
        ).execute(),
        what=f"create tab {sheet_name}",
    )
    if header:
        update_values(svc, sheet_name, "A1", [header])
    return sheet_name


def overwrite_tab(svc, sheet_name, header, rows):
    """Replace a tab's contents WITHOUT the clear-then-write data-loss window:
    write the new data first, then clear only the leftover cells below and to
    the right of it. A crash mid-way leaves the tab with data, never empty."""
    values = [header] + rows
    update_values(svc, sheet_name, "A1", values)
    clear_values(svc, sheet_name, f"A{len(values) + 1}:ZZ")
    width = max(len(r) for r in values)
    first_col_after = col_index_to_letter(width)  # letter of column width+1
    clear_values(svc, sheet_name, f"{first_col_after}1:ZZ{len(values)}")


def delete_rows(svc, sheet_name, row_numbers_1based):
    """Delete specific rows in ONE atomic batchUpdate (bottom-up, contiguous
    runs merged). Safer than rewriting the whole tab."""
    if not row_numbers_1based:
        return
    sheet_id = get_sheet_id(svc, sheet_name)
    if sheet_id is None:
        raise RuntimeError(f"Tab not found: {sheet_name}")

    zero_based = sorted({r - 1 for r in row_numbers_1based}, reverse=True)
    requests = []
    run_end = run_start = zero_based[0]
    for r in zero_based[1:]:
        if r == run_start - 1:
            run_start = r
        else:
            requests.append({"deleteDimension": {"range": {
                "sheetId": sheet_id, "dimension": "ROWS",
                "startIndex": run_start, "endIndex": run_end + 1}}})
            run_end = run_start = r
    requests.append({"deleteDimension": {"range": {
        "sheetId": sheet_id, "dimension": "ROWS",
        "startIndex": run_start, "endIndex": run_end + 1}}})

    with_backoff(
        lambda: svc.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID, body={"requests": requests}
        ).execute(),
        what=f"delete rows in {sheet_name}",
    )


def ensure_column(svc, sheet_name, header, header_map, field,
                  header_label=None, rows=None):
    """Make sure a tracking column exists on the tab; append it to the header
    row if missing. Pass the CURRENT header row (and, when available, all
    rows) and thread the returned header into any further ensure_column
    calls. Returns (header_map, header_row).

    The new column goes AFTER the widest data row, not just after the header
    — otherwise stray values in an unlabeled trailing column would be
    mistaken for sent-markers and those leads silently skipped."""
    header = list(header or [])
    if field in header_map:
        return header_map, header

    width = len(header)
    if rows:
        width = max(width, max(len(r) for r in rows))
    if width > len(header):
        print(f"⚠️ '{sheet_name}' has data beyond the labeled headers; "
              f"placing the {header_label or field} column after it "
              f"(column {col_index_to_letter(width)}).")

    # values.update can't write past the tab's grid edge — widen the grid
    # first if the new column wouldn't fit.
    props = get_sheet_props(svc, sheet_name)
    if props:
        col_count = props.get("gridProperties", {}).get("columnCount", 0)
        if width + 1 > col_count:
            with_backoff(
                lambda: svc.spreadsheets().batchUpdate(
                    spreadsheetId=SPREADSHEET_ID,
                    body={"requests": [{"appendDimension": {
                        "sheetId": props["sheetId"],
                        "dimension": "COLUMNS",
                        "length": width + 1 - col_count,
                    }}]},
                ).execute(),
                what=f"widen {sheet_name}",
            )

    new_header = header + [""] * (width - len(header)) + [header_label or field]
    update_values(svc, sheet_name, "A1", [new_header])
    return build_header_map(new_header), new_header


# ---------------------------------------------------------------------------
# Suppression list (CAN-SPAM)
# ---------------------------------------------------------------------------

def load_suppression_set(svc, tabs=None):
    """Every email address found in the bounce + unsubscribe tabs. Checked
    before ANY send: once someone bounces hard or opts out, they are never
    emailed again. Missing tabs are skipped."""
    tabs = tabs if tabs is not None else [INVALID_EMAIL_TAB, UNSUBSCRIBES_TAB]
    suppressed = set()
    existing = {t.strip().lower() for t in list_sheet_titles(svc)}
    for tab in tabs:
        if tab.strip().lower() not in existing:
            continue
        for row in get_values(svc, tab):
            for cell in row:
                for m in EMAIL_RE.findall(str(cell or "")):
                    suppressed.add(m.lower())
    return suppressed


# ---------------------------------------------------------------------------
# Reply parsing (for the unsubscribe scanner)
# ---------------------------------------------------------------------------

QUOTE_MARKERS = (
    re.compile(r"^\s*>"),
    re.compile(r"^On .{0,200}wrote:\s*$"),
    re.compile(r"^-{2,}\s*Original Message\s*-{2,}", re.I),
    re.compile(r"^From:\s.*@", re.I),
)


def strip_quoted_text(body):
    """Keep only the text the person actually typed — drop quoted history.
    Vital because every reply quotes our own footer, which contains the word
    'unsubscribe'; without this, every reply would look like an opt-out."""
    kept = []
    for line in (body or "").splitlines():
        if any(p.search(line) for p in QUOTE_MARKERS):
            break
        kept.append(line)
    return "\n".join(kept)


UNSUB_RE = re.compile(
    r"\b(unsubscribe|opt[\s\-]?out|remove me|take me off|stop email(ing)?|"
    r"do not (email|contact)|no more emails?)\b",
    re.I,
)


def is_unsubscribe_text(text):
    return bool(UNSUB_RE.search(text or ""))
