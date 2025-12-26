import os
import time
import base64
import re
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
from email.mime.image import MIMEImage

from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.exceptions import RefreshError

count = 0
BUSINESS_CARD_PATH = r"images\\JC_BusinessCard.png"

# --- Load environment variables ---
load_dotenv()
AGENT_NAME = os.getenv("AGENT_NAME")
AGENT_LICENSE = os.getenv("AGENT_NUMBER")
WORK_PHONE = os.getenv("WORK_PHONE")
WORK_EMAIL = os.getenv("WORK_EMAIL")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

# --- Gmail API (OAuth2) ---
GMAIL_SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

def authenticate_gmail():
    creds = None

    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", GMAIL_SCOPES)

    try:
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
            with open("token.json", "w") as token:
                token.write(creds.to_json())

        return build("gmail", "v1", credentials=creds)

    except RefreshError:
        # token is revoked/broken -> delete + re-auth
        if os.path.exists("token.json"):
            os.remove("token.json")

        flow = InstalledAppFlow.from_client_secrets_file("credentials.json", GMAIL_SCOPES)
        creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

        return build("gmail", "v1", credentials=creds)

# --- Google Sheets API (Service Account) ---
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "sheet_service_account.json"

# ✅ Edit these and press ▶ in VS Code
TARGET_SHEET_NAME = "testsheet"
TARGET_RANGE = "A1:ZZ"

# ✅ Header aliases (added "number" under phone)
ALIASES = {
    "first_name": ["first", "first name", "firstname", "fname", "given name"],
    "last_name": ["last", "last name", "lastname", "lname", "surname", "family name"],
    "full_name": ["name", "full name", "fullname", "client name", "prospect name"],
    "email": ["email", "e-mail", "email address", "mail"],
    "phone": ["phone", "phone number", "number", "mobile", "cell", "cell phone", "telephone", "tel", "contact number"],
    "email_sent": ["email_sent", "email sent", "emailed", "emailed_date", "email date", "sent at", "sent_on", "sent date"],
}

EMAIL_RE = re.compile(r"\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b", re.I)

def sheets_service():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SHEETS_SCOPES
    )
    return build("sheets", "v4", credentials=creds)

def normalize_header(h: str) -> str:
    h = (h or "").strip().lower()
    h = re.sub(r"[\s\-_]+", " ", h)
    h = re.sub(r"[^a-z0-9 ]+", "", h)
    return h

def build_header_map(header_row):
    headers = [normalize_header(h) for h in header_row]
    idx = {}

    # exact match
    for field, aliases in ALIASES.items():
        for i, h in enumerate(headers):
            if h in aliases:
                idx[field] = i
                break

    # contains match
    for field, aliases in ALIASES.items():
        if field in idx:
            continue
        for i, h in enumerate(headers):
            for a in aliases:
                if a in h:
                    idx[field] = i
                    break
            if field in idx:
                break

    return idx

def get_cell(row, col_idx):
    if col_idx is None:
        return ""
    return row[col_idx] if col_idx < len(row) else ""

def normalize_email(x: str) -> str:
    if not x:
        return ""
    m = EMAIL_RE.search(str(x).strip())
    return m.group(0).lower() if m else ""

def normalize_phone(x: str) -> str:
    if not x:
        return ""
    digits = re.sub(r"\D", "", str(x))
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits if len(digits) == 10 else str(x).strip()

def titlecase_name(x: str) -> str:
    return str(x).strip().title() if x else ""

def split_name(full: str):
    full = (full or "").strip()
    if not full:
        return "", ""
    parts = full.split()
    if len(parts) == 1:
        return parts[0].title(), ""
    return parts[0].title(), " ".join(parts[1:]).title()

def col_index_to_letter(idx0: int) -> str:
    n = idx0 + 1
    letters = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters

def now_timestamp_local():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def read_sheet_rows(sheets_svc):
    rng = f"'{TARGET_SHEET_NAME}'!{TARGET_RANGE}"
    result = sheets_svc.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=rng
    ).execute()
    return result.get("values", [])

def update_cell(sheets_svc, col_idx0: int, row_number_1based: int, value: str):
    col_letter = col_index_to_letter(col_idx0)
    a1 = f"'{TARGET_SHEET_NAME}'!{col_letter}{row_number_1based}"
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=a1,
        valueInputOption="RAW",
        body={"values": [[value]]}
    ).execute()

def ensure_email_sent_column_exists(rows, header_map, sheets_svc):
    header = rows[0] if rows else []
    if "email_sent" in header_map:
        return header_map, header

    new_header = header[:] + ["email_sent"]
    sheets_svc.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{TARGET_SHEET_NAME}'!A1",
        valueInputOption="RAW",
        body={"values": [new_header]}
    ).execute()

    header_map = build_header_map(new_header)
    return header_map, new_header

def format_phone_us(digits: str) -> str:
    """
    Takes 10-digit string and formats as (XXX) XXX-XXXX
    """
    if not digits:
        return ""
    digits = re.sub(r"\D", "", digits)
    if len(digits) != 10:
        return digits  # fallback: return as-is
    return f"({digits[0:3]}) {digits[3:6]}-{digits[6:10]}"


# --- Gmail Functions ---
def create_message(to, subject, body_html, image_path=None):
    message = MIMEMultipart("related")
    message["to"] = to
    message["subject"] = subject

    alternative = MIMEMultipart("alternative")
    message.attach(alternative)

    alternative.attach(MIMEText(body_html, "html", "utf-8"))

    # Attach image inline (CID)
    if image_path:
        if not os.path.exists(image_path):
            raise FileNotFoundError(f"Business card image not found: {image_path}")

        with open(image_path, "rb") as f:
            img = MIMEImage(f.read())
            img.add_header("Content-ID", "<businesscard>")
            img.add_header("Content-Disposition", "inline", filename=os.path.basename(image_path))
            message.attach(img)

    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {"raw": raw}



def send_email(gmail_service, to_name, to_email, to_phone):
    subject = "Something I noticed in your file..."

    body_html = f"""
    <html>
      <body style="font-family: Arial, Helvetica, sans-serif; font-size: 14px; color: #000; line-height: 1.6;">
        <p>Hi {to_name},</p>

        <p>
          I was reviewing some older records today and noticed your file was left
          <strong>open</strong>.
        </p>

        <p>
          In the time since we last spoke, the "safety net" most families rely on has become…
          <strong>different</strong>. You’ve likely felt the shift—the way certain protections
          aren’t as firm as they used to be. There is a specific
          <strong>blind spot</strong> in many older plans that often goes unnoticed until the
          moment it's actually needed.
        </p>

        <p>
          I’m not sure if your situation has evolved, but leaving that gap
          <strong>unattended</strong> is a risk that weighs more heavily now than it did a year ago.
        </p>

        <p>
          I’ve closed this gap for several others recently. It’s a quiet fix, but the
          <strong>relief</strong> it brings is immediate.
        </p>

        <p>
          Are you still at {to_phone}, or should we exchange a few notes here?
        </p>

        <p>
          Best,<br><br>
          {AGENT_NAME}<br>
          Life Insurance &amp; Annuities Broker<br>
          CA License: {AGENT_LICENSE}<br>
          📞 {WORK_PHONE}<br>
          📧 {WORK_EMAIL}<br>
          Book an appointment on my calendar:
          <a href="https://calendly.com/justingimho/life-insurance-consulting">https://calendly.com/justingimho/life-insurance-consulting</a>
        </p>
        <p style="margin-top:20px;">
        <img src="cid:businesscard"
        alt="Business Card"
        style="max-width:420px;width:100%;border-radius:6px;">
        </p>

        <hr style="border:none;border-top:1px solid #ddd;margin:20px 0;">

        <p style="font-size: 12px; color: #555;">
          If you’d prefer not to receive emails from me, just reply with “unsubscribe”
          and I’ll take you off my list.
        </p>
      </body>
    </html>
    """

    msg = create_message(to_email, subject, body_html, image_path=BUSINESS_CARD_PATH)
    gmail_service.users().messages().send(userId="me", body=msg).execute()

# --- Run Program ---
if __name__ == "__main__":
    gmail_service = authenticate_gmail()
    sheets_svc = sheets_service()

    rows = read_sheet_rows(sheets_svc)
    if not rows or len(rows) < 2:
        raise RuntimeError("Sheet is empty or missing data rows.")

    header = rows[0]
    header_map = build_header_map(header)

    if "email" not in header_map:
        raise RuntimeError(
            f"Couldn't find an EMAIL column in '{TARGET_SHEET_NAME}'.\n"
            f"Header row was: {header}\n"
            f"Rename header to one of: {ALIASES['email']}"
        )

    if "phone" not in header_map:
        raise RuntimeError(
            f"Couldn't find a PHONE column in '{TARGET_SHEET_NAME}'.\n"
            f"Header row was: {header}\n"
            f"Rename header to one of: {ALIASES['phone']}"
        )

    # Ensure email_sent exists
    header_map, header = ensure_email_sent_column_exists(rows, header_map, sheets_svc)

    first_idx = header_map.get("first_name")
    last_idx = header_map.get("last_name")
    full_idx = header_map.get("full_name")
    email_idx = header_map.get("email")
    phone_idx = header_map.get("phone")
    email_sent_idx = header_map.get("email_sent")

    print("✅ Detected header mapping:", header_map)

    # Find first unsent row
    start_row_number_1based = None
    for row_number_1based, row in enumerate(rows[1:], start=2):
        sent_val = str(get_cell(row, email_sent_idx)).strip()
        email_val = normalize_email(get_cell(row, email_idx))
        if email_val and not sent_val:
            start_row_number_1based = row_number_1based
            break

    if start_row_number_1based is None:
        print("✅ No unsent leads found (everyone has email_sent filled).")
        raise SystemExit(0)

    print(f"▶ Starting from first unsent lead at row {start_row_number_1based}...")

    for row_number_1based, row in enumerate(rows[start_row_number_1based - 1:], start=start_row_number_1based):
        email = normalize_email(get_cell(row, email_idx))
        if not email:
            continue

        sent_val = str(get_cell(row, email_sent_idx)).strip()
        if sent_val:
            continue

        first = titlecase_name(get_cell(row, first_idx)) if first_idx is not None else ""
        last = titlecase_name(get_cell(row, last_idx)) if last_idx is not None else ""
        if (not first and not last) and full_idx is not None:
            first, last = split_name(get_cell(row, full_idx))
        name_for_greeting = first or "there"

        raw_phone = normalize_phone(get_cell(row, phone_idx))
        to_phone = format_phone_us(raw_phone)

        if not to_phone:
            # If you want to skip rows with missing phone, keep this:
            # continue
            to_phone = "your current number"

        send_email(gmail_service, name_for_greeting, email, to_phone)

        ts = now_timestamp_local()
        update_cell(sheets_svc, email_sent_idx, row_number_1based, ts)

        count += 1
        print(f"✅ Sent email #{count} to {name_for_greeting} at {email} | phone={to_phone} | email_sent={ts}")
        time.sleep(2)

        if count == 50:
            quit("Reached 50 emails sent. Stopping to avoid rate limits.")
