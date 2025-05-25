import os
import time
import base64
from email.mime.text import MIMEText
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

# --- Load environment variables ---
load_dotenv()
AGENT_NAME = os.getenv("AGENT_NAME")
AGENT_LICENSE = os.getenv("AGENT_NUMBER")
WORK_PHONE = os.getenv("WORK_PHONE")
WORK_EMAIL = os.getenv("WORK_EMAIL")
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")

# --- Gmail API (OAuth2) ---
GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def authenticate_gmail():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', GMAIL_SCOPES)
    else:
        flow = InstalledAppFlow.from_client_secrets_file('credentials.json', GMAIL_SCOPES)
        creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

# --- Google Sheets API (Service Account) ---
SHEETS_SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
SERVICE_ACCOUNT_FILE = 'sheet_service_account.json'

def get_leads_from_sheets():
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SHEETS_SCOPES)
    service = build('sheets', 'v4', credentials=creds)
    sheet = service.spreadsheets()
    RANGE_NAME = "'Sheet1'!A1:D"
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    rows = result.get('values', [])
    leads = []
    for row in rows:
        if len(row) >= 4:
            name = row[0].strip().capitalize()
            email = row[3]
            leads.append((name, email))
    return leads

# --- Gmail Functions ---
def create_message(to, subject, body_text):
    message = MIMEText(body_text)
    message['to'] = to
    message['subject'] = subject
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {'raw': raw}

def send_email(service, to_name, to_email):
    subject = "Your Life Insurance Program Eligibility – Urgent"
    body = f"""Hi {to_name},

I hope this reaches you in time — I was just reviewing my file list, and your name popped up as someone who may still be eligible for the new California Life Insurance Programs with added living benefits.

These programs were created to help families like yours lock in permanent coverage that never expires and access your money early in case of a critical illness — something most older policies don’t cover.

If you already have coverage, great — but most of the families I speak to are shocked to find out:

Their policies expire or increase in price as they age

They can’t access any of the money if something happens while they’re still alive

They’re paying more than they should without realizing it

If you want to see how these new options compare to what you currently have (or what you could qualify for), I can send you a personalized breakdown. Just reply to this email or call/text me directly.

🕒 Time-sensitive: I only have access to your file for a short period before it gets closed out.

Talk soon,
{AGENT_NAME}
Licensed Broker - Family First Life  
CA License: {AGENT_LICENSE}  
📞 {WORK_PHONE}  
📧 {WORK_EMAIL}
"""
    message = create_message(to_email, subject, body)
    service.users().messages().send(userId='me', body=message).execute()

# --- Run Program ---
if __name__ == '__main__':
    gmail_service = authenticate_gmail()
    leads = get_leads_from_sheets()
    
    for name, email in leads:
        send_email(gmail_service, name, email)
        print(f"✅ Sent email to {name} at {email}")
        time.sleep(2)
