import os
import base64
import re
from dotenv import load_dotenv

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.oauth2 import service_account

# --- Configuration ---
load_dotenv()
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
TARGET_SHEET_NAME = "Invalid_Email" # Change this to the exact name of your tab

# --- Scopes ---
GMAIL_SCOPES = ['https://mail.google.com/']
SHEETS_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SERVICE_ACCOUNT_FILE = "sheet_service_account.json"

def get_email_body(payload):
    """Recursively parses the email payload to extract plain text."""
    body = ""
    if 'parts' in payload:
        for part in payload['parts']:
            body += get_email_body(part)
    elif payload.get('mimeType') == 'text/plain':
        data = payload['body'].get('data')
        if data:
            body = base64.urlsafe_b64decode(data).decode('utf-8', errors='ignore')
    return body

def append_to_sheet(data):
    """Appends the list of bounced data directly to the Google Sheet."""
    if not data:
        return
        
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SHEETS_SCOPES
    )
    sheets_svc = build("sheets", "v4", credentials=creds)
    
    # We append starting at column A
    range_name = f"'{TARGET_SHEET_NAME}'!A:C"
    
    body = {
        "values": data
    }
    
    sheets_svc.spreadsheets().values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=range_name,
        valueInputOption="RAW",
        insertDataOption="INSERT_ROWS",
        body=body
    ).execute()

def main():
    creds = None
    # The file token_cleanup.json stores your Gmail access.
    if os.path.exists('token_cleanup.json'):
        creds = Credentials.from_authorized_user_file('token_cleanup.json', GMAIL_SCOPES)
    
    # Authenticate Gmail
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token_cleanup.json', 'w') as token:
            token.write(creds.to_json())

    # Build the Gmail API service
    service = build('gmail', 'v1', credentials=creds)

    print("Searching for all bounce-backs (this might take a second if there are hundreds)...")
    
    messages = []
    page_token = None

    # This loop keeps asking Google for the next page of emails until it runs out
    while True:
        results = service.users().messages().list(
            userId='me', 
            q='from:mailer-daemon', 
            maxResults=500,  # Max allowed per request
            pageToken=page_token
        ).execute()
        
        messages.extend(results.get('messages', []))
        page_token = results.get('nextPageToken')
        
        if not page_token:
            break

    if not messages:
        print("No bounce-back messages found.")
        return

    bounced_data = []
    email_pattern = re.compile(r"to\s+([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})")
    total_messages = len(messages)
    
    print(f"Found {total_messages} bounce-backs. Processing them now...")

    for index, msg_meta in enumerate(messages, start=1):
        # This will update the same line in your terminal so it doesn't spam your screen
        print(f"Processing {index} of {total_messages}...", end='\r')
        
        msg_id = msg_meta['id']
        # Fetch the full message
        msg = service.users().messages().get(userId='me', id=msg_id, format='full').execute()
        
        payload = msg['payload']
        headers = payload.get('headers', [])
        
        subject = next((header['value'] for header in headers if header['name'].lower() == 'subject'), "Unknown Subject")
        
        body = get_email_body(payload)
        if not body:
            body = msg.get('snippet', '')

        # Regex to find the failed email
        failed_email_match = email_pattern.search(body.lower())
        failed_email = failed_email_match.group(1) if failed_email_match else "Could not extract"

        # Determine the error reason
        error_reason = "Unknown Error"
        if "address not found" in body.lower() or "address not found" in subject.lower():
            error_reason = "Address not found"
        elif "inbox full" in body.lower() or "inbox is full" in body.lower():
            error_reason = "Inbox full"
        elif "delivery incomplete" in body.lower():
            error_reason = "Delivery incomplete / Temporary problem"

        bounced_data.append([failed_email, error_reason, subject])

        # Move the email to the Trash
        service.users().messages().trash(userId='me', id=msg_id).execute()
        
    # Jump to the next line after the loop is done
    print("\nProcessing complete!")
        
    # Append the extracted data directly to Google Sheets
    print(f"Uploading {len(bounced_data)} rows to Google Sheets...")
    append_to_sheet(bounced_data)

    print(f"Successfully processed, trashed, and uploaded {len(bounced_data)} bounce-backs.")

if __name__ == '__main__':
    main()