from __future__ import print_function
import os
import base64
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
import pandas as pd
import time
from dotenv import load_dotenv
load_dotenv()

# --- CONFIG ---
SPREADSHEET_ID = os.getenv("SPREADSHEET_ID")
RANGE_NAME = os.getenv("RANGE_NAME", "Sheet1!A:C")
ATTACHMENT_PATH = os.getenv("ATTACHMENT_PATH", "test.pdf")
FIXED_CC = os.getenv("FIXED_CC")
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT", "Superloopz - Universal Commerce Platform Opportunity")

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send'
]

def auth_google_services():
    """Authenticate using stored credentials.json inside the function."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            import json
            creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
            if creds_json:
                with open("credentials_temp.json", "w") as f:
                    f.write(creds_json)
                flow = InstalledAppFlow.from_client_secrets_file("credentials_temp.json", SCOPES)
            else:
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    sheets_service = build('sheets', 'v4', credentials=creds)
    gmail_service = build('gmail', 'v1', credentials=creds)
    return sheets_service, gmail_service

def get_pending_rows(sheets_service):
    result = sheets_service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME
    ).execute()
    values = result.get('values', [])

    if not values:
        print("No data found in sheet.")
        return pd.DataFrame(), pd.DataFrame()

    headers = [h.strip().title() for h in values[0]]
    rows = []
    for row in values[1:]:
        while len(row) < len(headers):
            row.append('')
        rows.append(row)

    df = pd.DataFrame(rows, columns=headers)
    if 'Status' not in df.columns:
        df['Status'] = ''
    df['Status'] = df['Status'].fillna('')
    pending_rows = df[df['Status'].str.lower().isin(['pending', ''])]
    return df, pending_rows

def send_email(gmail_service, name, to_email):
    import uuid, requests

    TRACKER_BASE = os.getenv("TRACKER_BASE")  # Replace with your deployed tracker URL
    track_id = str(uuid.uuid4())
    tracking_pixel = f'<img src="{TRACKER_BASE}/pixel/{track_id}.png" width="1" height="1"/>'
    
    msg = MIMEMultipart()
    msg['to'] = to_email
    msg['cc'] = FIXED_CC
    msg['subject'] = EMAIL_SUBJECT

    body = f"""<html><body>
<p>Dear {name},</p>

<p>Here’s a paradox: a manufacturer knows exactly how much it shipped to a distributor — but once the truck leaves the warehouse, visibility ends. Did those products sell out in San Francisco/Mumbai but stall in Los Angeles/Delhi? Did one retailer run dry in a week while another sat on stock? Most manufacturers don’t know, because their systems stop at the distributor’s door.</p>

<p>That blind spot is enormous. It drives overproduction in some markets, empty shelves in others, wasted marketing spend, and billions in lost revenue.</p>

<p>And it’s not just commerce. Last year, airlines lost an estimated $35 billion to delays, cancellations, and missed connections not because planes failed, but because systems didn’t talk to each other. Millions of passengers were double-booked or stranded because one database showed “available” while another showed “sold.”</p>

<p>Commerce is suffering from the same dysfunction.</p>

<p>At Superloopz, we’re building the Universal Commerce Platform to solve this. A SaaS backbone that unifies B2C, B2B/B2B2B, and service channels into one real-time dashboard. Manufacturers see beyond distributors into retailers, marketplaces, and even end-consumer demand. On top of this, we’re integrating agentic commerce capabilities AI agents that automate catalog updates, generate personalized marketing, and optimize demand forecasting. These agents reduce manual work and act proactively, helping businesses scale smarter. From D2C brands to manufacturers, hotels, and QSRs, our system enables anyone to sell anything, anywhere, through any channel efficiently and intelligently.</p>

<p>This has never been built end-to-end. Commerce today is fragmented across marketplaces, distributors, and aggregators. Yet the demand for unification is massive. Global B2B e-commerce alone will reach $61.9T by 2030, and brands are ready to pay millions for visibility because it directly protects their revenue.</p>

<p>I’d love to share how we’re tackling this opportunity and why we believe this will be the backbone of global commerce in the next decade. I’ve attached a detailed PRD with stats and real-world examples to give you a deeper view. Would you be open to a quick conversation?</p>

<p>Best,<br>
Praharsha<br>
Co-Founder, Superloopz<br>
+1 949-755-9016 | <a href="https://www.linkedin.com/in/praharsha-nelaturi/">My LinkedIn</a> | <a href="https://www.linkedin.com/in/chaitanyasai-g/">Co-founder&apos;s LinkedIn</a></p>

<p><a href="https://Superloopz.com">Superloopz.com</a><br>
Password for the doc - Fundus</p>
</body></html>"""

    msg.attach(MIMEText(body, 'html'))

    if os.path.exists(ATTACHMENT_PATH):
        with open(ATTACHMENT_PATH, "rb") as f:
            part = MIMEApplication(f.read(), Name=os.path.basename(ATTACHMENT_PATH))
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(ATTACHMENT_PATH)}"'
        msg.attach(part)

    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    try:
        result = gmail_service.users().messages().send(userId="me", body={'raw': raw_message}).execute()
        gmail_message_id = result.get('id', None)

        # Register this email with the tracker
        requests.post(f"{TRACKER_BASE}/_register_send", json={
            "track_id": track_id,
            "recipient_email": to_email,
            "subject": EMAIL_SUBJECT,
            "gmail_message_id": gmail_message_id
        })

        print(f"✅ Email sent to {to_email}")
        return True
    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")
        return False

def update_status(sheets_service, df, row_index, status):
    cell = f"C{row_index+2}"
    sheets_service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=cell,
        valueInputOption="RAW",
        body={"values": [[status]]}
    ).execute()

def process_emails():
    sheets_service, gmail_service = auth_google_services()
    df, pending_rows = get_pending_rows(sheets_service)
    if pending_rows.empty:
        print("No pending rows found.")
        return "No pending rows found."
    for idx, row in pending_rows.iterrows():
        name, email = row['Name'], row['Email']
        success = send_email(gmail_service, name, email)
        update_status(sheets_service, df, idx, "Completed" if success else "Error")
        time.sleep(2)
    return "Email automation run complete."

# ---- Cloud Function entry point ----
def send_emails(request):
    return process_emails()