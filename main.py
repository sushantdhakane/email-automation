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
ATTACHMENT_PATH = os.getenv("ATTACHMENT_PATH", "SuperLoopz_ AI Native Universal Commerce PRD.pdf")
FIXED_CC = os.getenv("FIXED_CC")
EMAIL_SUBJECT = os.getenv("EMAIL_SUBJECT", "Back us up — Together, let’s close the $60T gap in global commerce")

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/gmail.send'
]

def auth_google_services():
    """Authenticate using stored credentials.json or environment secrets. Supports non-interactive TOKEN_JSON when provided."""
    creds = None
    # 1) If a token JSON was provided as an environment secret, use it (headless)
    token_json = os.getenv("TOKEN_JSON")
    if token_json:
        import json as _json
        try:
            creds = Credentials.from_authorized_user_info(_json.loads(token_json), SCOPES)
        except Exception as e:
            print(f"⚠️ Failed to load TOKEN_JSON: {e}")
            creds = None

    # 2) Fallback to local token.json file
    if not creds and os.path.exists('token.json'):
        try:
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        except Exception as e:
            print(f"⚠️ Failed to load token.json: {e}")
            creds = None

    # 3) If no valid creds yet, try to create them.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"⚠️ Failed to refresh credentials: {e}")
                creds = None

        if not creds:
            import json
            creds_json = os.getenv("GOOGLE_CREDENTIALS_JSON")
            if creds_json:
                # write a temporary credentials file for the flow
                with open("credentials_temp.json", "w") as f:
                    f.write(creds_json)
                flow = InstalledAppFlow.from_client_secrets_file("credentials_temp.json", SCOPES)
                creds = flow.run_local_server(port=0)
                # remove temp file after use
                try:
                    os.remove("credentials_temp.json")
                except Exception:
                    pass
            else:
                # no credentials in env — attempt to use local file (interactive)
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
                creds = flow.run_local_server(port=0)

        # persist token.json for next runs (if we have credentials)
        if creds:
            try:
                with open('token.json', 'w') as token:
                    token.write(creds.to_json())
            except Exception as e:
                print(f"⚠️ Failed to write token.json: {e}")

    sheets_service = build('sheets', 'v4', credentials=creds)
    gmail_service = build('gmail', 'v1', credentials=creds)
    return sheets_service, gmail_service

def get_sender_email(gmail_service):
    profile = gmail_service.users().getProfile(userId="me").execute()
    return profile.get("emailAddress")

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

<p>There are two kinds of people in commerce today.</p>

<p>On one side, established manufacturers, distributors, and retailers struggle with fragmented systems across online, offline, B2B, and D2C channels. On the other, a college student at 2 AM designs a T-shirt and then spends hours just figuring out how to sell it online.</p>

<p><strong>Superloopz bridges both worlds.</strong></p>

<p>We’re building an <strong>AI-native Universal Commerce Platform</strong> — a SaaS operating system that unifies and automates every channel a business sells through. Powered by autonomous AI agents, Superloopz learns from every transaction to automate catalog creation, inventory synchronization, pricing, demand forecasting, and partner coordination.</p>

<p>A large enterprise can connect its entire distributor and retail network into a single intelligent system. At the same time, an individual creator can upload a product image, and our AI generates studio-quality visuals, builds a storefront, and lists the product across marketplaces like Amazon and eBay in one click.</p>

<p>We’ve validated this need with <strong>20+ brands</strong> across FMCG, pharma, jewelry, and quick commerce, with <strong>6 ready for early access</strong>. The opportunity is massive: global B2B e-commerce is ~$32T today and projected to reach <strong>$61.9T by 2030</strong>, India’s D2C market is expected to hit <strong>$61B by 2027</strong>, and U.S. retail e-commerce already exceeds <strong>$1.3T</strong>. Yet no platform exists that truly bridges B2C and B2B end-to-end.</p>

<p>Commerce remains deeply fragmented across marketplaces, distributors, and aggregators. Brands are willing to pay millions for real-time visibility and control because it directly protects revenue — and that’s the gap Superloopz is built to close.</p>

<p>I’d love to share how we’re tackling this and why we believe Superloopz can become the backbone of global commerce over the next decade. I’ve attached a detailed PRD with data and real-world examples for deeper context. Would you be open to a brief conversation?</p>

<p>Best regards,<br>
Praharsha<br>
Co-Founder, Superloopz<br>
+1 949-755-9016 | 
<a href="https://www.linkedin.com/in/praharsha-nelaturi/">My LinkedIn</a> | 
<a href="https://www.linkedin.com/in/chaitanyasai-g/">Co-founder&apos;s LinkedIn</a>
</p>

<p>
<a href="https://superloopz.short.gy/Fundpromo">superloopz.short.gy/Fundpromo</a>
</p>
</body></html>"""

    # Embed tracking pixel into the email body
    import uuid
    cache_buster = uuid.uuid4()
    tracking_pixel = f'<img src="{TRACKER_BASE}/pixel/{track_id}.png?cb={cache_buster}" width="1" height="1"/>'
    body_with_pixel = body + tracking_pixel

    msg.attach(MIMEText(body_with_pixel, 'html'))

    if os.path.exists(ATTACHMENT_PATH):
        try:
            with open(ATTACHMENT_PATH, "rb") as f:
                part = MIMEApplication(f.read(), Name=os.path.basename(ATTACHMENT_PATH))
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(ATTACHMENT_PATH)}"'
            msg.attach(part)
        except Exception as e:
            print(f"⚠️ Failed to attach file {ATTACHMENT_PATH}: {e}")
    else:
        print(f"⚠️ Attachment file not found: {ATTACHMENT_PATH}")

    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    try:
        result = gmail_service.users().messages().send(userId="me", body={'raw': raw_message}).execute()
        gmail_message_id = result.get('id', None)

        # Identify sender email
        try:
            sender_email = gmail_service.users().getProfile(userId="me").execute().get("emailAddress")
        except Exception:
            sender_email = "pnelatur@uci.edu"

        # Register email with tracker (now includes sender_email)
        try:
            requests.post(f"{TRACKER_BASE}/_register_send", json={
                "track_id": track_id,
                "recipient_email": to_email,
                "sender_email": sender_email,
                "subject": EMAIL_SUBJECT,
                "gmail_message_id": gmail_message_id
            }, timeout=5)
        except Exception as tracker_err:
            print(f"⚠️ Tracker registration failed: {tracker_err}")

        print(f"✅ Email sent to {to_email}")
        return True
    except Exception as e:
        print(f"❌ Failed to send email to {to_email}: {e}")
        return False

def update_status(sheets_service, df, row_index, status):
    # Determine the column index for 'Status' dynamically to avoid hardcoding column letters
    try:
        status_col_idx = df.columns.get_loc('Status') + 1  # 1-based for sheet columns
        # convert to column letter (supports up to Z)
        col_letter = chr(64 + status_col_idx)
    except Exception:
        # fallback to column C if anything goes wrong
        col_letter = 'C'
    cell = f"{col_letter}{row_index+2}"
    try:
        sheets_service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=cell,
            valueInputOption="RAW",
            body={"values": [[status]]}
        ).execute()
    except Exception as e:
        print(f"⚠️ Failed to update sheet status at {cell}: {e}")

def process_emails():
    print("🚀 Starting email automation...")
    sheets_service, gmail_service = auth_google_services()
    df, pending_rows = get_pending_rows(sheets_service)
    print(f"✅ Loaded sheet with {len(df)} rows. Pending rows: {len(pending_rows)}")
    if pending_rows.empty:
        print("No pending rows found.")
        return "No pending rows found."
    for idx, row in pending_rows.iterrows():
        name, email = row.get('Name', ''), row.get('Email', '')
        print(f"📧 Sending email to: {email} ({name})")
        success = send_email(gmail_service, name, email)
        update_status(sheets_service, df, idx, "Completed" if success else "Error")
        time.sleep(2)
    print("🎉 Email automation complete.")
    return "Email automation run complete."

# ---- Cloud Function entry point ----
# def send_emails(request):
#     # Cloud Function entry: run automation and return a simple JSON-like dict
#     result = process_emails()
#     return {"status": "ok", "message": result}


if __name__ == "__main__":
    process_emails()
