import smtplib
import imaplib
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import time
import traceback
from googleapiclient.discovery import build
import json
import os

# === SMTP/IMAP Credentials ===
print("✅ Script loaded. Starting authentication...",flush=True)
SMTP_SERVER = "mail.b2bgrowthexpo.com"
SMTP_PORT = 587
SMTP_EMAIL = "speakersengagement@b2bgrowthexpo.com"
SMTP_PASSWORD = "jH!Ra[9q[f68"

IMAP_SERVER = "mail.b2bgrowthexpo.com"
IMAP_PORT = 143
IMAP_EMAIL = SMTP_EMAIL
IMAP_PASSWORD = SMTP_PASSWORD

SENDER_NAME = "Nagendra Mishra"

# === HTML Email Template ===
EMAIL_TEMPLATE = """
<html>
  <body style="font-family: Arial, sans-serif; font-size: 15px; color: #333; background-color: #ffffff; padding: 20px;">
    <div style="text-align: center; margin-bottom: 20px;">
      <img src="https://iili.io/FogC9l2.jpg" alt="B2B Growth Expo" style="max-width: 400px; height: auto;" />
    </div>
    <p>Hi {%name%},</p>
    <p>{%body%}</p>
    <p>
      If you would like to schedule a meeting with me at your convenient time,<br>
      please use the link below:<br>
      <a href="https://tidycal.com/nagendra/b2b-discovery-call" target="_blank">https://tidycal.com/nagendra/b2b-discovery-call</a>
    </p>
    <p style="margin-top: 30px;">
      Thanks & Regards,<br>
      <strong>Nagendra Mishra</strong><br>
      Director | B2B Growth Hub<br>
      Mo: +44 7913 027482<br>
      Email: <a href="mailto:nagendra@b2bgrowthexpo.com">nagendra@b2bgrowthexpo.com</a><br>
      <a href="https://www.b2bgrowthexpo.com" target="_blank">www.b2bgrowthexpo.com</a>
    </p>
    <p style="font-size: 13px; color: #888;">
      If you don’t want to hear from me again, please let me know.
    </p>
  </body>
</html>
"""

# === Authenticate Google Sheets ===
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file('/etc/secrets/service_account.json', scopes=SCOPES)
sheets_api = build("sheets", "v4", credentials=creds)
gc = gspread.authorize(creds)
print("✅ Google Sheets credentials loaded.",flush=True)
print("✅ Opening worksheet...",flush=True)
sheet = gc.open("Expo-Sales-Management").worksheet("OB-speakers")
print(f"✅ Worksheet '{sheet.title}' opened successfully.",flush=True)

# === Follow-up Templates ===
FOLLOWUP_EMAILS = [
    "I hope you’re doing well.I’m reaching out to invite you to speak at the upcoming {%show%}, an event that brings together founders, business leaders, and professionals for a day of meaningful connections and insight sharing.We’d be honoured to have you as one of our speakers. While this is an unpaid opportunity, here’s what you can expect:<ul><li>A platform to showcase your expertise to a high-quality audience</li><li>Increased visibility within your industry</li><li>Opportunities to expand your professional network</li></ul>Our previous expos have featured startup founders, SME owners, and decision-makers from various sectors. It’s a fantastic opportunity to build presence and influence in a focused business environment.",
    "I just wanted to follow up on my previous message about speaking at the upcoming {%show%}.We’d love to have you share your expertise with our audience of founders, professionals, and decision-makers from across industries.",
    "We’re finalizing the speaker lineup for the {%show%} and I didn’t want you to miss the chance to be part of it.It’s a high-visibility opportunity to connect with a professional audience and share your knowledge. While it’s not a paid spot, the exposure and networking can be incredibly valuable.",
    "Just wanted to check in one more time in case the timing works out for you to join us as a speaker at the {%show%}.We’re seeing strong interest from founders, professionals, and decision-makers across the region, and I’d love to include your voice in the mix.Thanks again for your time and consideration."
]

FOLLOWUP_SUBJECTS = [
    "Invitation to Speak at the {%show%}",
    "Still Time to Join as a Speaker at {%show%}",
    " We’d Love to Include You – {%show%} Speaker Invite",
    "Reminder: We’d Love You to Speak at {%show%}"
]

FINAL_EMAIL = (
    "Just one last note,if you're still interested in speaking at the {%show%}, we’d love to include you. We're locking the final agenda this week."
    "Even if the timing doesn’t work this time, I’d be happy to stay in touch for future opportunities."
    "Either way, appreciate your time!"
)

def send_email(to_email, subject, body, name=""):
    print(f"Preparing to send email to: {to_email}")
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = f"{SENDER_NAME} <{SMTP_EMAIL}>"
    msg["To"] = to_email
    html_body = EMAIL_TEMPLATE.replace("{%name%}", name).replace("{%body%}", body)
    msg.attach(MIMEText(html_body, "html"))

    try:
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.starttls()
            server.login(SMTP_EMAIL, SMTP_PASSWORD)
            server.sendmail(SMTP_EMAIL, to_email, msg.as_string())
        print(f"✅ Email sent to {to_email}",flush=True)
    except Exception as e:
        print(f"❌ SMTP Error while sending to {to_email}: {e}")

    try:
        imap = imaplib.IMAP4_SSL(IMAP_SERVER)
        imap.login(SMTP_EMAIL, SMTP_PASSWORD)
        imap.append("INBOX.Sent", '', imaplib.Time2Internaldate(time.time()), msg.as_bytes())
        imap.logout()
    except Exception as e:
        print(f"❌ IMAP Error while saving to Sent folder for {to_email}: {e}")

def get_reply_emails():
    print("Checking for new replies in inbox...")
    replied = set()
    try:
        with imaplib.IMAP4_SSL(IMAP_SERVER) as mail:
            mail.login(IMAP_EMAIL, IMAP_PASSWORD)
            mail.select("inbox")
            status, messages = mail.search(None, 'UNSEEN')
            for num in messages[0].split():
                _, data = mail.fetch(num, "(RFC822)")
                msg = email.message_from_bytes(data[0][1])
                from_addr = email.utils.parseaddr(msg["From"])[1].lower().strip()
                replied.add(from_addr)
    except Exception as e:
        print(f"❌ IMAP Error while checking replies: {e}")
    print(f"✅ Found {len(replied)} new replies.",flush=True)
    return replied

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return {
        "red": int(hex_color[0:2], 16) / 255,
        "green": int(hex_color[2:4], 16) / 255,
        "blue": int(hex_color[4:6], 16) / 255
    }

def get_all_row_colors(sheet_id, sheet_name, start_row=2, end_row=1000):
    try:
        range_ = f"{sheet_name}!A{start_row}:A{end_row}"
        result = sheets_api.spreadsheets().get(
            spreadsheetId=sheet_id,
            includeGridData=True,
            ranges=[range_],
            fields="sheets.data.rowData.values.effectiveFormat.backgroundColor"
        ).execute()

        rows = result['sheets'][0]['data'][0].get('rowData', [])
        row_colors = []

        for row in rows:
            cell = row.get('values', [{}])[0]
            color = cell.get('effectiveFormat', {}).get('backgroundColor', {})
            rgb = (
                int(color.get('red', 1) * 255),
                int(color.get('green', 1) * 255),
                int(color.get('blue', 1) * 255)
            )
            row_colors.append(rgb)

        # Padding to avoid index errors if fewer rows were returned
        expected_rows = end_row - start_row + 1
        while len(row_colors) < expected_rows:
            row_colors.append((255, 255, 255))  # Default to white

        return row_colors

    except Exception as e:
        print(f"❌ Failed to fetch all row colors: {e}")
        return []

def batch_update_cells(sheet_id, updates, chunk_size=100):
    try:
        for i in range(0, len(updates), chunk_size):
            chunk = updates[i:i+chunk_size]
            body = {
                "valueInputOption": "USER_ENTERED",
                "data": chunk
            }
            sheets_api.spreadsheets().values().batchUpdate(
                spreadsheetId=sheet_id,
                body=body
            ).execute()
            print(f"✅ Batch update for rows {i+1} to {i+len(chunk)} complete.")
            time.sleep(1)  # Slight delay to avoid rate limits
    except Exception as e:
        print(f"❌ Failed batch cell update: {e}")

def batch_color_rows(spreadsheet_id, start_row_index_color_map, sheet_id):
    requests = []
    for row_idx, hex_color in start_row_index_color_map.items():
        rgb = hex_to_rgb(hex_color)
        print(f"🎨 Coloring row {row_idx} with {hex_color} => RGB {rgb}")
        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": row_idx - 1,
                    "endRowIndex": row_idx,
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": rgb
                    }
                },
                "fields": "userEnteredFormat.backgroundColor"
            }
        })

    try:
        response = sheets_api.spreadsheets().batchUpdate(
            spreadsheetId=spreadsheet_id, body={"requests": requests}
        ).execute()
        print(f"✅ Batch row coloring done. Response: {json.dumps(response, indent=2)}",flush=True)
    except Exception as e:
        print(f"❌ Batch row coloring failed: {e}",flush=True)

def set_row_color(sheet, row_number, color_hex):
    print(f"Coloring row {row_number} with color {color_hex}")
    try:
        sheet_format = {
            "requests": [{
                "repeatCell": {
                    "range": {
                        "sheetId": sheet._properties['sheetId'],
                        "startRowIndex": row_number - 1,
                        "endRowIndex": row_number,
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColorStyle": {
                                "rgbColor": hex_to_rgb(color_hex)
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColorStyle"
                }
            }]
        }
        sheet.spreadsheet.batch_update(sheet_format)
    except Exception as e:
        print(f"❌ Google Sheets Error while coloring row {row_number}: {e}",flush=True)

def get_row_background_color(sheet_id, sheet_name, row_number):
    try:
        range_ = f"{sheet_name}!A{row_number}"
        result = sheets_api.spreadsheets().get(
            spreadsheetId=sheet_id,
            ranges=[range_],
            fields="sheets.data.rowData.values.effectiveFormat.backgroundColor"
        ).execute()

        cell_format = result['sheets'][0]['data'][0]['rowData'][0]['values'][0]['effectiveFormat']['backgroundColor']
        rgb = (
            int(cell_format.get('red', 0) * 255),
            int(cell_format.get('green', 0) * 255),
            int(cell_format.get('blue', 0) * 255)
        )
        print(f"Row {row_number} color fetched: RGB{rgb}",flush=True)
        return rgb
    except Exception as e:
        print(f"❌ Error getting background color for row {row_number}: {e}",flush=True)
        return None

def process_replies():
    print("Processing replies...")
    try:
        data = sheet.get_all_records()
        replied_emails = get_reply_emails()
        updates = []
        color_updates = {}

        row_colors = get_all_row_colors(sheet.spreadsheet.id, sheet.title, 2, len(data) + 1)
        for idx, row in enumerate(data, start=2):
            if not any(row.values()):
                continue

            email_addr = row.get("Email", "").lower().strip()
            if not email_addr or row.get("Reply Status", "") == "Replied":
                continue

            rgb = row_colors[idx - 2]
            if rgb and rgb != (255, 255, 255):
                continue

            if email_addr in replied_emails:
                updates.append({
                    "range": f"{sheet.title}!G{idx}",
                    "values": [["Replied"]]
                })
                color_updates[idx] = "#FFFF00"

        if updates:
            batch_update_cells(sheet.spreadsheet.id, updates)
        if color_updates:
            batch_color_rows(sheet.spreadsheet.id, color_updates, sheet._properties['sheetId'])

    except Exception as e:
        print("❌ Error in processing replies:", e)

def process_followups():
    print("Processing follow-up emails...")
    try:
        data = sheet.get_all_records()
        today = datetime.today().strftime('%Y-%m-%d')
        updates = []
        color_updates = {}

        row_colors = get_all_row_colors(sheet.spreadsheet.id, sheet.title, 2, len(data) + 1)
        sent_tracker = set()

        for idx, row in enumerate(data, start=2):
            if not any(row.values()):
                continue

            rgb = row_colors[idx - 2]
            if rgb and rgb != (255, 255, 255):
                continue

            email_addr = row.get("Email", "").lower().strip()
            if not email_addr or email_addr in sent_tracker:
                continue

            name = row.get("First_Name", "").strip()

            try:
                count = int(row.get("Follow-Up Count"))
                if count < 0:
                    count = 0
            except:
                count = 0

            last_date = row.get("Last Follow-Up Date", "")
            reply_status = row.get("Reply Status", "").strip()

            if reply_status in ["Replied", "No Reply After 4 Followups"]:
                continue

            if last_date:
                last_dt = datetime.strptime(last_date, "%Y-%m-%d")
                if (datetime.now() - last_dt).total_seconds() < 86400:
                    continue

            if count >= 4:
                send_email(email_addr, "Last Reminder: Speaker Spot!", FINAL_EMAIL, name=name)
                updates.append({"range": f"{sheet.title}!G{idx}", "values": [["No Reply After 4 Followups"]]})
                color_updates[idx] = "#FF0000"
                sent_tracker.add(email_addr)
                continue

            next_count = count

            try:
                followup_text = FOLLOWUP_EMAILS[next_count].replace("{%name%}", name)
                subject = FOLLOWUP_SUBJECTS[next_count]

                if next_count == 0:
                    show = row.get("Show", "").strip()
                    if not show:
                        continue
                    followup_text = followup_text.replace("{%show%}", show)
                    subject = subject.replace("{%show%}", show)

                send_email(email_addr, subject, followup_text, name=name)
                sent_tracker.add(email_addr)

                print(f"Row {idx}: Sent template {next_count + 1} to {email_addr}")

                updates.extend([
                    {"range": f"{sheet.title}!E{idx}", "values": [[str(next_count + 1)]]},
                    {"range": f"{sheet.title}!F{idx}", "values": [[today]]},
                    {"range": f"{sheet.title}!G{idx}", "values": [["Pending"]]}
                ])

            except Exception as e:
                print(f"❌ Failed to prepare/send follow-up email to {email_addr}: {e}")
                continue

            if (idx - 1) % 3 == 0:
                print("Sleeping for 3 seconds...")
                time.sleep(3)

        if updates:
            batch_update_cells(sheet.spreadsheet.id, updates)
        if color_updates:
            batch_color_rows(sheet.spreadsheet.id, color_updates, sheet._properties['sheetId'])

    except Exception as e:
        print("❌ Error in processing followups:", e)

# === Entry Point ===
if __name__ == "__main__":
    print("🚀 Sales follow-up automation started...",flush=True)
    next_followup_check = time.time()

    while True:
        try:
            print(f"⏱ Loop at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",flush=True)
            #print("\n--- Checking for replies ---",flush=True)
            #process_replies()

            current_time = time.time()
            if current_time >= next_followup_check:
                print("\n--- Sending follow-up emails ---",flush=True)
                process_followups()
                next_followup_check = current_time + 86400  # every 24 hours

        except Exception:
            print("❌ Fatal error:",flush=True)
            traceback.print_exc()

        time.sleep(30)
