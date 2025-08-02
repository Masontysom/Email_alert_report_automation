import imaplib
import email
from email.header import decode_header
from datetime import datetime
import os
from dotenv import load_dotenv
import pandas as pd
from bs4 import BeautifulSoup
import re
from collections import defaultdict
from datetime import timedelta
import time
import socket
import sys

# Load credentials from .env file
load_dotenv()
EMAIL = os.getenv("EMAIL_ID")
PASSWORD = os.getenv("APP_PASSWORD")

# Ensure environment variables are loaded
if not EMAIL or not PASSWORD:
    print("Error: EMAIL_ID or APP_PASSWORD not found in .env file.")
    sys.exit()

# Format dates
today = datetime.today().strftime("%Y-%m-%d")
yesterday = (datetime.today() - timedelta(days=1)).strftime("%d-%b-%Y")
tomorrow = (datetime.today() + timedelta(days=1)).strftime("%d-%b-%Y") 
def get_email_body(msg):
    body_content = ""
    html_content_found = False
    if msg.is_multipart():
        for part in msg.walk():
            ctype = part.get_content_type()
            cdispo = str(part.get('Content-Disposition'))
            if ctype == 'text/plain' and 'attachment' not in cdispo:
                try:
                    charset = part.get_content_charset()
                    body_content = part.get_payload(decode=True).decode(charset if charset else "utf-8", errors='ignore')
                    return body_content
                except Exception as e:
                    print(f"Warning: Error decoding plain text part: {e}")
            elif ctype == 'text/html' and 'attachment' not in cdispo:
                try:
                    charset = part.get_content_charset()
                    html_content_found = True
                    body_content = part.get_payload(decode=True).decode(charset if charset else "utf-8", errors='ignore')
                except Exception as e:
                    print(f"Warning: Error decoding HTML part: {e}")
    else:
        try:
            charset = msg.get_content_charset()
            body_content = msg.get_payload(decode=True).decode(charset if charset else "utf-8", errors='ignore')
            if msg.get_content_type() == 'text/html':
                html_content_found = True
        except Exception as e:
            print(f"Warning: Error decoding single part message: {e}")

    if html_content_found and body_content:
        soup = BeautifulSoup(body_content, "html.parser")
        return soup.get_text(separator=' ', strip=True)
    return body_content

print("🔐 Connecting to Gmail...")
try:
    mail = imaplib.IMAP4_SSL("imap.gmail.com")
    mail.login(EMAIL, PASSWORD)
    mail.select("inbox")
    print("✅ Connected and logged in successfully.")
except Exception as e:
    print(f"❌ Error connecting or logging in to Gmail: {e}")
    exit()

print(f"🔍 Fetching emails from {yesterday} to {tomorrow} (last 2 days)...")
try:
    status, messages = mail.search(None, f'(SINCE "{yesterday}" BEFORE "{tomorrow}")')
    email_ids = messages[0].split()
    print(f"📬 Total emails found: {len(email_ids)}")
except Exception as e:
    print(f"❌ Error searching for emails: {e}")
    email_ids = []

unique_daily_alerts = defaultdict(lambda: {"count": 0, "time": None, "device": None,"mail_date": None})

# --- Daily Status Alert function here (unchanged from earlier) ---
def Daily_Status_Alert(mail, email_ids, unique_alerts_dict):
    """
    Parses daily status report emails for critical alerts and populates unique_alerts_dict.
    Uses original logic for HTML parsing, with refined group name extraction.
    """
    print("\n--- Starting Daily Status Alert Parsing ---")
    for i, num in enumerate(email_ids, start=1):
        extracted_mail_date = today  # Default mail date
        print(f"\n➡️  Processing email {i}/{len(email_ids)}")
        time.sleep(0.5)  # To avoid hitting Gmail's rate limits
        try:
            status, data = mail.fetch(num, "(RFC822)")
            msg = email.message_from_bytes(data[0][1])

            subject, encoding = decode_header(msg["Subject"])[0]
            if isinstance(subject, bytes):
                subject = subject.decode(encoding if encoding else "utf-8")

            print(f"   ➤ Subject: {subject}")

            critical_match = re.search(r"Critical:\s*(\d+)", subject, re.IGNORECASE)
            critical_count = int(critical_match.group(1)) if critical_match else 0

            # --- Group name extraction for Daily Status to get content after LAST '>' ---
            group_match = re.search(r"group:\s*(?:[^>]*(>)\s*)?([^\)]+)\)", subject)
            if group_match:
                if group_match.group(1):
                    group_name = group_match.group(2).strip()
                else:
                    group_name = group_match.group(2).strip()
            else:
                group_name = "-"
            
            group_name = re.sub(r'[^\w\s().&-]', '', group_name).strip()
            print(f"      🏢 Extracted Group Name: '{group_name}'")


            if "DAILY STATUS REPORT" in subject.upper() and critical_count > 0:
                # Extract Mail Date from Subject
                mail_date_match = re.search(r"ON\s+([A-Za-z]{3,9} \d{1,2}, \d{4})", subject)
                if mail_date_match:
                    extracted_mail_date_str = mail_date_match.group(1)
                    try:
                        extracted_mail_date = datetime.strptime(extracted_mail_date_str, "%b %d, %Y").strftime("%Y-%m-%d")
                    except ValueError:
                        extracted_mail_date = today  # fallback
                else:
                    extracted_mail_date = today  # fallback
                print(f"      🗓️ Extracted Mail Date: {extracted_mail_date}")

                print(f"   ✅ {critical_count} critical alert(s) found — parsing content...")

                html_content = None
                for part in msg.walk():
                    if part.get_content_type() == "text/html":
                        html_content = part.get_payload(decode=True).decode(errors="ignore")
                        break

                if not html_content:
                    print("   ⚠️ No HTML content found — skipping.")
                    continue

                soup = BeautifulSoup(html_content, "html.parser")
                all_divs = soup.find_all("div")

                current_time = None
                last_message = None
                next_is_device = False

                for div in all_divs:
                    text = div.get_text(strip=True)

                    if re.search(r"\w{3,9} \d{1,2}, \d{4}, \d{1,2}:\d{2}", text):
                        current_time = text

                    elif text.lower() == "device":
                        next_is_device = True

                    elif next_is_device and text:
                        device = text
                        next_is_device = False
                        if last_message:
                            unique_key = (group_name, last_message)
                            current_alert = unique_alerts_dict[unique_key]

                            if current_alert["count"] == 0:
                                print(f"   📌 New alert: {group_name} → {last_message[:60]}...")
                                current_alert["time"] = current_time if current_time else "-"
                                current_alert["device"] = device
                                current_alert["mail_date"] = extracted_mail_date
                            else:
                                current_alert["count"] += 1

                                try:
                                    existing_dt = datetime.strptime(current_alert["mail_date"], "%Y-%m-%d")
                                    new_dt = datetime.strptime(extracted_mail_date, "%Y-%m-%d")
                                    if new_dt > existing_dt:
                                        current_alert["mail_date"] = extracted_mail_date
                                        current_alert["time"] = current_time if current_time else current_alert["time"]
                                        current_alert["device"] = device or current_alert["device"]
                                except:
                                    pass  # Graceful fallback
                            last_message = None
                    elif any(key in text.lower() for key in ["offline","azure", "quota ", "backup failed", "incident detected","expired","exceeded"]):
                        last_message = text
            else:
                print("   ❌ Skipped — not a daily status report or 0 critical alerts.")

        except socket.error as e:
            print(f"🔌 Connection dropped: {e} — attempting to reconnect to Gmail...")
            try:
                if mail.state != 'LOGOUT':
                    
                    mail.logout()
            except:
                pass  # Ignore if logout fails

            time.sleep(3)
            try:
                mail = imaplib.IMAP4_SSL("imap.gmail.com")
                mail.login(EMAIL, PASSWORD)
                mail.select("inbox")
                print("🔁 Reconnected successfully.")
            except Exception as reconnect_err:
                print(f"❌ Reconnection failed: {reconnect_err}")
                continue  # Skip this email and move on

        except Exception as e:
            print(f"   ❌ General error processing email {num}: {e}")
            continue
    print("--- Finished Daily Status Alert Parsing ---")


# --- Main Execution Starts Here ---
Daily_Status_Alert(mail, email_ids, unique_daily_alerts)

# Convert parsed data
daily_status_rows = []
for (group, message), values in unique_daily_alerts.items():
    alert_dt_str = values["time"]
    try:
        alert_dt = datetime.strptime(alert_dt_str, "%b %d, %Y, %I:%M:%S %p")
        mail_date_str = alert_dt.strftime("%Y-%m-%d")
        # Filter: Only include today or yesterday
        if mail_date_str not in [today, (datetime.today() - timedelta(days=1)).strftime("%Y-%m-%d")]:
            continue
    except Exception:
        continue  # Skip invalid/missing dates

    daily_status_rows.append({
        "Mail Date": values["mail_date"] or today,
        "Alerts Date & Time": values["time"],
        "Customers": group,
        "Device": values["device"] if values["device"] else "-",
        "Message": message
    })


daily_df = pd.DataFrame(daily_status_rows)

# Save to Excel
# Save to Excel
output_dir = r"\\K2\g\maildocuments\Acronics_Daily_Report"
os.makedirs(output_dir, exist_ok=True)
base_filename = f"Daily_Status_Alerts_Report_{datetime.today().strftime('%Y%m%d')}"
output_file = os.path.join(output_dir, base_filename + ".xlsx")

# If file exists, append (1), (2), ...
file_counter = 1
while os.path.exists(output_file):
    output_file = os.path.join(output_dir, f"{base_filename} ({file_counter}).xlsx")
    file_counter += 1

print(f"\n✍️ Generating Excel report: {output_file}...")
try:
    with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
        workbook = writer.book
        center_align_format = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
        left_align_format = workbook.add_format({'align': 'left', 'valign': 'top', 'text_wrap': True})
        header_format = workbook.add_format({'bold': True, 'bg_color': '#D9E1F2', 'align': 'center', 'valign': 'vcenter'})

        sheet_name = 'Alerts'
        worksheet = workbook.add_worksheet(sheet_name)
        writer.sheets[sheet_name] = worksheet

        # Write column headers
        for col_num, header in enumerate(daily_df.columns):
            worksheet.write(0, col_num, header, header_format)

        # Write data rows with spacing between customers
        current_row = 1
        previous_customer = None

        for _, row in daily_df.iterrows():
            current_customer = row["Customers"]

            # Insert empty row between different groups
            if previous_customer is not None and current_customer != previous_customer:
                current_row += 1

            for col_num, value in enumerate(row):
                format_to_use = left_align_format if daily_df.columns[col_num].lower() == "message" else center_align_format
                worksheet.write(current_row, col_num, value, format_to_use)

            previous_customer = current_customer
            current_row += 1

        # Set column widths
        for col_num, header in enumerate(daily_df.columns):
            if header.lower() != "message":
                worksheet.set_column(col_num, col_num, 20, center_align_format)
            else:
                worksheet.set_column(col_num, col_num, 70, left_align_format)

        # Apply conditional formatting to Message column
        message_column_letter = 'E'  # Adjust if your column order changes
        worksheet.conditional_format(
            f'{message_column_letter}2:{message_column_letter}{current_row}',
            {'type': 'text', 'criteria': 'containing', 'value': 'failed',
             'format': workbook.add_format({'bg_color': "#EC283F"})}
        )
        worksheet.conditional_format(
            f'{message_column_letter}2:{message_column_letter}{current_row}',
            {'type': 'text', 'criteria': 'containing', 'value': 'quota',
             'format': workbook.add_format({'bg_color': "#FFC7C7"})}
        )
        worksheet.conditional_format(
            f'{message_column_letter}2:{message_column_letter}{current_row}',
            {'type': 'text', 'criteria': 'containing', 'value': 'Incident',
             'format': workbook.add_format({'bg_color': "#72F072"})}
        )

        print("   ✅ Daily Status Alerts formatted with spacing.")
          # --- Summary Sheet ---
        summary_data = [
            ["Total Emails Processed", len(email_ids)],
            ["Total Daily Status Alerts", len(daily_df)],
            ["Report Generated On", datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
        ]
        summary_df = pd.DataFrame(summary_data, columns=["Metric", "Value"])
        summary_df.to_excel(writer, index=False, sheet_name="Summary")

        summary_ws = writer.sheets["Summary"]
        summary_ws.set_column("A:A", 30, center_align_format)
        summary_ws.set_column("B:B", 25, center_align_format)
    print(f"\n✅ Excel report generated successfully: {output_file}")
except Exception as e:
    print(f"❌ Error saving Excel file: {e}")

try:
    mail.logout()
    print("🔚 Logged out from Gmail.")
except Exception as e:
    print(f"⚠️ Could not logout cleanly: {e}")

