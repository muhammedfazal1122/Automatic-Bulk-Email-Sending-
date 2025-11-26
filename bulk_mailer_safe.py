# bulk_mailer_safe.py
# Usage:
# 1) preview only (no sends) : leave PREVIEW_MODE = True and run `python bulk_mailer_safe.py`
# 2) single real test send  : set PREVIEW_MODE = False and TEST_EMAIL = "fasal@nuronics.com" then run
# 3) full send              : set PREVIEW_MODE = False and TEST_EMAIL = "" (empty) then run

import os
import time
import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT   = int(os.getenv("SMTP_PORT", "587"))
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")

# CONFIG - change these
CSV_FILE = "emails_clean.csv"        # your cleaned CSV
TEMPLATE_FILE = "email_template.txt"
SEND_RATE_PER_MIN = 10               # safe default (e.g. 10 emails per minute)
PREVIEW_MODE = False
TEST_EMAIL = ""           # empty string
LOG_FILE = "send_log.csv"
SUBJECT = "Application — Python / Django + React Developer"


# add these imports at the top of the file (if not already present)
from email.mime.base import MIMEBase
from email import encoders

# set this near the CONFIG area
ATTACH_PATH = "resume.pdf"    # change if your file has a different name


# ---------------- helpers ----------------
def load_template(path):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return "Dear Hiring Manager,\n\n(Template missing)\n"



# ---------------- new send function with attachment ----------------
def send_via_smtp(smtp_conn, sender, to_email, subject, body):
    # create multipart message
    msg = MIMEMultipart()
    msg["From"] = sender
    msg["To"] = to_email
    msg["Subject"] = subject

    # attach the plain text body
    msg.attach(MIMEText(body, "plain"))

    # attach resume if file exists
    if os.path.exists(ATTACH_PATH):
        try:
            part = MIMEBase("application", "octet-stream")
            with open(ATTACH_PATH, "rb") as f:
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                f'attachment; filename="{os.path.basename(ATTACH_PATH)}"'
            )
            msg.attach(part)
        except Exception as e:
            # if attachment fails, raise so caller can log it
            raise Exception(f"Attachment error: {e}")

    # send the email
    smtp_conn.sendmail(sender, to_email, msg.as_string())


# ---------------- main ----------------
def main():
    template = load_template(TEMPLATE_FILE)

    # load CSV
    df = pd.read_csv(CSV_FILE)
    df.fillna("", inplace=True)

    logs = []

    smtp = None
    if not PREVIEW_MODE:
        # open SMTP connection
        smtp = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        smtp.ehlo()
        smtp.starttls()
        smtp.login(SENDER_EMAIL, APP_PASSWORD)

    delay = 60.0 / SEND_RATE_PER_MIN if SEND_RATE_PER_MIN > 0 else 0

    # If TEST_EMAIL specified and PREVIEW_MODE is False, we will send only one email to TEST_EMAIL.
    if TEST_EMAIL and not PREVIEW_MODE:
        to = TEST_EMAIL.strip()
        print(">>> TEST MODE: sending single email to", to)
        body = template
        try:
            send_via_smtp(smtp, SENDER_EMAIL, to, SUBJECT, body)
            logs.append({"email": to, "status": "sent_test", "error": "", "timestamp": datetime.utcnow().isoformat()})
            print("Sent test email to", to)
        except Exception as e:
            logs.append({"email": to, "status": "error_test", "error": str(e), "timestamp": datetime.utcnow().isoformat()})
            print("Error sending test email:", e)
        if smtp:
            smtp.quit()
        pd.DataFrame(logs).to_csv(LOG_FILE, index=False)
        print("Test finished. Log:", LOG_FILE)
        return

    # Normal loop (preview or full send)
    for idx, row in df.iterrows():
        to_email = row.get("Email", "").strip()
        if not to_email:
            logs.append({"email": "", "status": "skipped", "error": "missing email", "timestamp": datetime.utcnow().isoformat()})
            continue

        body = template  # no merge tags needed

        if PREVIEW_MODE:
            print("------ PREVIEW ------")
            print("To:", to_email)
            print("Subject:", SUBJECT)
            print(body[:800])
            print("---------------------\n")
            logs.append({"email": to_email, "status": "preview", "error": "", "timestamp": datetime.utcnow().isoformat()})
            continue

        # actual send
        try:
            send_via_smtp(smtp, SENDER_EMAIL, to_email, SUBJECT, body)
            logs.append({"email": to_email, "status": "sent", "error": "", "timestamp": datetime.utcnow().isoformat()})
            print("Sent →", to_email)
        except Exception as e:
            logs.append({"email": to_email, "status": "error", "error": str(e), "timestamp": datetime.utcnow().isoformat()})
            print("Error sending to", to_email, ":", e)

        if delay > 0:
            time.sleep(delay)

    if smtp:
        smtp.quit()

    pd.DataFrame(logs).to_csv(LOG_FILE, index=False)
    print("Done. Log saved to", LOG_FILE)

if __name__ == "__main__":
    main()
