# bulk_mailer_safe.py — send the cover letter + resume to every HR contact in a CSV.
#
# Usage (from the project folder):
#   ./venv/bin/python bulk_mailer_safe.py                     # PREVIEW only, nothing is sent
#   ./venv/bin/python bulk_mailer_safe.py --test you@x.com    # send ONE real email to yourself
#   ./venv/bin/python bulk_mailer_safe.py --send              # send to everyone not already sent
#   ./venv/bin/python bulk_mailer_safe.py --send --limit 50   # send to the next 50 only
#
# Safety:
#   * Anyone already marked "sent" in send_log.csv is skipped, so re-running never double-sends.
#   * Each result is appended to send_log.csv the moment it happens, so Ctrl+C or a crash
#     loses nothing; just run the same command again to continue where it stopped.

import argparse
import csv
import os
import re
import smtplib
import sys
import time
from datetime import datetime, timezone
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

SMTP_SERVER = os.getenv("SMTP_SERVER", "smtp.gmail.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")

# ---------------- CONFIG ----------------
CSV_FILE = "Gulf IT Hiring Contacts - Gulf HR Contacts.csv"   # downloaded from the Google Sheet
TEMPLATE_FILE = "email_template.txt"
ATTACH_PATH = "resume.pdf"
LOG_FILE = "send_log.csv"
SENDER_NAME = "Muhammed Fasal C"
SUBJECT = "Senior Python / Django Developer – Muhammed Fasal C – Resume Attached"
SEND_RATE_PER_MIN = 10        # 6 s between emails; keeps Gmail happy
MAX_RETRIES = 2               # per email, after a dropped connection

# Column names in the sheet (matched case-insensitively; only Email is required)
EMAIL_COL = "Email"
COMPANY_COL = "Company / hiring firm"
NAME_COL = "HR / recruiter name"

LOG_FIELDS = ["email", "company", "hr_name", "status", "error", "timestamp"]
EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+(\.[\w-]+)+$")


# ---------------- helpers ----------------
def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


PLACEHOLDER_RE = re.compile(r"^(not named|not stated|n/?a|none|unknown|-+)\b", re.I)


def clean_company(raw):
    """'Touchworld Technology (via Netlyft)' -> 'Touchworld Technology';
    '(not named – personal email)' -> ''  (empty = no company known)."""
    s = re.sub(r"\s*[\(\[].*?[\)\]]", "", raw or "")   # drop (…) and […] notes
    s = " ".join(s.split()).strip(" ,;-–|")
    return "" if PLACEHOLDER_RE.match(s) else s


def find_col(df, wanted):
    for c in df.columns:
        if c.strip().lower() == wanted.lower():
            return c
    return None


def load_contacts(path):
    """Return a list of dicts {email, company, hr_name}, cleaned and de-duplicated."""
    df = pd.read_csv(path, encoding="utf-8-sig", dtype=str).fillna("")
    email_col = find_col(df, EMAIL_COL) or find_col(df, "email")
    if not email_col:
        sys.exit(f"ERROR: no '{EMAIL_COL}' column in {path}. Columns are: {list(df.columns)}")
    company_col = find_col(df, COMPANY_COL)
    name_col = find_col(df, NAME_COL)

    contacts, seen, invalid = [], set(), []
    for _, row in df.iterrows():
        email = row[email_col].strip().lower()
        if not email:
            continue
        if not EMAIL_RE.match(email):
            invalid.append(email)
            continue
        if email in seen:
            continue
        seen.add(email)
        contacts.append({
            "email": email,
            "company": clean_company(row[company_col]) if company_col else "",
            "hr_name": row[name_col].strip() if name_col else "",
        })
    for e in invalid:
        print(f"  ! skipping invalid address: {e}")
    return contacts


def read_log_rows(path):
    """Yield (header, row-dict) for every data row. The log may hold several
    header blocks (old 4-column format, then the new one), blank lines and
    separators; each row is read with the header above it."""
    if not os.path.exists(path):
        return
    header = None
    with open(path, newline="", encoding="utf-8") as f:
        for row in csv.reader(f):
            if not row or not any(cell.strip() for cell in row):
                continue
            if row[0].strip().lower() == "email":
                header = [h.strip() for h in row]
                continue
            if header:
                yield header, dict(zip(header, row))


def load_already_sent(path):
    """Emails never to send again: 'sent', or 'rejected' (address doesn't exist)."""
    return {
        (r.get("email") or "").strip().lower()
        for _, r in read_log_rows(path)
        if (r.get("status") or "").strip() in ("sent", "rejected")
    }


def last_log_header(path):
    header = None
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8") as f:
            for row in csv.reader(f):
                if row and row[0].strip().lower() == "email":
                    header = [h.strip() for h in row]
    return header


class Log:
    """Appends one row per result and flushes immediately."""

    def __init__(self, path):
        self.path = path
        need_header = last_log_header(path) != LOG_FIELDS
        if os.path.exists(path) and os.path.getsize(path) > 0:
            with open(path, "rb") as r:
                r.seek(-1, os.SEEK_END)
                ends_with_newline = r.read(1) == b"\n"
        else:
            ends_with_newline = True
        self.f = open(path, "a", newline="", encoding="utf-8")
        if not ends_with_newline:
            self.f.write("\n")
        if need_header:
            # old 4-column log (or a new file): keep it, start a new header block below
            if os.path.getsize(path) > 0:
                self.f.write("\n")
            self.f.write(",".join(LOG_FIELDS) + "\n")
        self.w = csv.DictWriter(self.f, fieldnames=LOG_FIELDS, extrasaction="ignore")

    def write(self, contact, status, error=""):
        self.w.writerow({**contact, "status": status, "error": error, "timestamp": now()})
        self.f.flush()

    def close(self):
        self.f.close()


def build_message(to_email, body, attachment):
    msg = MIMEMultipart()
    msg["From"] = formataddr((SENDER_NAME, SENDER_EMAIL))
    msg["To"] = to_email
    msg["Subject"] = SUBJECT
    msg.attach(MIMEText(body, "plain", "utf-8"))
    if attachment:
        part = MIMEBase("application", "pdf")
        part.set_payload(attachment)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", "attachment",
                        filename=os.path.basename(ATTACH_PATH))
        msg.attach(part)
    return msg


def connect():
    smtp = smtplib.SMTP(SMTP_SERVER, SMTP_PORT, timeout=60)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(SENDER_EMAIL, APP_PASSWORD)
    return smtp


def send_with_retry(smtp, msg, to_email):
    """Send; reconnect after a dropped connection. Returns the (possibly new) connection."""
    for attempt in range(MAX_RETRIES + 1):
        try:
            smtp.sendmail(SENDER_EMAIL, [to_email], msg.as_string())
            return smtp
        # only real connection drops are retried; SMTP errors such as a
        # rejected address (also OSError subclasses) go straight to the caller
        except (smtplib.SMTPServerDisconnected, ConnectionError, TimeoutError):
            if attempt == MAX_RETRIES:
                raise
            print("    connection dropped, reconnecting…")
            time.sleep(5)
            smtp = connect()


# ---------------- main ----------------
def main():
    p = argparse.ArgumentParser(description="Send resume + cover letter to HR contacts.")
    mode = p.add_mutually_exclusive_group()
    mode.add_argument("--send", action="store_true", help="really send to all pending contacts")
    mode.add_argument("--test", metavar="EMAIL", help="send one real email to EMAIL only")
    p.add_argument("--limit", type=int, default=0, help="send at most N emails this run")
    p.add_argument("--csv", default=CSV_FILE, help=f"contacts CSV (default: {CSV_FILE})")
    args = p.parse_args()

    if not os.path.exists(TEMPLATE_FILE):
        sys.exit(f"ERROR: template {TEMPLATE_FILE} not found")
    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        body = f.read()

    attachment = None
    if os.path.exists(ATTACH_PATH):
        with open(ATTACH_PATH, "rb") as f:
            attachment = f.read()
    elif args.send or args.test:
        sys.exit(f"ERROR: {ATTACH_PATH} not found — refusing to send without the resume")

    if (args.send or args.test) and not (SENDER_EMAIL and APP_PASSWORD):
        sys.exit("ERROR: SENDER_EMAIL / APP_PASSWORD missing from .env")

    # ---- single test email ----
    if args.test:
        to = args.test.strip().lower()
        print(f">>> TEST: sending one email to {to}")
        smtp = connect()
        send_with_retry(smtp, build_message(to, body, attachment), to)
        smtp.quit()
        print("Sent. Check that inbox (test sends are not written to the log).")
        return

    contacts = load_contacts(args.csv)
    already = load_already_sent(LOG_FILE)
    pending = [c for c in contacts if c["email"] not in already]
    if args.limit > 0:
        pending = pending[: args.limit]

    delay = 60.0 / SEND_RATE_PER_MIN if SEND_RATE_PER_MIN > 0 else 0
    done = sum(c["email"] in already for c in contacts)
    print(f"Contacts in CSV: {len(contacts)} | already sent/rejected: {done} "
          f"| to send this run: {len(pending)}")

    # ---- preview ----
    if not args.send:
        print(f"\nSubject: {SUBJECT}\nAttachment: {ATTACH_PATH if attachment else 'MISSING!'}\n")
        print(body)
        print("------ recipients ------")
        for i, c in enumerate(pending, 1):
            print(f"{i:4}. {c['email']:45} {c['company']}")
        mins = len(pending) * delay / 60
        print(f"\nPREVIEW ONLY — nothing was sent. "
              f"To send these {len(pending)} emails (~{mins:.0f} min), run:\n"
              f"  ./venv/bin/python bulk_mailer_safe.py --send")
        return

    if not pending:
        print("Nothing to send — everyone in the CSV is already in the log as sent.")
        return

    # ---- real send ----
    log = Log(LOG_FILE)
    smtp = connect()
    sent = failed = 0
    try:
        for i, c in enumerate(pending, 1):
            try:
                smtp = send_with_retry(smtp, build_message(c["email"], body, attachment), c["email"])
                log.write(c, "sent")
                sent += 1
                print(f"[{i}/{len(pending)}] Sent → {c['email']}  ({c['company']})")
            except smtplib.SMTPRecipientsRefused as e:
                log.write(c, "rejected", str(e))
                failed += 1
                print(f"[{i}/{len(pending)}] REJECTED (bad address, won't retry) → {c['email']}")
            except smtplib.SMTPAuthenticationError as e:
                log.write(c, "error", str(e))
                print("Gmail rejected the login — stopping.", e)
                break
            except Exception as e:
                log.write(c, "error", str(e))
                failed += 1
                print(f"[{i}/{len(pending)}] ERROR → {c['email']}: {e}")
                if "daily" in str(e).lower() or "5.4.5" in str(e):
                    print("Gmail daily sending limit reached — stopping. Re-run tomorrow to continue.")
                    break
            if delay and i < len(pending):
                time.sleep(delay)
    except KeyboardInterrupt:
        print("\nStopped by you. Re-run the same command to continue.")
    finally:
        log.close()
        try:
            smtp.quit()
        except Exception:
            pass

    print(f"\nDone. sent={sent} failed={failed}. Log: {LOG_FILE}")


if __name__ == "__main__":
    main()
