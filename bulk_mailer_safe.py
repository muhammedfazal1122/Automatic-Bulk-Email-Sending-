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
CSV_FILE = "Job Hunt Tracker - Jobs new.csv"   # downloaded from the Google Sheet (Jobs tab)
TEMPLATE_FILE = "email_template.txt"
TEMPLATE = ""   # loaded in main()
ATTACH_PATH = "resume.pdf"
LOG_FILE = "send_log.csv"
SENDER_NAME = "Muhammed Fasal C"
SUBJECT = "Senior Python / Django Developer – Muhammed Fasal C – Resume Attached"
SEND_RATE_PER_MIN = 10        # 6 s between emails; keeps Gmail happy
MAX_RETRIES = 2               # per email, after a dropped connection

# Column names in the sheet (matched case-insensitively; only the email column is required).
# Each name may be a list: the first one found in the CSV wins, so the older HR-contact
# sheets (Email / Company / hiring firm / HR / recruiter name) still work.
EMAIL_COLS = ["HR Email", "Email"]
COMPANY_COLS = ["Company", "Company / hiring firm"]
NAME_COLS = ["HR / recruiter name"]
ROLE_COLS = ["Title"]

LOG_FIELDS = ["email", "company", "hr_name", "status", "error", "timestamp"]
EMAIL_RE = re.compile(r"^[\w.+-]+@[\w-]+(\.[\w-]+)+$")
FIND_EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")


# ---------------- helpers ----------------
def now():
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


PLACEHOLDER_RE = re.compile(r"^(not named|not stated|not found|n/?a|none|unknown|-+)\b", re.I)


def clean_company(raw):
    """'Touchworld Technology (via Netlyft)' -> 'Touchworld Technology';
    '(not named – personal email)' -> ''  (empty = no company known)."""
    s = re.sub(r"\s*[\(\[].*?[\)\]]", "", raw or "")   # drop (…) and […] notes
    s = " ".join(s.split()).strip(" ,;-–|")
    return "" if PLACEHOLDER_RE.match(s) else s


def find_col(df, wanted):
    """First column whose name matches `wanted` (a name or a list of names), else None."""
    for w in ([wanted] if isinstance(wanted, str) else wanted):
        for c in df.columns:
            if c.strip().lower() == w.lower():
                return c
    return None


def load_contacts(path):
    """Return a list of dicts {email, company, hr_name, role}, cleaned and de-duplicated.
    A cell may hold several addresses; cells like 'Not found' or blank are skipped.
    If one address appears on several rows (different companies/roles) the contact is
    kept once with company and role blanked, so the letter stays generic and correct."""
    df = pd.read_csv(path, encoding="utf-8-sig", dtype=str).fillna("")
    email_col = find_col(df, EMAIL_COLS)
    if not email_col:
        sys.exit(f"ERROR: no {EMAIL_COLS} column in {path}. Columns are: {list(df.columns)}")
    company_col = find_col(df, COMPANY_COLS)
    name_col = find_col(df, NAME_COLS)
    role_col = find_col(df, ROLE_COLS)

    by_email, invalid = {}, []
    for _, row in df.iterrows():
        cell = row[email_col].strip()
        found = [e.lower() for e in FIND_EMAIL_RE.findall(cell)]
        if not found:
            if cell and not PLACEHOLDER_RE.match(cell):
                invalid.append(cell)
            continue
        for email in found:
            if not EMAIL_RE.match(email):
                invalid.append(email)
                continue
            c = {
                "email": email,
                "company": clean_company(row[company_col]) if company_col else "",
                "hr_name": row[name_col].strip() if name_col else "",
                "role": row[role_col].strip() if role_col else "",
            }
            if email not in by_email:
                by_email[email] = c
            elif (by_email[email]["company"], by_email[email]["role"]) != (c["company"], c["role"]):
                by_email[email].update(company="", role="")   # ambiguous -> generic letter
    for e in invalid:
        print(f"  ! skipping invalid address: {e}")
    return list(by_email.values())


def render(contact):
    """Return (subject, body) for one contact, filling the {greeting} and {opening} slots."""
    company, role = contact.get("company", ""), contact.get("role", "")
    if company:
        greeting = f"Dear {company} Hiring Team,"
    else:
        greeting = "Dear Hiring Team,"
    if role and company:
        opening = f"I'm writing to apply for the {role} position at {company}."
    elif role:
        opening = f"I'm writing to apply for the {role} position."
    else:
        opening = "I'm writing to ask whether you have an opening for a Python / Django or full-stack role."
    subject = f"Application: {role} – {SENDER_NAME}" if role else SUBJECT
    body = TEMPLATE.replace("{greeting}", greeting).replace("{opening}", opening)
    return subject, body


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


def build_message(to_email, subject, body, attachment):
    msg = MIMEMultipart()
    msg["From"] = formataddr((SENDER_NAME, SENDER_EMAIL))
    msg["To"] = to_email
    msg["Subject"] = subject
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
    global TEMPLATE
    with open(TEMPLATE_FILE, encoding="utf-8") as f:
        TEMPLATE = f.read()

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
        subject, body = render({})
        send_with_retry(smtp, build_message(to, subject, body, attachment), to)
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
        print(f"\nAttachment: {ATTACH_PATH if attachment else 'MISSING!'}\n")
        if pending:
            subj, body = render(pending[0])
            print(f"--- sample letter (first recipient: {pending[0]['email']}) ---\nSubject: {subj}\n\n{body}")
        print("------ recipients ------")
        for i, c in enumerate(pending, 1):
            print(f"{i:4}. {c['email']:42} {c['company']:28} {c['role']}")
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
                subject, body = render(c)
                smtp = send_with_retry(smtp, build_message(c["email"], subject, body, attachment), c["email"])
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
