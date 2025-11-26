# clean_emails.py
# Run: python clean_emails.py
#
# Reads emails.csv and writes emails_clean.csv with columns:
# Email,CompanyName

import csv
import re

INPUT = "emails.csv"
OUTPUT = "emails_clean.csv"

# simple email pattern (will catch most normal emails)
EMAIL_RE = re.compile(r'[\w\.-]+@[\w\.-]+\.\w+')

def normalize_company(text):
    if text is None:
        return ""
    s = text.strip()
    # remove repeated spaces, newlines and trailing commas
    s = " ".join(s.split())
    s = s.rstrip(",")
    return s

def main():
    seen = set()
    rows_out = []

    with open(INPUT, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        # skip header if present
        first = next(reader, None)
        # detect if header looks like header row
        if first:
            joined = ",".join(first).lower()
            if "email" in joined and ("first" in joined or "company" in joined):
                # header present; continue with next rows
                pass
            else:
                # not header — process the first row as data
                # put it back by processing it manually
                reader = [first] + list(reader)

        for row in reader:
            if not row:
                continue

            # Join row into a single string to find emails if they are in quoted/multiline cells
            row_join = ",".join(row)

            # find ALL emails in the row string
            emails = EMAIL_RE.findall(row_join)

            # company candidate: prefer column index 1 if exists
            company = ""
            if len(row) >= 2:
                company = normalize_company(row[1])

            # fallback: look for any non-empty field (besides email) that looks like a company
            if (not company) or company.upper() in ("ERROR", "NO_RECIPIENT"):
                for cell in row[1:]:
                    if cell and cell.strip() and cell.strip().upper() not in ("ERROR", "NO_RECIPIENT"):
                        candidate = normalize_company(cell)
                        if candidate:
                            company = candidate
                            break

            # If still empty, try to infer company from later columns
            if (not company) or company.upper() in ("ERROR", "NO_RECIPIENT"):
                # try text after first email in joined string (very last resort)
                if emails:
                    after = row_join.split(emails[0], 1)[-1]
                    parts = after.split(",")
                    if len(parts) > 1:
                        company = normalize_company(parts[1])

            # Add each found email as its own output row
            for e in emails:
                e_norm = e.strip().lower()
                if not e_norm:
                    continue
                key = (e_norm, company)
                if key in seen:
                    continue
                seen.add(key)
                rows_out.append((e_norm, company))

    # write cleaned CSV
    with open(OUTPUT, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Email", "CompanyName"])
        for e, c in rows_out:
            writer.writerow([e, c])

    print(f"Done. {len(rows_out)} records written to {OUTPUT}")

if __name__ == "__main__":
    main()
