import os
import csv
import smtplib
import time
from datetime import datetime
from email.message import EmailMessage
from dotenv import load_dotenv
import textwrap
import pandas as pd

# --------------------------------------------------------------
# Load configuration from .env (kept out of version control)
# --------------------------------------------------------------
load_dotenv()  # .env should be in the same directory as this script

# ----- Core configuration (required) -----
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))
EMAIL = os.getenv("EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
PDF_PATH = os.getenv("PDF_PATH")
CSV_PATH = os.getenv("CSV_PATH")  # historical log of successfully sent emails
CLEANER_XLSX_PATH = os.getenv("CLEANER_XLSX_PATH")  # source of clients to email
NO_MAIL_CSV_PATH = os.getenv("NO_MAIL_CONTACTS_CSV_PATH")

# ----- Optional runtime parameters -----
# Number of emails to send in this run (default 10)
MAX_EMAILS = int(os.getenv("MAX_EMAILS", "10"))
# Seconds to wait between each send (default 5 seconds)
DELAY_SECONDS = float(os.getenv("DELAY_SECONDS", "5"))

# Basic validation – abort early if anything essential is missing
required = {
    "SMTP_SERVER": SMTP_SERVER,
    "SMTP_PORT": SMTP_PORT,
    "EMAIL": EMAIL,
    "APP_PASSWORD": APP_PASSWORD,
    "PDF_PATH": PDF_PATH,
    "CSV_PATH": CSV_PATH,
    "CLEANER_XLSX_PATH": CLEANER_XLSX_PATH,
    "NO_MAIL_CONTACTS_CSV_PATH": NO_MAIL_CSV_PATH,
}
missing = [k for k, v in required.items() if not v]
if missing:
    raise RuntimeError(f"Missing required env variables: {', '.join(missing)}")

# --------------------------------------------------------------
# Shared helpers for reading Cleaner.xlsx and writing the contact CSVs
# (No_Mail_Contacts.csv and client_data.csv both use this schema).
# --------------------------------------------------------------

CONTACT_CSV_HEADER = ["Sr.No", "Manufacturer/ Shop Name", "Type", "Contact No.", "Website URL", "Email", "Status"]


def _is_missing_email(value):
    """True when a cell is blank/NaN or the literal text 'Not Found'."""
    text = "" if pd.isna(value) else str(value).strip()
    return text == "" or text.lower() == "not found"


def _cell(value):
    """Render a cell for CSV output, turning NaN into an empty string."""
    return "" if pd.isna(value) else str(value).strip()


def _detect_source_columns(df):
    """Locate the Email column and positionally map the rest onto
    (name, type, phone, website), matching Cleaner.xlsx's layout."""
    email_col = next((c for c in df.columns if str(c).strip().lower() == "email"), None)
    if email_col is None:
        raise RuntimeError("No 'Email' column found in the source spreadsheet")

    other_cols = [c for c in df.columns if c != email_col]
    return {
        "email": email_col,
        "name": other_cols[0] if len(other_cols) > 0 else None,
        "type": other_cols[1] if len(other_cols) > 1 else None,
        "phone": other_cols[2] if len(other_cols) > 2 else None,
        "website": other_cols[3] if len(other_cols) > 3 else None,
    }


def _next_sr_no(csv_path):
    """Next Sr.No to use when appending to a contact CSV, continuing the existing count."""
    if not (os.path.isfile(csv_path) and os.path.getsize(csv_path) > 0):
        return 1
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        return max(sum(1 for _ in csv.reader(f)) - 1, 0) + 1


def _append_contact_rows(csv_path, rows):
    """Append rows to a contact CSV, creating it with CONTACT_CSV_HEADER if needed."""
    file_exists = os.path.isfile(csv_path) and os.path.getsize(csv_path) > 0
    os.makedirs(os.path.dirname(csv_path) or ".", exist_ok=True)
    with open(csv_path, "a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if not file_exists:
            writer.writerow(CONTACT_CSV_HEADER)
        writer.writerows(rows)


def sync_no_mail_contacts(xlsx_path, csv_path):
    """Append Cleaner.xlsx rows with a blank/'Not Found' e‑mail to the no‑mail CSV.

    Existing rows in the CSV are preserved; new rows are appended below them.
    Creates the CSV with headers if it doesn't exist yet.
    """
    df = pd.read_excel(xlsx_path)
    cols = _detect_source_columns(df)

    flagged = df[df[cols["email"]].apply(_is_missing_email)]
    if flagged.empty:
        print("📋 Cleaner sync: no blank/'Not Found' emails found – nothing to append.")
        return

    next_sr_no = _next_sr_no(csv_path)
    rows = []
    for offset, (_, row) in enumerate(flagged.iterrows()):
        email_text = _cell(row[cols["email"]])
        status = "Blank Email" if email_text == "" else "Not Found"
        rows.append([
            next_sr_no + offset,
            _cell(row[cols["name"]]) if cols["name"] else "",
            _cell(row[cols["type"]]) if cols["type"] else "",
            _cell(row[cols["phone"]]) if cols["phone"] else "",
            _cell(row[cols["website"]]) if cols["website"] else "",
            email_text,
            status,
        ])
    _append_contact_rows(csv_path, rows)

    print(f"📋 Cleaner sync: appended {len(flagged)} row(s) with missing/blank email to {csv_path}")


if os.path.isfile(CLEANER_XLSX_PATH):
    try:
        sync_no_mail_contacts(CLEANER_XLSX_PATH, NO_MAIL_CSV_PATH)
    except Exception as exc:  # pragma: no cover
        print(f"⚠️  Cleaner sync skipped due to error: {type(exc).__name__}: {exc}")
else:
    print(f"⚠️  Cleaner sync skipped: {CLEANER_XLSX_PATH} not found.")

# --------------------------------------------------------------
# Load the PDF once – the same file will be attached to every mail
# --------------------------------------------------------------
with open(PDF_PATH, "rb") as f:
    pdf_bytes = f.read()

# --------------------------------------------------------------
# Load the first MAX_EMAILS clients with a usable e‑mail from Cleaner.xlsx.
# Rows without a valid e‑mail were already flagged in the sync step above.
# --------------------------------------------------------------
source_df = pd.read_excel(CLEANER_XLSX_PATH)
source_cols = _detect_source_columns(source_df)
candidates = source_df[~source_df[source_cols["email"]].apply(_is_missing_email)].head(MAX_EMAILS)


def parse_emails(raw: str):
    """Split a raw e‑mail string into a list, handling commas and semicolons."""
    if not raw:
        return []
    return [e.strip() for e in raw.replace(";", ",").split(",") if e.strip()]


def _decode(value):
    return value.decode(errors="replace") if isinstance(value, bytes) else str(value)


def _classify_failure(exc):
    """Decide whether a send failure is permanent (re-running won't fix it) or transient.

    A refused/invalid recipient address or a 5xx SMTP response is permanent —
    the same email to the same address will fail again. Connection issues,
    timeouts, and 4xx responses are treated as transient and left for a later
    run to retry.
    """
    if isinstance(exc, smtplib.SMTPRecipientsRefused):
        details = "; ".join(
            f"{addr}: {code} {_decode(msg)}" for addr, (code, msg) in exc.recipients.items()
        )
        return True, f"Recipient refused – {details}"
    if isinstance(exc, smtplib.SMTPResponseException):
        return 500 <= exc.smtp_code < 600, f"{exc.smtp_code} {_decode(exc.smtp_error)}"
    return False, f"{type(exc).__name__}: {exc}"


def _finalize_client(row_index, client, status_text):
    """Append the client's row to the historical CSV log and drop it from the
    source spreadsheet so it is neither re-sent nor re-attempted."""
    global source_df
    row = [
        _next_sr_no(CSV_PATH),
        _cell(client[source_cols["name"]]) if source_cols["name"] else "",
        _cell(client[source_cols["type"]]) if source_cols["type"] else "",
        _cell(client[source_cols["phone"]]) if source_cols["phone"] else "",
        _cell(client[source_cols["website"]]) if source_cols["website"] else "",
        _cell(client[source_cols["email"]]) if source_cols["email"] else "",
        status_text,
    ]
    _append_contact_rows(CSV_PATH, [row])
    source_df = source_df.drop(index=row_index)
    source_df.to_excel(CLEANER_XLSX_PATH, index=False)

# --------------------------------------------------------------
# Send personalized e‑mail to each client
# --------------------------------------------------------------
for row_index, client in candidates.iterrows():
    manufacturer = _cell(client[source_cols["name"]]) if source_cols["name"] else ""
    raw_emails = _cell(client[source_cols["email"]])
    email_list = parse_emails(raw_emails)

    if not email_list:
        print(f"⚠️  Skipping '{manufacturer}': no e‑mail address found.")
        # Still respect the delay before moving to the next row
        time.sleep(DELAY_SECONDS)
        continue

    to_addr = email_list[0]
    cc_addrs = email_list[1:]

    body = f"""Dear {manufacturer},

We are writing to formally introduce Monolithic Refractory Corporation as a reliable manufacturer and global exporter. We are interested in becoming an approved supplier for your high‑temperature lining requirements.

Since 1993, we have partnered with heavy industries worldwide to deliver high‑performance monolithic refractories designed to maximize equipment lifecycles and minimize operational downtime. Our product portfolio includes:

- Premium castables, ramming masses, plastic refractories, patching mixes, fire bricks, pre‑cast shapes, mortars, ceramic fiber blankets, ropes, and boards.
- Engineered excellence: high thermal‑shock resistance, mechanical strength, and abrasion resistance tailored for harsh environments.
- Proven reliability: over 30 years of manufacturing expertise ensuring strict consistency and seamless international logistics.

Could you please share your vendor registration procedure or direct us to the appropriate portal to join your approved‑supplier list?

Best regards,

Monolithic Refractory Corporation
Mob: +91‑8488824477
Email: sales@monolithicrefractory.com
Website: www.monolithicrefractory.com
"""

    body = textwrap.dedent(body).strip()
    msg = EmailMessage()
    msg["Subject"] = f"Monolithic Refractory – Introduction for {manufacturer}"
    msg["From"] = EMAIL
    msg["To"] = to_addr
    if cc_addrs:
        msg["Cc"] = ", ".join(cc_addrs)
    msg.set_content(body)
    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=os.path.basename(PDF_PATH),
    )

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:
            smtp.login(EMAIL, APP_PASSWORD)
            smtp.send_message(msg)
        cc_display = ", ".join(cc_addrs) if cc_addrs else "none"
        print(f"✅ Sent to {to_addr} (CC: {cc_display}) – {manufacturer}")
        _finalize_client(row_index, client, f"Sent {datetime.now().isoformat()}")
    except Exception as exc:  # pragma: no cover
        print(
            f"❌ Failed for {manufacturer} (to={to_addr}): {type(exc).__name__}: {exc}"
        )
        permanent, reason = _classify_failure(exc)
        if permanent:
            _finalize_client(row_index, client, f"Failed {datetime.now().isoformat()} – {reason}")
            print(f"🗂️  Not retryable – moved to {CSV_PATH} as Failed ({reason}).")
    finally:
        # Respect the configured pause before the next e‑mail
        time.sleep(DELAY_SECONDS)
