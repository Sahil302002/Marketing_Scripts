import os
import csv
import smtplib
import time
from email.message import EmailMessage
from dotenv import load_dotenv
import textwrap

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
CSV_PATH = os.getenv("CSV_PATH")

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
}
missing = [k for k, v in required.items() if not v]
if missing:
    raise RuntimeError(f"Missing required env variables: {', '.join(missing)}")

# --------------------------------------------------------------
# Load the PDF once – the same file will be attached to every mail
# --------------------------------------------------------------
with open(PDF_PATH, "rb") as f:
    pdf_bytes = f.read()

def parse_emails(raw: str):
    """Split a raw e‑mail string into a list, handling commas and semicolons."""
    if not raw:
        return []
    return [e.strip() for e in raw.replace(";", ",").split(",") if e.strip()]

def load_clients(csv_path):
    """Read the CSV into a list of row dicts and return the fieldnames too."""
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader), reader.fieldnames

def write_clients(csv_path, rows, fieldnames):
    """Atomically rewrite the CSV. Uses a .tmp file then os.replace so a
    crash mid‑write cannot leave a half‑written file on disk."""
    tmp = csv_path + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    os.replace(tmp, csv_path)

# --------------------------------------------------------------
# Load the CSV (column A = manufacturer, column G = e‑mail).
# We read the whole file so we can rewrite it after each successful send,
# dropping sent rows to prevent duplicate e‑mails on re‑runs.
# --------------------------------------------------------------
clients, fieldnames = load_clients(CSV_PATH)
pending_rows = list(clients)

# --------------------------------------------------------------
# Send personalized e‑mail to each client. After a successful send,
# the row is dropped from pending_rows and the CSV is rewritten so a
# re‑run will not contact the same person again.
# --------------------------------------------------------------
sent_in_run = 0
while pending_rows and sent_in_run < MAX_EMAILS:
    client = pending_rows[0]
    manufacturer = client.get("Manufacturer/ Shop Name", "").strip()
    raw_emails = client.get("Email", "")
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

        # Drop the sent row from the CSV so this client is not contacted again
        pending_rows.pop(0)
        try:
            write_clients(CSV_PATH, pending_rows, fieldnames)
        except OSError as write_exc:
            # Don't lose the in‑memory progress: keep the row removed locally,
            # but warn the operator so they can inspect the file.
            print(
                f"⚠️  Could not update CSV after sending to {manufacturer}: "
                f"{type(write_exc).__name__}: {write_exc}"
            )
        sent_in_run += 1
    except Exception as exc:  # pragma: no cover
        print(
            f"❌ Failed for {manufacturer} (to={to_addr}): {type(exc).__name__}: {exc}"
        )
    finally:
        # Respect the configured pause before the next e‑mail
        time.sleep(DELAY_SECONDS)
