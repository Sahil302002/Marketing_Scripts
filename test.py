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

# --------------------------------------------------------------
# Read the first MAX_EMAILS rows from the CSV (column A = manufacturer, column G = e‑mail)
# --------------------------------------------------------------
clients = []
with open(CSV_PATH, newline="", encoding="utf-8") as csvfile:
    reader = csv.DictReader(csvfile)
    for i, row in enumerate(reader):
        if i >= MAX_EMAILS:
            break
        clients.append(row)

def parse_emails(raw: str):
    """Split a raw e‑mail string into a list, handling commas and semicolons."""
    if not raw:
        return []
    return [e.strip() for e in raw.replace(";", ",").split(",") if e.strip()]

# --------------------------------------------------------------
# Send personalized e‑mail to each client
# --------------------------------------------------------------
for client in clients:
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
    except Exception as exc:  # pragma: no cover
        print(
            f"❌ Failed for {manufacturer} (to={to_addr}): {type(exc).__name__}: {exc}"
        )
    finally:
        # Respect the configured pause before the next e‑mail
        time.sleep(DELAY_SECONDS)
