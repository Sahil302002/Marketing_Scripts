
import smtplib
from email.message import EmailMessage

# ==========================================
# CONFIGURATION
# ==========================================

SMTP_SERVER = "smtp.zoho.in"
SMTP_PORT = 465

EMAIL = "sales@monolithicrefractory.net"
APP_PASSWORD = "nymFKnAniyGp"

TO_EMAIL = "sahilgupta302002@gmail.com"

# ==========================================

msg = EmailMessage()

msg["Subject"] = "SMTP Test from Python"

msg["From"] = EMAIL

msg["To"] = TO_EMAIL

msg.set_content(
"""Hi,

This is a test email sent successfully using Python and Zoho SMTP.

Everything is working correctly.

Regards,
Sahil
"""
)

try:

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as smtp:

        smtp.login(EMAIL, APP_PASSWORD)

        smtp.send_message(msg)

    print("=" * 50)
    print("SUCCESS")
    print("=" * 50)
    print("Email sent successfully.")

except Exception as e:

    print("=" * 50)
    print("FAILED")
    print("=" * 50)

    print(type(e).__name__)
    print(e)