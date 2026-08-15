# Neha General Marketing - Email Automation Tools

This project contains Python scripts designed to automate the process of finding prospective client emails from company websites and sending them personalized outreach emails with attachments.

## Project Structure

```text
├── README.md
├── data/
└── scripts/
    ├── Python_Email_Fetcher.py  # Scrapes email addresses from website URLs
    └── email_sender.py          # Automates personalized email outreach
```

---

## Scripts Overview

### 1. Python Email Fetcher (`scripts/Python_Email_Fetcher.py`)

This script automatically crawls and extracts email addresses from a list of company websites stored in an Excel spreadsheet.

#### Key Features
* **Excel Integration:** Reads website URLs from `Cleaner.xlsx`, processes them, and writes the results back to an `Email` column.
* **Targeted Page Crawling:** If emails are not found on the home page, it automatically scans common pages (like Contact Us, About, Support, Careers, etc.) or tries direct common URLs.
* **Robust Network Handling:** Implements an exponential backoff retry strategy for handling temporary server errors or rate limits (HTTP 429, 500, 502, 503, 504) and enforces a 15-second request timeout.
* **Email Filtering:** Uses regular expressions to match and isolate standard email structures.

#### Setup & Requirements
* Requires an input file at: `C:\Users\sahil\Downloads\Neha General marketting\Cleaner.xlsx`
* Dependencies: `pandas`, `requests`, `beautifulsoup4`, `urllib3`, `openpyxl` (for reading Excel files)

---

### 2. Email Sender (`scripts/email_sender.py`)

This script automates sending personalized cold outreach emails with PDF attachments to prospective clients using SMTP and credentials loaded securely from local environment variables.

#### Key Features
* **Secure Configurations:** Keeps sensitive details (SMTP server, passwords, ports) safe by loading them from a local, uncommitted `.env` file.
* **Personalized Outreach:** Dynamically addresses each email to the client's shop/manufacturer name.
* **Anti-Spam & Delivery Protection:**
  * **Throttling:** Implements a configurable delay (default: 5 seconds) between sends to protect domain reputation.
  * **Batch Limits:** Restricts the maximum number of emails sent per execution (default: 10 emails) to stay within safe sending limits.
* **Attachments:** Automatically loads and attaches a specified PDF brochure to every message.

#### Setup & Configuration
To run this script, create a `.env` file in the same directory as the script (`scripts/`) with the following variables:

```env
SMTP_SERVER=your.smtp.server.com
SMTP_PORT=465
EMAIL=your-email@domain.com
APP_PASSWORD=your_secure_app_password
PDF_PATH=path/to/your/attachment.pdf
CSV_PATH=path/to/your/client_leads.csv
MAX_EMAILS=10
DELAY_SECONDS=5.0
```

* Dependencies: `python-dotenv`
