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
* **Duplicate Removal:** Before scraping starts, drops rows that repeat a `Manufacturer/Shop Name` already seen earlier in `Cleaner.xlsx` and saves the cleaned sheet immediately, so no company is scraped or listed more than once. Rows with a blank name are left alone (can't be confidently matched). Among duplicates, the row kept is (in order of preference): one with a real `Email` result (not blank or `Not Found`), otherwise one that at least has a `Website URL` to scrape, otherwise the first occurrence.
* **Excel Integration:** Reads website URLs from `Cleaner.xlsx`, processes them, and writes the results back to an `Email` column.
* **Targeted Page Crawling:** If emails are not found on the home page, it automatically scans common pages (like Contact Us, About, Support, Careers, etc.) or tries direct common URLs.
* **Robust Network Handling:** Implements an exponential backoff retry strategy for handling temporary server errors or rate limits (HTTP 429, 500, 502, 503, 504) and enforces a 15-second request timeout.
* **Email Filtering:** Uses regular expressions to match and isolate standard email structures.

#### Setup & Requirements
* Reads its input file from the `CLEANER_XLSX_PATH` variable in `.env` (the same one used by `email_sender.py` — see below).
* Dependencies: `pandas`, `requests`, `beautifulsoup4`, `urllib3`, `openpyxl`, `python-dotenv`

---

### 2. Email Sender (`scripts/email_sender.py`)

This script automates sending personalized cold outreach emails with PDF attachments to prospective clients using SMTP and credentials loaded securely from local environment variables. It treats `data/Cleaner.xlsx` as a **live queue** — rows leave it as soon as they're successfully emailed or found to be permanently undeliverable — and keeps a running history of every outcome in `data/client_data.csv`.

#### The Three Data Files

| File | Role |
|---|---|
| `data/Cleaner.xlsx` | **The queue.** Every prospective client still waiting to be emailed. Rows are removed from here once they've been sent or permanently failed. |
| `data/client_data.csv` | **The history.** Append-only log of every email that was successfully `Sent` or that `Failed` permanently, each with a timestamp (and, for failures, the reason). Existing rows are never touched — this is how you track total emails sent over time. |
| `data/No_Mail_Contacts.csv` | **The blockers.** Contacts from `Cleaner.xlsx` whose `Email` column is blank or `Not Found` — flagged here for manual follow-up. Also append-only. |

#### What Happens on Each Run

1. **No-Mail Contact Sync (Step 0):** Scans `Cleaner.xlsx` for rows with a blank or `Not Found` email and appends them to `No_Mail_Contacts.csv`. These rows are *not* removed from `Cleaner.xlsx`, so a contact still missing an email gets re-flagged (and re-appended) on every run until it's fixed or deleted from the spreadsheet.
2. **Candidate selection:** Takes up to `MAX_EMAILS` rows from `Cleaner.xlsx` that *do* have a usable email address.
3. **Send loop:** For each candidate, sends a personalized email (with the PDF attached), then waits `DELAY_SECONDS` before moving on. What happens next depends on the outcome:
   * ✅ **Sent successfully** → the row is appended to `client_data.csv` with `Status = Sent <timestamp>` and removed from `Cleaner.xlsx`. It will never be emailed again.
   * 🗂️ **Failed permanently** (e.g. the address doesn't exist, or the server issued a 5xx rejection) → the row is appended to `client_data.csv` with `Status = Failed <timestamp> – <reason>` and removed from `Cleaner.xlsx`, since re-running the script would just fail the same way again.
   * ❌ **Failed transiently** (e.g. a dropped connection, timeout, or a 4xx "try again later" response) → the row is left untouched in `Cleaner.xlsx` and not logged anywhere, so it's automatically retried the next time the script runs.

#### Key Features
* **Secure Configurations:** Keeps sensitive details (SMTP server, passwords, ports) safe by loading them from a local, uncommitted `.env` file.
* **Personalized Outreach:** Dynamically addresses each email to the client's shop/manufacturer name.
* **Anti-Spam & Delivery Protection:**
  * **Throttling:** Implements a configurable delay (default: 5 seconds) between sends to protect domain reputation.
  * **Batch Limits:** Restricts the maximum number of emails sent per execution (default: 10 emails) to stay within safe sending limits.
* **Attachments:** Automatically loads and attaches a specified PDF brochure to every message.
* **No Repeat Sends:** a client is physically removed from `Cleaner.xlsx` the moment it's sent or permanently failed, so re-running the script never re-sends to the same contact.
* **Smart Failure Handling:** distinguishes failures worth retrying (network hiccups, server timeouts) from ones that never will succeed (invalid/refused addresses), so the queue doesn't get stuck retrying a bad address forever, and doesn't lose track of a good one on a temporary glitch.

#### Setup & Configuration
To run this script, create a `.env` file in the same directory as the script (`scripts/`) with the following variables:

```env
SMTP_SERVER=your.smtp.server.com
SMTP_PORT=465
EMAIL=your-email@domain.com
APP_PASSWORD=your_secure_app_password
PDF_PATH=path/to/your/attachment.pdf
CSV_PATH=data\client_data.csv
MAX_EMAILS=10
DELAY_SECONDS=5.0
CLEANER_XLSX_PATH=data\Cleaner.xlsx
NO_MAIL_CONTACTS_CSV_PATH=data\No_Mail_Contacts.csv
```

* Dependencies: `python-dotenv`, `pandas`, `openpyxl`
