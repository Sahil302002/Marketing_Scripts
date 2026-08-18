import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
from urllib.parse import urljoin

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# Email regex
EMAIL_REGEX = r'[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}'

# Common pages where emails are often found

CONTACT_PAGES = [
    "contact",
    "contact-us",
    "contactus",
    "about",
    "about-us",
    "support",
    "team",
    "our-team",
    "staff",
    "management",
    "career",
    "careers",
    "jobs",
    "dealer",
    "dealers",
    "distributor",
    "branch",
    "branches",
    "office",
    "locations",
    "privacy-policy",
    "terms"
]

# Create a requests session with retries
session = requests.Session()

retry = Retry(
    total=3,
    backoff_factor=1,
    status_forcelist=[429, 500, 502, 503, 504]
)

adapter = HTTPAdapter(max_retries=retry)

session.mount("http://", adapter)
session.mount("https://", adapter)

headers = {
    "User-Agent": "Mozilla/5.0"
}


def find_emails(text):
    emails = set(re.findall(EMAIL_REGEX, text))
    return emails


def scrape_url(url):

    try:
        response = session.get(url, headers=headers, timeout=15)
        html = response.text

        emails = find_emails(html)

        if emails:
            return emails

        soup = BeautifulSoup(html, "html.parser")

        # Check links on homepage
        for link in soup.find_all("a", href=True):

            href = link["href"].lower()

            if any(page in href for page in CONTACT_PAGES):

                contact_url = urljoin(url, href)

                try:
                    contact_response = session.get(
                        contact_url,
                        headers=headers,
                        timeout=15
                    )

                    contact_emails = find_emails(contact_response.text)

                    if contact_emails:
                        emails.update(contact_emails)

                except:
                    pass

        # Try common URLs directly
        if not emails:

            for page in CONTACT_PAGES:

                page_url = urljoin(url + "/", page)

                try:
                    page_response = session.get(
                        page_url,
                        headers=headers,
                        timeout=15
                    )

                    page_emails = find_emails(page_response.text)

                    if page_emails:
                        emails.update(page_emails)

                except:
                    pass

        return emails

    except Exception:
        return set()


# ------------------------------

input_file = r"C:\Users\sahil\Documents\Neha General marketting\data\Cleaner.xlsx"
df = pd.read_excel(input_file)

# Create Email column if it doesn't exist
if "Email" not in df.columns:
    df["Email"] = ""

for index, row in df.iterrows():

    website = row["Website URL"]

    # Skip blank/NaN values
    if pd.isna(website):
        print(f"Row {index + 2}: Website is blank. Skipping...")
        continue

    website = str(website).strip()

    # Skip empty string or dash
    if website == "" or website == "-":
        print(f"Row {index + 2}: Website is empty or '-'. Skipping...")
        continue

    # Add https:// if missing
    if not website.startswith(("http://", "https://")):
        website = "https://" + website

    print(f"Checking: {website}")

    emails = scrape_url(website)

    df.at[index, "Email"] = ", ".join(sorted(emails)) if emails else "Not Found"


df.to_excel(input_file, index=False)

print("\nDone!")
print("Email column updated successfully.")