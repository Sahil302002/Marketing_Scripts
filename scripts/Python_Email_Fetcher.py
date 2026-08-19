import os
import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
from urllib.parse import urljoin

from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from dotenv import load_dotenv

load_dotenv()

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


def dedupe_by_manufacturer(df):
    """Drop rows that repeat a Manufacturer/Shop Name already seen earlier in
    the sheet, so no company appears more than once in Cleaner.xlsx. Rows
    with a blank name are left alone since they can't be confidently matched
    to one another. Among duplicates, preference goes to (in order): a row
    with a real Email result (not blank or 'Not Found'), then a row that at
    least has a Website URL to scrape, otherwise the first occurrence.
    """
    # .fillna("") before stringifying so a real NaN can't slip through as the
    # literal text "nan" and be mistaken for a filled-in value.
    name_col = df.columns[0]
    normalized = df[name_col].fillna("").astype(str).str.strip().str.lower()
    is_blank_name = normalized == ""

    email = df["Email"] if "Email" in df.columns else pd.Series("", index=df.index)
    email_text = email.fillna("").astype(str).str.strip()
    no_result = email_text.isin(["", "Not Found"])

    website = df["Website URL"] if "Website URL" in df.columns else pd.Series("", index=df.index)
    website_text = website.fillna("").astype(str).str.strip()
    no_website = website_text.isin(["", "-"])

    sort_key = pd.DataFrame({
        "normalized": normalized,
        "no_result": no_result,
        "no_website": no_website,
        "order": range(len(df)),
    }, index=df.index)

    keep = is_blank_name.copy()
    ranked = sort_key[~is_blank_name].sort_values(["normalized", "no_result", "no_website", "order"])
    keep.loc[ranked.drop_duplicates(subset="normalized", keep="first").index] = True

    removed = int((~keep).sum())
    if removed:
        print(f"🧹 Removed {removed} duplicate manufacturer/shop-name row(s) before scraping.")

    return df[keep].reset_index(drop=True)


# ------------------------------

input_file = os.getenv("CLEANER_XLSX_PATH")
if not input_file:
    raise RuntimeError("Missing required env variable: CLEANER_XLSX_PATH")

df = pd.read_excel(input_file)

# Create Email column if it doesn't exist
if "Email" not in df.columns:
    df["Email"] = ""

# Remove duplicate manufacturers/shops up front and persist the cleaned
# sheet before scraping starts, so no company appears more than once.
df = dedupe_by_manufacturer(df)
df.to_excel(input_file, index=False)

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