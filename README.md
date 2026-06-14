# Google Company Information Extractor

This project extracts company information from Google Places API and saves the results into Excel.

The project was designed for a large list of companies, for example more than 3,000 records, where the script must continue safely even if the internet connection is lost, Google API returns temporary errors, or the program stops unexpectedly.

The final output contains company information such as:

* Company name found on Google
* Address
* Phone number
* International phone number
* Opening hours
* Website
* Google Maps URL
* Business status
* Google Place ID

---

## Project Purpose

The goal of this project is to automate the process of collecting public company information from Google based on company names and countries.

In this project, we did the following:

1. Extracted unique company and country records from one of our SQL Server table.
2. Saved the unique company-country list into an Excel file.
3. Used that Excel file as input for the Google Places API scraper.
4. Sent company search requests to Google Places Text Search API.
5. Extracted address, phone number, opening hours, website, Google Maps URL, and business status.
6. Saved results into Excel.
7. Added checkpoint/resume functionality so that already processed companies are not requested again.
8. Added retry logic for connection errors, timeout errors, rate-limit errors, and temporary Google API errors.

---

## Important Note

This project does not scrape Google HTML pages directly.

Instead, it uses the official Google Places API.

This is more stable and safer than using dynamic XPath values from Google pages because Google-generated HTML IDs change frequently and can break the scraper.

---

## Project Structure

Recommended repository structure:

```text
google-company-info-extractor/
│
├── README.md
├── requirements.txt
├── .gitignore
│
├── export_unique_companies_from_sql.py
├── google_places_company_info.py
│
├── input/
│   └── companies.xlsx
│
├── output/
│   ├── unique_companies_by_country.xlsx
│   ├── google_company_info.xlsx
│   └── google_company_info_checkpoint.csv
```

---

## Requirements

This project uses Python 3.9 or newer.

Required Python packages:

```text
requests
pandas
openpyxl
tqdm
pyodbc
```




---

## Google Places API Setup

To use this project, you need a Google Maps Platform API key with Places API enabled.

### Step 1: Open Google Cloud Console

Go to Google Cloud Console and log in with your Google account.

### Step 2: Create a New Project

Create a project, for example:

```text
Company Info Extractor
```

### Step 3: Enable Billing

Google Places API requires billing to be enabled.

Even if your usage is small, billing must be connected to the Google Cloud project.

### Step 4: Enable Places API

In Google Cloud Console:

```text
APIs & Services
→ Library
→ Search for "Places(new) API"
→ Enable Places(new) API
```

Use Places API, especially Places API New, because this project uses this endpoint:

```text
https://places.googleapis.com/v1/places:searchText
```

### Step 5: Create an API Key

Go to:

```text
APIs & Services
→ Credentials
→ Create Credentials
→ API Key
```

Copy the generated API key.

### Step 6: Restrict the API Key

For security, do not leave the API key unrestricted.

Recommended restrictions:

```text
API restrictions:
Restrict key → Places API
```

For a local Python script, you can restrict by IP address if you are running the script from a fixed server IP.

Avoid putting the API key directly inside the Python file. instead, add them into your environment using : setx Scrapper "YOUR API KEY" in windows powershel or export GOOGLE_MAPS_API_KEY="your_google_api_key_here" in Linux or macOS

---

## Setting the API Key

This project reads the API key from an environment variable.

The script supports either of these environment variable names:

```text
Scrapper
GOOGLE_MAPS_API_KEY
```

Using `GOOGLE_MAPS_API_KEY` is recommended because the name is clearer.

---



## Input Excel Format

The Google scraper expects an Excel file with this structure:

```text
company_name | location
```

Example:

```text
company_name       | location
Siemens            | Germany
Raykasoft          | UK
Afra Trading       | Iran
```

Default input file:

```text
companies.xlsx
```

Recommended path:

```text
input/companies.xlsx
```

---

## Google Places API Script

The main script is:

```text
google_places_company_info.py
```

It reads companies from Excel and sends requests to Google Places Text Search API.

The API request uses this endpoint:

```text
https://places.googleapis.com/v1/places:searchText
```

The request searches for:

```text
company_name + location
```

For example:

```text
Siemens Germany
Raykasoft UK
Afra Trading Iran
```

---

## Extracted Fields

The script requests these fields from Google:

```text
places.id
places.displayName
places.formattedAddress
places.nationalPhoneNumber
places.internationalPhoneNumber
places.regularOpeningHours
places.currentOpeningHours
places.websiteUri
places.googleMapsUri
places.businessStatus
```

The final Excel contains:

```text
unique_key
input_company
input_location
search_query
status
google_name
address
phone_number
international_phone_number
hours
website
google_maps_url
business_status
place_id
```

---

## Checkpoint and Resume System

Because the project may send thousands of requests, the script saves every result immediately.

It creates this checkpoint file:

```text
google_company_info_checkpoint.csv
```

This file is updated after every company request.

If the script stops because of:

```text
Internet disconnection
Timeout
Google temporary error
Program crash
Computer restart
```

you can run the script again.

The script will read the checkpoint file and skip companies that were already completed.

---

## Completed Statuses

The script skips records with these statuses:

```text
Found
Not found
```

This means the company was already checked successfully.

---

## Retryable Statuses

The script retries temporary errors such as:

```text
429 Too Many Requests
500 Internal Server Error
502 Bad Gateway
503 Service Unavailable
504 Gateway Timeout
Timeout error
Connection error
```

If a company fails after all retries, the error is saved in the checkpoint file.

Failed records are not treated as completed, so they can be retried when you run the script again.

---

## Output Files

The project creates these files:

### 1. Checkpoint CSV

```text
google_company_info_checkpoint.csv
```

This is the most important file during execution.

It protects your progress.

### 2. Final Excel

```text
google_company_info.xlsx
```

This file contains the clean final results.

### 3. Summary Sheet

The final Excel contains a summary sheet showing how many records have each status:

```text
Found
Not found
HTTP error
Failed after retries
```

---

## How to Run the Full Workflow

### Step 1: Install dependencies

```bash
pip install -r requirements.txt
```

### Step 2: Set Google API key

Windows CMD:

```cmd
set GOOGLE_MAPS_API_KEY=your_google_api_key_here
```

PowerShell:

```powershell
$env:GOOGLE_MAPS_API_KEY="your_google_api_key_here"
```

Linux/macOS:

```bash
export GOOGLE_MAPS_API_KEY="your_google_api_key_here"
```

### Step 3: Export unique companies from SQL Server

```bash
python export_unique_companies_from_sql.py
```

This creates:

```text
unique_companies_by_country.xlsx
```

### Step 4: Rename or copy the SQL output file

The Google scraper expects:

```text
companies.xlsx
```

Make sure the file contains:

```text
company_name
location
```

### Step 5: Run Google Places scraper

```bash
python google_places_company_info.py
```

### Step 6: Check output

After running, check:

```text
google_company_info.xlsx
google_company_info_checkpoint.csv
```

---

## Recommended Git Ignore

Create a `.gitignore` file:

```text
__pycache__/
*.pyc

.env
*.log

input/*.xlsx
output/*.xlsx
*.xlsx
*.csv

google_company_info_checkpoint.csv
google_company_info.xlsx
companies.xlsx
unique_companies_by_country.xlsx
```

Important:

Do not upload your API key to GitHub.

Do not hard-code your API key inside the Python source code.

---

## Common Problems

### Problem: API key is not set

Error:

```text
API key is not set
```

Solution:

Set the environment variable:

```cmd
set GOOGLE_MAPS_API_KEY=your_google_api_key_here
```

or:

```powershell
$env:GOOGLE_MAPS_API_KEY="your_google_api_key_here"
```

---

### Problem: SQL Server SSL or certificate error

If you get an SSL or certificate error with ODBC Driver 18, use:

```text
TrustServerCertificate=yes
```

Example:

```python
connection_string = (
    "DRIVER={ODBC Driver 18 for SQL Server};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    "Trusted_Connection=yes;"
    "Encrypt=yes;"
    "TrustServerCertificate=yes;"
)
```

---

### Problem: ODBC Driver not found

Error example:

```text
Data source name not found
ODBC Driver 18 for SQL Server not found
```

Solution:

Install Microsoft ODBC Driver 18 for SQL Server.

Then check your installed drivers with this Python code:

```python
import pyodbc
print(pyodbc.drivers())
```

---

### Problem: Script is too slow

The script sends one request per company.

For 3,000+ companies, this can take time.

You can change this value:

```python
REQUEST_DELAY_SECONDS = 0.2
```

Be careful when reducing the delay because too many requests too quickly may cause rate-limit errors.

---

### Problem: Duplicate company names

The script uses this key:

```text
company_name + location
```

For example:

```text
siemens|||germany
```

This prevents repeated requests for the same company-country pair.

---

## Safety and Cost Notes

Google Places API is a paid API.

Before running thousands of requests:

1. Check your Google Cloud billing settings.
2. Set a budget alert in Google Cloud.
3. Restrict your API key.
4. Test the script with 5 to 10 companies first.
5. Only run the full dataset after confirming the output is correct.

---

## Recommended Test Input

Before running thousands of records, create a small `companies.xlsx` file:

```text
company_name | location
Siemens      | Germany
Microsoft    | USA
Samsung      | South Korea
```

Run:

```bash
python google_places_company_info.py
```

Then check:

```text
google_company_info.xlsx
```

If the result is correct, run the full list.

---

## Final Result

At the end, this project gives you an Excel file containing unique companies and their Google business information.

The final result can be used for:

* Company enrichment
* Export/import company databases
* CRM preparation
* Market research
* Contact data collection
* Business intelligence dashboards

---

## License

This project is for internal data processing and research use.

Before using the collected data commercially, make sure your usage follows Google Maps Platform terms and local data protection regulations.
