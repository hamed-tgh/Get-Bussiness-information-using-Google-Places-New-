import os
import time
import csv
import random
from pathlib import Path

import requests
import pandas as pd
from tqdm import tqdm


# =========================
# CONFIG
# =========================

API_KEY = os.getenv("Scrapper") or os.getenv("GOOGLE_MAPS_API_KEY")

if not API_KEY:
    raise ValueError(
        "API key is not set. Please set environment variable 'Scrapper' or 'GOOGLE_MAPS_API_KEY'."
    )

INPUT_FILE = "companies.xlsx"
OUTPUT_FILE = "google_company_info.xlsx"
CHECKPOINT_FILE = "google_company_info_checkpoint.csv"

SEARCH_URL = "https://places.googleapis.com/v1/places:searchText"

REQUEST_DELAY_SECONDS = 0.2
SAVE_EXCEL_EVERY = 25

MAX_RETRIES = 4
RETRY_BASE_SECONDS = 2

RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}

FIELD_MASK = ",".join([
    "places.id",
    "places.displayName",
    "places.formattedAddress",
    "places.nationalPhoneNumber",
    "places.internationalPhoneNumber",
    "places.regularOpeningHours",
    "places.currentOpeningHours",
    "places.websiteUri",
    "places.googleMapsUri",
    "places.businessStatus",
])


OUTPUT_COLUMNS = [
    "unique_key",
    "input_company",
    "input_location",
    "search_query",
    "status",
    "google_name",
    "address",
    "phone_number",
    "international_phone_number",
    "hours",
    "website",
    "google_maps_url",
    "business_status",
    "place_id",
]




def normalize_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()

# in order to save space in checkpoint and excel, we create a unique key based on company name and location. 
def make_unique_key(company_name: str, location: str) -> str:
    """
    This key is used to know whether this company/location was already checked.
    """
    company_name = normalize_text(company_name).lower()
    location = normalize_text(location).lower()
    return f"{company_name}|||{location}"

#if the company is not found or there is an error, we return this empty result with status. This way we can save all attempts in checkpoint and know which ones were successful, which were not found, and which had errors. We can also retry errors in the next run, while skipping found/not found.
def empty_result(
    unique_key: str,
    company_name: str,
    location: str,
    query: str,
    status: str
) -> dict:
    return {
        "unique_key": unique_key,
        "input_company": company_name,
        "input_location": location,
        "search_query": query,
        "status": status,
        "google_name": "",
        "address": "",
        "phone_number": "",
        "international_phone_number": "",
        "hours": "",
        "website": "",
        "google_maps_url": "",
        "business_status": "",
        "place_id": "",
    }


# this function appends each result to a CSV checkpoint file immediately after processing each company. This way, if the program crashes or is stopped, we won't lose all progress and can resume from the last saved state. The checkpoint file will contain all attempts, including found, not found, and errors, which allows us to analyze results and retry errors in future runs.
def append_result_to_checkpoint(result: dict):
    """
    Save one row immediately.
    This protects your data if the program crashes.
    """
    file_exists = Path(CHECKPOINT_FILE).exists()

    with open(CHECKPOINT_FILE, mode="a", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=OUTPUT_COLUMNS)

        if not file_exists:
            writer.writeheader()

        writer.writerow(result)
        f.flush()

# this function is tend to load the existing checkpoint file and return a set of unique keys that were already processed with status "Found" or "Not found". This allows the main program to skip those records in future runs, while still allowing records with errors to be retried. By keeping track of completed keys, we can avoid redundant API calls and save time and resources.
def load_existing_completed_keys() -> set:
    """
    Load already completed companies from checkpoint.
    Only Found / Not found are considered completed.
    Errors are not considered completed, so they can be retried in the next run.
    """
    if not Path(CHECKPOINT_FILE).exists():
        return set()

    df = pd.read_csv(CHECKPOINT_FILE, dtype=str).fillna("")

    completed_statuses = {
        "Found",
        "Not found",
    }

    completed_df = df[df["status"].isin(completed_statuses)]

    return set(completed_df["unique_key"].astype(str).tolist())


def save_checkpoint_to_excel():
    """
    Rebuild Excel from checkpoint CSV.
    It keeps the last result for each unique company/location.
    """
    if not Path(CHECKPOINT_FILE).exists():
        return

    df = pd.read_csv(CHECKPOINT_FILE, dtype=str).fillna("")

    if df.empty:
        return

    # If same company was retried, keep the latest row
    df = df.drop_duplicates(subset=["unique_key"], keep="last")

    
    summary_df = (
        df.groupby("status")
        .size()
        .reset_index(name="count")
        .sort_values("count", ascending=False)
    )

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Google Info", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)



# this function performs the actual search for a company using the Google Places API. It constructs the search query based on the company name and location, makes the API request, and processes the response. If the company is found, it extracts relevant information such as address, phone number, hours, website, etc. If there are errors or if the company is not found, it returns an appropriate status. The function also implements retry logic for handling temporary errors and rate limits.
def search_company(company_name: str, location: str = "") -> dict:
    company_name = normalize_text(company_name)
    location = normalize_text(location)

    unique_key = make_unique_key(company_name, location)

    query = company_name

    if location:
        query = f"{company_name} {location}".strip()

    headers = {
        "Content-Type": "application/json",
        "X-Goog-Api-Key": API_KEY,
        "X-Goog-FieldMask": FIELD_MASK,
    }

    payload = {
        "textQuery": query,
        "maxResultCount": 1,
        "languageCode": "en",
    }

    last_error = ""

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.post(
                SEARCH_URL,
                headers=headers,
                json=payload,
                timeout=30
            )

            status_code = response.status_code

            # Retry for temporary Google/server/rate-limit errors
            if status_code in RETRYABLE_STATUS_CODES:
                last_error = f"Retryable HTTP {status_code}: {response.text[:300]}"

                sleep_time = RETRY_BASE_SECONDS * attempt + random.uniform(0, 1.5)
                time.sleep(sleep_time)
                continue

            # Non-retryable HTTP error
            if status_code >= 400:
                return empty_result(
                    unique_key,
                    company_name,
                    location,
                    query,
                    f"HTTP error {status_code}: {response.text[:300]}"
                )

            data = response.json()
            places = data.get("places", [])

            if not places:
                return empty_result(
                    unique_key,
                    company_name,
                    location,
                    query,
                    "Not found"
                )

            place = places[0]

            display_name = place.get("displayName", {}).get("text", "")
            address = place.get("formattedAddress", "")
            national_phone = place.get("nationalPhoneNumber", "")
            international_phone = place.get("internationalPhoneNumber", "")
            website = place.get("websiteUri", "")
            maps_url = place.get("googleMapsUri", "")
            business_status = place.get("businessStatus", "")
            place_id = place.get("id", "")

            regular_hours = place.get("regularOpeningHours", {})
            current_hours = place.get("currentOpeningHours", {})

            hours_list = (
                regular_hours.get("weekdayDescriptions")
                or current_hours.get("weekdayDescriptions")
                or []
            )

            hours_text = "\n".join(hours_list)

            return {
                "unique_key": unique_key,
                "input_company": company_name,
                "input_location": location,
                "search_query": query,
                "status": "Found",
                "google_name": display_name,
                "address": address,
                "phone_number": national_phone,
                "international_phone_number": international_phone,
                "hours": hours_text,
                "website": website,
                "google_maps_url": maps_url,
                "business_status": business_status,
                "place_id": place_id,
            }

        except requests.exceptions.Timeout:
            last_error = "Timeout error"

        except requests.exceptions.ConnectionError:
            last_error = "Connection error"

        except requests.exceptions.RequestException as e:
            last_error = f"Request error: {e}"

        except Exception as e:
            last_error = f"Unexpected error: {e}"

        sleep_time = RETRY_BASE_SECONDS * attempt + random.uniform(0, 1.5)
        time.sleep(sleep_time)

    return empty_result(
        unique_key,
        company_name,
        location,
        query,
        f"Failed after retries: {last_error}"
    )



def main():
    input_df = pd.read_excel(INPUT_FILE)

    if "company_name" not in input_df.columns:
        raise ValueError("Input Excel must contain a column named 'company_name'.")

    if "location" not in input_df.columns:
        input_df["location"] = ""

    completed_keys = load_existing_completed_keys()

    print(f"Already completed records: {len(completed_keys)}")

    processed_counter = 0
    skipped_counter = 0

    for _, row in tqdm(input_df.iterrows(), total=len(input_df)):
        company_name = normalize_text(row.get("company_name", ""))
        location = normalize_text(row.get("location", ""))

        if not company_name:
            continue

        unique_key = make_unique_key(company_name, location)

        if unique_key in completed_keys:
            skipped_counter += 1
            continue

        result = search_company(company_name, location)

        append_result_to_checkpoint(result)

        if result["status"] in {"Found", "Not found"}:
            completed_keys.add(unique_key)

        processed_counter += 1

        if processed_counter % SAVE_EXCEL_EVERY == 0:
            save_checkpoint_to_excel()

        time.sleep(REQUEST_DELAY_SECONDS)

    save_checkpoint_to_excel()

    print("Finished.")
    print(f"Processed new records: {processed_counter}")
    print(f"Skipped existing records: {skipped_counter}")
    print(f"Checkpoint file: {CHECKPOINT_FILE}")
    print(f"Excel file: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()