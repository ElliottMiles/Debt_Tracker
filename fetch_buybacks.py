"""
Fetches U.S. Treasury buyback operation results (CUSIP-level par amounts
Treasury has repurchased before maturity) and keeps Treasury_Buyback_History.csv
up to date.

This script only maintains the raw data file -- build_report.py doesn't read
it yet. It's a peer fetch script to fetch_auctions.py and fetch_yield_curve.py,
kept separate because the data source has nothing in common with either
(different API, different shape, no announcement/stub lifecycle to track).
See fetch_all.py to run all three together.

Source: api.fiscaldata.treasury.gov's "Treasury Securities Buybacks" dataset.
Unlike the yield curve (see fetch_yield_curve.py's note on that one), this
data *is* on the newer Fiscal Data JSON API -- confirmed by querying it
directly. Two tables are used:
  - buybacks_security_details: one row per CUSIP per operation -- cusip, par
    amount accepted, coupon, maturity date. This is the per-security data
    the rest of this file is built around.
  - buybacks_operations: one row per operation, used only to look up each
    operation's settlement_date (the date accepted debt is actually retired
    -- not published on the per-security table, since it's the same for
    every CUSIP in a given operation). Small enough (about 220 rows total,
    ever) to fetch in full every run and join in memory rather than
    filtering it incrementally.

Only nominal coupon Notes/Bonds and TIPS are eligible for buyback -- Bills,
FRNs, and STRIPS never appear in this feed.

Quirk: the API represents some missing values as the literal string "null"
(not JSON null or "") -- e.g. weighted_avg_accepted_price on a CUSIP that was
offered but not accepted. convert() treats all three the same.

Incremental: safe to re-run. With no existing file, this pulls full history
in one request -- the whole dataset is currently under 6,000 rows, well
inside this API's page size limit, though pagination is still implemented in
case that ever changes. A later run only re-requests operations from the
latest operation_date already on file, less LOOKBACK_DAYS, in case an
earlier run caught an operation before every CUSIP in it had posted --
overlapping rows are overwritten in place, not duplicated. Unlike
fetch_auctions.py, there's no "announced but not yet held" stub state to
handle here: this endpoint only ever publishes finished results (confirmed
by querying for operations dated after today -- empty).

Usage:
    python fetch_buybacks.py

Dependencies:
    pip install requests
"""

import csv
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import requests

API_BASE = "https://api.fiscaldata.treasury.gov/services/api/fiscal_service/v1/accounting/od"
SECURITY_DETAILS_ENDPOINT = f"{API_BASE}/buybacks_security_details"
OPERATIONS_ENDPOINT = f"{API_BASE}/buybacks_operations"
OUTPUT_PATH = Path(__file__).resolve().parent / "Treasury_Buyback_History.csv"

# Mirrors fetch_auctions.py's LOOKBACK_DAYS -- re-request a few days of
# already-saved operations on every incremental run, in case an earlier run
# caught an operation before all of its CUSIPs had posted.
LOOKBACK_DAYS = 7

PAGE_SIZE = 10000  # comfortably above the whole dataset's current size (~6,000 rows)
REQUEST_TIMEOUT = 30
MAX_RETRIES = 3
USER_AGENT = "Treasury-Buyback-History-Fetcher/1.0 (personal project)"

# Output column order and how to convert each field from the API's JSON.
# api_field is the exact key the Fiscal Data API returns; kind controls type
# conversion ("str", "num", "date"). settlement_date isn't on the source
# table (see module docstring) -- it's joined in from buybacks_operations by
# operation_date before conversion. cusip (not the source's own cusip_nbr) is
# used as the output column name to match Treasury_Auction_History.xlsx's
# own "cusip" column, so a future join between the two files is a plain
# same-name key.
COLUMNS = [
    ("operation_date", "operation_date", "date"),
    ("settlement_date", "settlement_date", "date"),
    ("cusip", "cusip_nbr", "str"),
    ("coupon_rate_pct", "coupon_rate_pct", "num"),
    ("maturity_date", "maturity_date", "date"),
    ("par_amt_accepted", "par_amt_accepted", "num"),
    ("weighted_avg_accepted_price", "weighted_avg_accepted_price", "num"),
]
HEADERS_OUT = [header for header, _, _ in COLUMNS]


def convert(raw_value, kind):
    """Convert one field from the Fiscal Data API's JSON to a CSV-ready value.
    The API represents some missing values as the literal string "null"
    (not JSON null or "") -- treat all three the same as missing."""
    if raw_value in (None, "", "null"):
        return None
    if kind == "str":
        return raw_value
    if kind == "num":
        try:
            return float(raw_value)
        except (TypeError, ValueError):
            return None
    if kind == "date":
        try:
            return datetime.strptime(raw_value, "%Y-%m-%d").date()
        except (TypeError, ValueError):
            return None
    raise ValueError(f"unknown column kind: {kind!r}")


def fetch_pages(endpoint, params):
    """Fetch every page of one Fiscal Data API endpoint for the given base
    params, returning the concatenated list of raw records."""
    headers = {"User-Agent": USER_AGENT}
    all_records = []
    page_number = 1
    while True:
        page_params = {**params, "page[size]": PAGE_SIZE, "page[number]": page_number}

        last_error = None
        payload = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = requests.get(endpoint, params=page_params, headers=headers, timeout=REQUEST_TIMEOUT)
                resp.raise_for_status()
                payload = resp.json()
                break
            except (requests.RequestException, ValueError) as exc:
                last_error = exc
                if attempt == MAX_RETRIES:
                    raise RuntimeError(
                        f"Failed to fetch {endpoint} after {MAX_RETRIES} attempts"
                    ) from last_error
                wait = 2**attempt
                print(f"  request failed ({exc}); retrying in {wait}s...", file=sys.stderr)
                time.sleep(wait)

        all_records.extend(payload.get("data", []))

        total_pages = payload.get("meta", {}).get("total-pages", 1)
        if page_number >= total_pages:
            break
        page_number += 1

    return all_records


def fetch_settlement_dates():
    """operation_date -> settlement_date for every buyback operation ever
    held. The whole table is small enough to always fetch in full (see
    module docstring) rather than tracked incrementally."""
    records = fetch_pages(OPERATIONS_ENDPOINT, {"fields": "operation_date,settlement_date"})
    return {rec["operation_date"]: rec.get("settlement_date") for rec in records}


def fetch_security_details(start_date):
    """Every CUSIP-level buyback result with operation_date >= start_date."""
    params = {
        "fields": "operation_date,cusip_nbr,coupon_rate_pct,maturity_date,par_amt_accepted,weighted_avg_accepted_price",
        "filter": f"operation_date:gte:{start_date.isoformat()}",
        "sort": "operation_date",
    }
    return fetch_pages(SECURITY_DETAILS_ENDPOINT, params)


def load_existing():
    """Read the existing CSV, if any, into {(operation_date, cusip): row_dict}.
    Empty dict if the file doesn't exist yet."""
    if not OUTPUT_PATH.exists():
        return {}
    with OUTPUT_PATH.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames != HEADERS_OUT:
            raise RuntimeError(
                "Treasury_Buyback_History.csv header row doesn't match what "
                "fetch_buybacks.py expects. Expected:\n"
                f"  {HEADERS_OUT}\nFound:\n  {reader.fieldnames}\n"
                "Fix the header row (or clear the file to start fresh) and re-run."
            )
        # Parse every stored cell back into the same Python types convert()
        # produces (date objects, floats, None) rather than leaving them as
        # the raw strings csv.DictReader hands back -- otherwise a row
        # reloaded from disk (all-string values) would never compare equal
        # to, or even share a dict key with, the freshly-fetched version of
        # the same record (a real date vs. its string repr, e.g.), silently
        # breaking dedup and the "updated" count.
        kind_by_header = {header: kind for header, _, kind in COLUMNS}
        rows = {}
        for raw_row in reader:
            row = {header: convert(raw_row[header], kind_by_header[header]) for header in HEADERS_OUT}
            rows[(row["operation_date"], row["cusip"])] = row
    return rows


def should_replace(old_row, new_row):
    """Whether new_row should overwrite old_row for the same (operation_date,
    cusip) key. Almost always yes -- a freshly-fetched row is authoritative,
    whether it's replacing an older run's data or a duplicate seen earlier in
    this same fetch. The one exception, confirmed empirically: 26 rows (all
    dated 2024-08-21, apparently a one-time source glitch) are published
    *twice* under the same key -- once with a real par_amt_accepted, once
    with the "null" placeholder (see convert()'s docstring) standing in for
    it. Never let the null variant clobber a real value, regardless of which
    one the API happens to list first."""
    if old_row is None:
        return True
    if new_row["par_amt_accepted"] is None and old_row["par_amt_accepted"] is not None:
        return False
    return True


def format_for_csv(row):
    """Typed row (date objects, floats, None) -> plain strings for csv.DictWriter."""
    return {
        header: "" if value is None else (value.isoformat() if isinstance(value, date) else value)
        for header, value in row.items()
    }


def main():
    existing_rows = load_existing()

    if not existing_rows:
        start_date = date(2000, 1, 1)  # program's actual start is 2000-03-09; rounds down to be safe
        print(f"No existing data found. Fetching full buyback history from {start_date}...")
    else:
        max_operation_date = max(op_date for op_date, _ in existing_rows)
        start_date = max_operation_date - timedelta(days=LOOKBACK_DAYS)
        print(
            f"Latest operation date on file: {max_operation_date}. "
            f"Fetching operations from {start_date} onward..."
        )

    settlement_by_operation = fetch_settlement_dates()
    records = fetch_security_details(start_date)
    print(f"Retrieved {len(records)} buyback record(s) from Fiscal Data.")

    net_new = 0
    updated = 0
    for rec in records:
        op_date_raw = rec.get("operation_date")
        rec["settlement_date"] = settlement_by_operation.get(op_date_raw)
        row = {header: convert(rec.get(api_field), kind) for header, api_field, kind in COLUMNS}

        key = (row["operation_date"], row["cusip"])
        old_row = existing_rows.get(key)
        if not should_replace(old_row, row):
            continue
        if old_row is None:
            net_new += 1
        elif old_row != row:
            updated += 1
        existing_rows[key] = row

    if net_new or updated:
        sorted_keys = sorted(existing_rows)
        with OUTPUT_PATH.open("w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=HEADERS_OUT)
            writer.writeheader()
            for key in sorted_keys:
                writer.writerow(format_for_csv(existing_rows[key]))
        print(f"Added {net_new} new record(s), updated {updated} existing record(s).")
        print(f"{len(existing_rows):,} total record(s). Saved {OUTPUT_PATH.name}")
    else:
        print(f"No changes ({len(existing_rows):,} record(s) already up to date).")


if __name__ == "__main__":
    main()
