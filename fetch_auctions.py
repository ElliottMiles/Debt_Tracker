"""
Fetches U.S. Treasury marketable-security auction history from TreasuryDirect
and keeps Treasury_Auction_History.xlsx up to date.

Safe to re-run: each run only requests auctions that could be new (based on
the latest issue date already saved). A record whose (cusip, auction_date)
is already in the sheet is normally skipped -- except when the existing row
is a "stub" (an auction that was announced but not yet held when it was
first fetched, so it has no results yet), in which case the row is updated
in place once real results are available.

Usage:
    python fetch_auctions.py

Dependencies (install in your own venv):
    pip install requests openpyxl
"""

import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path

import requests
from openpyxl import Workbook, load_workbook

API_URL = "https://www.treasurydirect.gov/TA_WS/securities/search"
EXCEL_PATH = Path(__file__).resolve().parent / "Treasury_Auction_History.xlsx"
SHEET_NAME = "Auctions"

# TreasuryDirect's structured auction records start around here.
EARLIEST_ISSUE_DATE = date(1980, 1, 1)

# IMPORTANT: the /search endpoint's startDate/endDate filter on ISSUE DATE,
# not auction date (verified empirically -- results stay strictly within the
# requested issueDate window even though auctionDate spills outside it, since
# auctions happen before their security is issued). On a re-run we therefore
# look back from the latest *issue* date we've saved, and look forward past
# today so we still catch securities that were already auctioned but haven't
# issued yet.
LOOKBACK_DAYS = 7
FUTURE_BUFFER_DAYS = 45

REQUEST_TIMEOUT = 30
MAX_RETRIES = 3
USER_AGENT = "Treasury-Auction-History-Fetcher/1.0 (personal project)"

# Spreadsheet column order and how to convert each field from the API's JSON.
# api_field is the exact key TreasuryDirect returns; kind controls type
# conversion ("str", "num", "date", "bool"). Keep this in sync with the
# header row already written into Treasury_Auction_History.xlsx.
COLUMNS = [
    ("cusip", "cusip", "str"),
    ("type", "type", "str"),
    ("security_term", "securityTerm", "str"),
    ("term", "term", "str"),
    ("reopening", "reopening", "bool"),
    ("announcement_date", "announcementDate", "date"),
    ("auction_date", "auctionDate", "date"),
    ("issue_date", "issueDate", "date"),
    ("dated_date", "datedDate", "date"),
    ("maturity_date", "maturityDate", "date"),
    ("interest_rate", "interestRate", "num"),
    ("high_yield", "highYield", "num"),
    ("high_discount_rate", "highDiscountRate", "num"),
    ("high_investment_rate", "highInvestmentRate", "num"),
    ("high_discount_margin", "highDiscountMargin", "num"),
    ("spread", "spread", "num"),
    ("price_per_100", "pricePer100", "num"),
    ("offering_amount", "offeringAmount", "num"),
    ("total_accepted", "totalAccepted", "num"),
    ("total_tendered", "totalTendered", "num"),
    ("bid_to_cover_ratio", "bidToCoverRatio", "num"),
    ("direct_bidder_accepted", "directBidderAccepted", "num"),
    ("indirect_bidder_accepted", "indirectBidderAccepted", "num"),
    ("primary_dealer_accepted", "primaryDealerAccepted", "num"),
    ("currently_outstanding", "currentlyOutstanding", "num"),
]
EXPECTED_HEADERS = [header for header, _, _ in COLUMNS]
TOTAL_ACCEPTED_IDX = EXPECTED_HEADERS.index("total_accepted")


def convert(raw_value, kind):
    """Convert one field from TreasuryDirect's JSON to a spreadsheet-ready value.
    TreasuryDirect uses "" (not null) for fields that don't apply to a given
    security type -- e.g. interestRate is blank for Bills."""
    if raw_value in (None, ""):
        return None
    if kind == "str":
        return raw_value
    if kind == "bool":
        return raw_value == "Yes"
    if kind == "num":
        try:
            return float(raw_value)
        except (TypeError, ValueError):
            return None
    if kind == "date":
        try:
            return datetime.fromisoformat(raw_value).date()
        except (TypeError, ValueError):
            return None
    raise ValueError(f"unknown column kind: {kind!r}")


def to_date(value):
    """Normalize a cell value (openpyxl may hand back datetime or date) to date."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return None


def fetch_records(start_date, end_date):
    """Fetch every auction whose issueDate falls in [start_date, end_date]."""
    params = {
        "format": "json",
        "startDate": start_date.isoformat(),
        "endDate": end_date.isoformat(),
    }
    headers = {"User-Agent": USER_AGENT}

    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(
                API_URL, params=params, headers=headers, timeout=REQUEST_TIMEOUT
            )
            resp.raise_for_status()
            data = resp.json()
            if not isinstance(data, list):
                raise ValueError(
                    f"unexpected response shape from TreasuryDirect: {type(data)}"
                )
            return data
        except (requests.RequestException, ValueError) as exc:
            last_error = exc
            if attempt == MAX_RETRIES:
                break
            wait = 2**attempt
            print(f"  request failed ({exc}); retrying in {wait}s...", file=sys.stderr)
            time.sleep(wait)

    raise RuntimeError(
        f"Failed to fetch auction data from TreasuryDirect after {MAX_RETRIES} attempts"
    ) from last_error


def load_workbook_state():
    """Open (or create) the sheet, validate its header row, and summarize
    what's already in it: for every (cusip, auction_date) already present,
    which sheet row it's on and whether that row is a "stub" (no results
    yet -- total_accepted is blank, meaning the auction hadn't happened yet
    the last time it was fetched); plus the latest issue_date seen."""
    if EXCEL_PATH.exists():
        wb = load_workbook(EXCEL_PATH)
        ws = wb[SHEET_NAME] if SHEET_NAME in wb.sheetnames else wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME

    header_row = [cell.value for cell in next(ws.iter_rows(min_row=1, max_row=1))] if ws.max_row >= 1 else []

    if not header_row or all(v is None for v in header_row):
        ws.append(EXPECTED_HEADERS)
        header_row = EXPECTED_HEADERS
    elif header_row != EXPECTED_HEADERS:
        raise RuntimeError(
            "Treasury_Auction_History.xlsx header row doesn't match what "
            "fetch_auctions.py expects. Expected:\n"
            f"  {EXPECTED_HEADERS}\nFound:\n  {header_row}\n"
            "Fix the header row (or clear the sheet to start fresh) and re-run."
        )

    cusip_col = header_row.index("cusip")
    auction_date_col = header_row.index("auction_date")
    issue_date_col = header_row.index("issue_date")
    total_accepted_col = header_row.index("total_accepted")

    existing_rows = {}  # (cusip, auction_date) -> sheet row number
    stub_keys = set()  # subset of existing_rows with no results yet
    max_issue_date = None
    for row_num, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
        if row[cusip_col] is None:
            continue
        auction_dt = to_date(row[auction_date_col])
        key = (row[cusip_col], auction_dt)
        existing_rows[key] = row_num
        if row[total_accepted_col] is None:
            stub_keys.add(key)

        issue_dt = to_date(row[issue_date_col])
        if issue_dt and (max_issue_date is None or issue_dt > max_issue_date):
            max_issue_date = issue_dt

    return wb, ws, existing_rows, stub_keys, max_issue_date


def main():
    wb, ws, existing_rows, stub_keys, max_issue_date = load_workbook_state()

    if max_issue_date is None:
        start_date = EARLIEST_ISSUE_DATE
        print(f"No existing data found. Fetching full auction history from {start_date}...")
    else:
        start_date = max_issue_date - timedelta(days=LOOKBACK_DAYS)
        print(
            f"Latest issue date on file: {max_issue_date}. "
            f"Fetching auctions with issue dates from {start_date} onward..."
        )

    end_date = date.today() + timedelta(days=FUTURE_BUFFER_DAYS)

    records = fetch_records(start_date, end_date)
    print(f"Retrieved {len(records)} auction record(s) from TreasuryDirect.")

    new_rows = 0
    updated_rows = 0
    skipped = 0
    for rec in records:
        cusip = rec.get("cusip")
        auction_dt = convert(rec.get("auctionDate"), "date")
        key = (cusip, auction_dt)
        row_values = [convert(rec.get(api_field), kind) for _, api_field, kind in COLUMNS]

        if key not in existing_rows:
            ws.append(row_values)
            existing_rows[key] = ws.max_row
            if row_values[TOTAL_ACCEPTED_IDX] is None:
                stub_keys.add(key)
            new_rows += 1
            continue

        # Already have this auction. If it was previously a stub (no results
        # yet) and this fetch has real results now, fill the row in rather
        # than leaving it permanently incomplete.
        if key in stub_keys and row_values[TOTAL_ACCEPTED_IDX] is not None:
            row_num = existing_rows[key]
            for col_idx, value in enumerate(row_values, start=1):
                ws.cell(row=row_num, column=col_idx, value=value)
            stub_keys.discard(key)
            updated_rows += 1
            continue

        skipped += 1

    if new_rows or updated_rows:
        wb.save(EXCEL_PATH)
        print(
            f"Added {new_rows} new record(s), filled in {updated_rows} "
            f"previously-incomplete record(s) ({skipped} already complete, skipped)."
        )
        print(f"Saved {EXCEL_PATH}")
    else:
        print(f"No changes ({skipped} already complete). Sheet is already up to date.")


if __name__ == "__main__":
    main()
