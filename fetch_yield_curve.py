"""
Fetches the U.S. Treasury Daily Par Yield Curve Rates (the "Constant Maturity
Treasury" series -- 1 Month through 30 Year, one row per business day) and
saves it to Yield_Curve_History.csv.

Not currently read by build_report.py or shown in the report -- this is a
peer fetch script to fetch_auctions.py, kept separate because the two data
sources have nothing in common (different API, different pagination and
incremental strategy, different output shape). See fetch_all.py to run both.

Source note: this dataset is NOT available on the newer JSON API at
api.fiscaldata.treasury.gov (checked directly -- 404). It's only published
through home.treasury.gov's older per-year CSV/XML interface, confirmed by
inspecting a real response. That means one HTTP request per calendar year,
not one request for full history like the TA_WS auction endpoint allows --
and unlike that endpoint, there's no bulk/multi-year option here either
(checked: a path with no year 403s, and a year-range value is silently
ignored in favor of the single year in the URL path).

Concurrent fetching: each request to this endpoint costs a near-fixed ~9s at
Treasury's CDN edge layer regardless of how much data comes back (confirmed
via the `server-timing` response header -- the actual origin answers in
~140ms; the rest is edge cache revalidation), so with 37 requests for a full
history pull, that fixed cost dominates and the run is dominated by *request
count*, not data volume or network throughput. Since the origin isn't the
bottleneck, years are fetched several at a time with a thread pool instead of
one at a time, so multiple ~9s edge waits overlap instead of stacking up
sequentially. MAX_WORKERS is a measured choice, not a guess: a small pool (5) was tested
against a much larger one (37, i.e. one thread per year, no cap) on
comparable fresh year batches, and the larger pool came back faster *and*
cleaner (0 retries vs. several). There's also no sign this endpoint is
throttling by request volume in the first place -- every response, at any
concurrency level, has come back a plain 200 with no 429, no Retry-After, and
no rate-limit headers. 15 is a middle ground: fast, and -- unlike using one
thread per year -- doesn't quietly send more and more simultaneous
connections as full history keeps growing by a year every year this runs.
Output order doesn't depend on fetch order -- results are always sorted by
date before being written, regardless of which years' requests happened to
finish first (see the note by the CSV-writing code below).

Incremental: this endpoint can't be asked for "just the new rows" -- the
finest granularity it offers is a whole calendar year -- but past years never
change once they're over, so there's no need to re-fetch them once they're
already saved. If Yield_Curve_History.csv doesn't exist yet, this pulls full
history (one request per year, ~37 requests). On every later run, it only
re-fetches the latest year already in the file (in case that run happened
mid-year and missed some trailing days) through the current year -- normally
just 1 request instead of ~37, since most runs happen within the same year as
the last one.

The set of maturities published has changed over time -- 1 Mo starts in 2001
(when the 4-week bill was introduced), 2 Mo starts in 2018, and 20/30 Yr
briefly disappear from the header entirely during the 2002-2006 30-Year Bond
suspension. This script doesn't assume a fixed column set; it unions whatever
columns actually appear across the years it fetches.

Usage:
    python fetch_yield_curve.py [start_year]

    start_year only matters on the very first run (no existing CSV) -- it
    sets how far back to go, and defaults to 1990, the earliest year this
    dataset covers (years before that return no data). Ignored on later
    incremental runs.

Dependencies:
    pip install requests
"""

import csv
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime
from pathlib import Path

import requests

OUTPUT_PATH = Path(__file__).resolve().parent / "Yield_Curve_History.csv"
DEFAULT_START_YEAR = 1990
REQUEST_TIMEOUT = 25
MAX_RETRIES = 3
# See the module docstring's "Concurrent fetching" section for how this
# number was chosen -- tested against a much larger pool, not guessed.
MAX_WORKERS = 15
USER_AGENT = "Treasury-Yield-Curve-Fetcher/1.0 (personal project)"

# Sensible fixed output-column order. A given row may be blank for any of
# these -- either the maturity didn't exist yet, was temporarily
# discontinued, or that specific day's value just wasn't published.
MATURITY_ORDER = [
    "1 Mo", "1.5 Month", "2 Mo", "3 Mo", "4 Mo", "6 Mo",
    "1 Yr", "2 Yr", "3 Yr", "5 Yr", "7 Yr", "10 Yr", "20 Yr", "30 Yr",
]


def fetch_year(year):
    """Fetch one calendar year of daily par yield curve rates.
    Returns a list of (date, {maturity: value_or_None}) tuples, or [] if the
    year has no data (e.g. before 1990) or the request ultimately fails."""
    url = (
        "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/"
        f"daily-treasury-rates.csv/{year}/all"
        f"?type=daily_treasury_yield_curve&field_tdr_date_value={year}&page&_format=csv"
    )
    headers = {"User-Agent": USER_AGENT}

    last_error = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = requests.get(url, headers=headers, timeout=REQUEST_TIMEOUT)
            resp.raise_for_status()
            text = resp.text.strip()
            if not text:
                return []
            reader = csv.DictReader(text.splitlines())
            rows = []
            for row in reader:
                date_str = row.pop("Date", None)
                if not date_str:
                    continue
                parsed_date = datetime.strptime(date_str, "%m/%d/%Y").date()
                values = {k.strip(): (v.strip() if v and v.strip() else None) for k, v in row.items()}
                rows.append((parsed_date, values))
            return rows
        except (requests.RequestException, ValueError) as exc:
            last_error = exc
            if attempt == MAX_RETRIES:
                print(f"  {year}: failed after {MAX_RETRIES} attempts ({exc})", file=sys.stderr)
                return []
            wait = 2**attempt
            print(f"  {year}: request failed ({exc}); retrying in {wait}s...", file=sys.stderr)
            time.sleep(wait)
    return []


def load_existing():
    """Read the existing CSV, if any, into ({date: {maturity: value}}, columns).
    columns is [] and the dict is empty if the file doesn't exist yet."""
    if not OUTPUT_PATH.exists():
        return {}, []
    with OUTPUT_PATH.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        columns = [c for c in reader.fieldnames if c != "date"]
        rows = {}
        for row in reader:
            d = date.fromisoformat(row["date"])
            rows[d] = {c: (row.get(c) or None) for c in columns}
    return rows, columns


def main():
    cli_start_year = int(sys.argv[1]) if len(sys.argv) > 1 else DEFAULT_START_YEAR
    end_year = date.today().year

    existing_rows, existing_columns = load_existing()

    if not existing_rows:
        fetch_start_year = cli_start_year
        print(f"No existing file found. Fetching full history, {fetch_start_year}-{end_year}...")
    else:
        # Re-fetch the latest year already saved (that run may have happened
        # mid-year and missed later days) through the current year. Every
        # earlier year is untouched -- it's already complete and can't change.
        fetch_start_year = max(existing_rows).year
        print(
            f"Existing file found, latest date {max(existing_rows)}. "
            f"Re-fetching {fetch_start_year}-{end_year} ({end_year - fetch_start_year + 1} year(s))..."
        )

    t0 = time.time()

    # Keep years strictly before the re-fetch range as-is; years in range are
    # replaced wholesale by the fresh fetch below, not merged row-by-row --
    # simpler, and correct even in the (very unlikely) case a source row gets
    # corrected or removed upstream.
    all_rows = {d: v for d, v in existing_rows.items() if d.year < fetch_start_year}
    seen_columns = set(existing_columns)
    years_requested = range(fetch_start_year, end_year + 1)
    years_with_data = 0
    print_lock = threading.Lock()  # print() itself is fine to call from multiple threads, but
                                    # without a lock two threads' start/timing/row-count lines
                                    # could still land interleaved mid-line; the lock keeps each
                                    # progress line intact as one unbroken line of output.

    def fetch_with_progress(year):
        year_t0 = time.time()
        year_rows = fetch_year(year)
        year_elapsed = time.time() - year_t0
        with print_lock:
            # flush=True: a full run can take a while even in parallel, and
            # without flushing, a non-interactive stdout (piped, logged, run
            # via fetch_all.py) would buffer every line here until the whole
            # pool finishes -- silent for a while, indistinguishable from a
            # hang.
            print(f"  {year}: {len(year_rows)} row(s) in {year_elapsed:.1f}s", flush=True)
        return year, year_rows

    results_by_year = {}
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        futures = [executor.submit(fetch_with_progress, year) for year in years_requested]
        for future in as_completed(futures):
            year, year_rows = future.result()
            results_by_year[year] = year_rows

    # Threads finish in whatever order their requests happen to complete, not
    # necessarily year order -- results_by_year is keyed by year regardless,
    # so this merge (and the sort just below, before writing) is unaffected
    # either way. The CSV is always written in date order because it's built
    # from `sorted(all_rows.keys())`, never from insertion/completion order.
    for year in years_requested:
        year_rows = results_by_year.get(year, [])
        if year_rows:
            years_with_data += 1
        for d, values in year_rows:
            all_rows[d] = values
            seen_columns.update(values.keys())

    elapsed = time.time() - t0

    if not all_rows:
        print("No data fetched -- check connectivity or whether the endpoint has changed.")
        return

    columns = [c for c in MATURITY_ORDER if c in seen_columns]
    unexpected = seen_columns - set(columns)
    if unexpected:
        print(f"Note: column(s) not in MATURITY_ORDER, appended at the end: {sorted(unexpected)}")
        columns += sorted(unexpected)

    sorted_dates = sorted(all_rows.keys())
    net_new_rows = len(all_rows) - len(existing_rows)

    with OUTPUT_PATH.open("w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["date"] + columns)
        for d in sorted_dates:
            values = all_rows[d]
            writer.writerow([d.isoformat()] + [values.get(c) or "" for c in columns])

    print(f"\nDone in {elapsed:.1f}s -- {years_with_data}/{len(years_requested)} requested year(s) returned data.")
    if existing_rows:
        print(f"{net_new_rows:,} new row(s) added ({len(sorted_dates):,} total).")
    print(f"{len(sorted_dates):,} trading days, {sorted_dates[0]} to {sorted_dates[-1]}.")
    print(f"Columns ({len(columns)}): {columns}")
    print(f"Saved to {OUTPUT_PATH.name}")

    print("\nMost recent 3 rows:")
    for d in sorted_dates[-3:]:
        values = all_rows[d]
        preview = ", ".join(f"{c}={values[c]}" for c in columns if values.get(c))
        print(f"  {d}: {preview}")


if __name__ == "__main__":
    main()
