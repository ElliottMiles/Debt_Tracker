"""
Reads Treasury_Auction_History.xlsx, computes U.S. marketable-debt structure
metrics, and writes Debt_Report.html from report_template.html.

Every run fully regenerates Debt_Report.html from scratch -- it never reads
or patches the previous output. report_template.html is never modified; only
Debt_Report.html (a separate file) is written.

Usage:
    python build_report.py

Dependencies (install in your own venv):
    pip install pandas openpyxl
"""

import json
from datetime import date, datetime
from pathlib import Path

import pandas as pd

EXCEL_PATH = Path(__file__).resolve().parent / "Treasury_Auction_History.xlsx"
YIELD_CURVE_PATH = Path(__file__).resolve().parent / "Yield_Curve_History.csv"
TEMPLATE_PATH = Path(__file__).resolve().parent / "report_template.html"
OUTPUT_PATH = Path(__file__).resolve().parent / "Debt_Report.html"
SHEET_NAME = "Auctions"
DATA_PLACEHOLDER = "/*__REPORT_DATA__*/"

# Fixed order everywhere a security type appears (chart series, legends,
# tables) so color assignment stays consistent report-wide.
TYPE_ORDER = ["Bill", "Note", "Bond", "TIPS", "FRN", "CMB"]

# (label, min_days_inclusive, max_days_exclusive)
MATURITY_BUCKETS = [
    ("<1wk", 0, 7),
    ("1wk-1mo", 7, 30),
    ("1-3mo", 30, 91),
    ("3-6mo", 91, 182),
    ("6-12mo", 182, 365),
    ("1-2yr", 365, 730),
    ("2-5yr", 730, 1825),
    ("5-10yr", 1825, 3650),
    ("10-20yr", 3650, 7300),
    ("20-30yr", 7300, 1_000_000),
]
BUCKET_LABELS = [b[0] for b in MATURITY_BUCKETS]

# The 7 conventional "yield curve" benchmark maturities, for the by-maturity
# auction history chart. Keyed by (type, term) rather than term alone --
# TIPS and FRN reuse the same term labels as Notes/Bonds (see the note by
# latest_rate_by_type_term below), so term alone would silently pull in a
# TIPS real yield under a "10-Year" label meant to mean the nominal Note.
DURATION_DEFS = [
    ("1 Month", "Bill", "4-Week"),
    ("3 Month", "Bill", "13-Week"),
    ("1 Year", "Bill", "52-Week"),
    ("2 Year", "Note", "2-Year"),
    ("5 Year", "Note", "5-Year"),
    ("10 Year", "Note", "10-Year"),
    ("30 Year", "Bond", "30-Year"),
]

# Earliest date TreasuryDirect's structured auction records go back to -- also
# where the "Show history" sparklines start (see historical_snapshots below).
HISTORY_START = date(1980, 1, 1)

# Same 7 benchmark maturities as DURATION_DEFS, same display labels -- mapped
# here to Yield_Curve_History.csv's own column names (fetch_yield_curve.py's
# source labels its columns "1 Mo"/"3 Mo"/etc., not the same strings used for
# the auction-history side of the app).
YIELD_CURVE_DURATIONS = [
    ("1 Month", "1 Mo"),
    ("3 Month", "3 Mo"),
    ("1 Year", "1 Yr"),
    ("2 Year", "2 Yr"),
    ("5 Year", "5 Yr"),
    ("10 Year", "10 Yr"),
    ("30 Year", "30 Yr"),
]


def bucket_for_days(days):
    for label, lo, hi in MATURITY_BUCKETS:
        if lo <= days < hi:
            return label
    return BUCKET_LABELS[-1]


def load_auctions():
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME)
    for col in ("announcement_date", "auction_date", "issue_date", "dated_date", "maturity_date"):
        df[col] = pd.to_datetime(df[col])
    # Stub rows (auction announced but not yet held) carry no results yet --
    # they're not part of any debt-structure calculation.
    return df[df["total_accepted"].notna()].copy()


def load_yield_curve():
    """Read Yield_Curve_History.csv (from fetch_yield_curve.py) into a plain
    dict shaped for direct use by the yield-curve slider: parallel arrays,
    one value per trading day per benchmark maturity, aligned by index so the
    frontend can pick "day i" and read every maturity's value for it without
    any date lookups of its own."""
    if not YIELD_CURVE_PATH.exists():
        raise FileNotFoundError(
            f"{YIELD_CURVE_PATH.name} not found. Run fetch_yield_curve.py first."
        )
    df = pd.read_csv(YIELD_CURVE_PATH, parse_dates=["date"]).sort_values("date")
    dates = [iso_date(d) for d in df["date"]]
    values = {}
    for label, column in YIELD_CURVE_DURATIONS:
        if column in df.columns:
            values[label] = [None if pd.isna(v) else round(float(v), 3) for v in df[column]]
        else:
            values[label] = [None] * len(df)
    return {
        "dates": dates,
        "durations": [label for label, _ in YIELD_CURVE_DURATIONS],
        "values": values,
    }


def compute_effective_rate(df):
    """A type-appropriate annualized rate, comparable across security types.

    Note/Bond/TIPS -> auction high_yield (reflects price actually paid).
    Bill/CMB       -> high_investment_rate (bond-equivalent yield; comparable
                      basis to high_yield, unlike the traditional discount rate).
    FRN            -> spread + the most recent 13-week Bill's high_investment_rate
                      as of that FRN's own auction date. Approximation: an FRN's
                      real coupon resets weekly against the 13-week Bill discount
                      rate, so this is a snapshot, not its literal cash flow.
    """
    rate = pd.Series(index=df.index, dtype="float64")

    note_like = df["type"].isin(["Note", "Bond", "TIPS"])
    rate[note_like] = df.loc[note_like, "high_yield"]

    bill_like = df["type"].isin(["Bill", "CMB"])
    rate[bill_like] = df.loc[bill_like, "high_investment_rate"]

    is_frn = df["type"] == "FRN"
    if is_frn.any():
        ref = (
            df[(df["type"] == "Bill") & (df["term"] == "13-Week")]
            [["auction_date", "high_investment_rate"]]
            .dropna()
            .sort_values("auction_date")
        )
        frn_rows = df.loc[is_frn, ["auction_date", "spread"]].sort_values("auction_date")
        merged = pd.merge_asof(frn_rows, ref, on="auction_date", direction="backward")
        # merge_asof always returns a fresh 0..n-1 index, discarding frn_rows'
        # original index -- restore it (row order/count is preserved 1:1) so
        # this lines up with `rate`, which is still indexed like `df`.
        merged.index = frn_rows.index
        rate.loc[merged.index] = merged["high_investment_rate"].to_numpy() + merged["spread"].to_numpy()

    return rate


def compute_cash_interest_rate(df, effective_rate):
    """The rate that determines actual cash interest obligations -- used only
    for the annual-interest-cost figures, never for the rate/cost-of-borrowing
    comparisons elsewhere in this app (those all use effective_rate_pct).

    Note/Bond/TIPS -> the literal coupon (interest_rate): the fixed cash
                      payment made twice a year for the security's whole
                      life, set at original issuance and unrelated to the
                      price actually paid at auction. This is where it
                      differs from effective_rate_pct's high_yield, which
                      reflects price paid and is the right basis for
                      cross-time/cross-issue comparisons -- but a reopening
                      priced away from par (routine, not an edge case: ~95%
                      of outstanding Note/Bond/TIPS records at the time this
                      was written) means yield isn't what's actually
                      disbursed in cash.
    Bill/CMB, FRN  -> same as effective_rate_pct. Bills/CMB have no coupon at
                      all (the discount to face value *is* the return, so
                      there's nothing else to use); FRN's rate genuinely
                      floats and effective_rate_pct is already the best
                      available estimate of what it's currently paying.
    """
    rate = effective_rate.copy()
    note_like = df["type"].isin(["Note", "Bond", "TIPS"])
    rate[note_like] = df.loc[note_like, "interest_rate"]
    return rate


def weighted_avg(amounts, rates):
    weight = amounts.where(rates.notna())
    total = weight.sum()
    if not total or pd.isna(total):
        return None
    return float((weight * rates).sum() / total)


def iso_date(value):
    return None if pd.isna(value) else value.strftime("%Y-%m-%d")


def month_start_dates(start, end):
    """First-of-month timestamps from `start` through `end`, inclusive."""
    return list(pd.date_range(start=pd.Timestamp(start), end=pd.Timestamp(end), freq="MS"))


def historical_snapshots(df, dates):
    """Point-in-time replay of the five headline metrics.

    For each date D, use only securities issued by then (issue_date <= D)
    that hadn't yet matured as of D (maturity_date > D) -- i.e. exactly what
    this dashboard would have shown on D, using the same definitions as the
    live summary tiles. Each security's own effective_rate_pct and
    cash_interest_rate_pct never change (both are fixed at auction); only
    which securities are "outstanding as of D" and their years-to-maturity
    as of D change per snapshot. weighted_avg_rate_pct uses effective_rate_pct
    (yield, for cost-of-borrowing comparability); annual_interest_cost uses
    cash_interest_rate_pct (coupon for Note/Bond/TIPS, i.e. actual cash
    obligations) -- see compute_cash_interest_rate for why they differ.
    """
    history = {
        "dates": [], "total_outstanding": [], "weighted_avg_rate_pct": [], "annual_interest_cost": [],
        "weighted_avg_maturity_years": [], "amount_maturing_12mo": [],
        "pct_maturing_12mo": [], "amount_maturing_90d": [], "pct_maturing_90d": [],
    }
    for d in dates:
        snap = df[(df["issue_date"] <= d) & (df["maturity_date"] > d)]
        total = float(snap["total_accepted"].sum())
        history["dates"].append(d.strftime("%Y-%m-%d"))
        history["total_outstanding"].append(round(total, 0))
        if total <= 0:
            history["weighted_avg_rate_pct"].append(None)
            history["annual_interest_cost"].append(0.0)
            history["weighted_avg_maturity_years"].append(None)
            history["amount_maturing_12mo"].append(0.0)
            history["pct_maturing_12mo"].append(None)
            history["amount_maturing_90d"].append(0.0)
            history["pct_maturing_90d"].append(None)
            continue

        days_to_mat = (snap["maturity_date"] - d).dt.days
        years_to_mat = days_to_mat / 365.25
        w_rate = weighted_avg(snap["total_accepted"], snap["effective_rate_pct"])
        annual_interest = float((snap["total_accepted"] * snap["cash_interest_rate_pct"]).sum() / 100)
        w_mat = float((snap["total_accepted"] * years_to_mat).sum() / total)
        amt_12mo = float(snap.loc[days_to_mat <= 365, "total_accepted"].sum())
        amt_90d = float(snap.loc[days_to_mat <= 90, "total_accepted"].sum())

        history["weighted_avg_rate_pct"].append(round(w_rate, 3) if w_rate is not None else None)
        history["annual_interest_cost"].append(round(annual_interest, 0))
        history["weighted_avg_maturity_years"].append(round(w_mat, 2))
        history["amount_maturing_12mo"].append(round(amt_12mo, 0))
        history["pct_maturing_12mo"].append(round(amt_12mo / total * 100, 2))
        history["amount_maturing_90d"].append(round(amt_90d, 0))
        history["pct_maturing_90d"].append(round(amt_90d / total * 100, 2))

    return history


def main():
    df = load_auctions()
    df["effective_rate_pct"] = compute_effective_rate(df)
    df["cash_interest_rate_pct"] = compute_cash_interest_rate(df, df["effective_rate_pct"])
    yield_curve = load_yield_curve()

    today_date = date.today()
    today = pd.Timestamp(today_date)
    # issue_date <= today matches historical_snapshots' definition of
    # "outstanding as of D" exactly, so the live tiles and the point-in-time
    # history line up (a security auctioned but not yet issued isn't counted
    # as outstanding debt yet either way).
    outstanding = df[(df["issue_date"] <= today) & (df["maturity_date"] > today)].copy()
    outstanding["days_to_maturity"] = (outstanding["maturity_date"] - today).dt.days
    outstanding["years_to_maturity"] = outstanding["days_to_maturity"] / 365.25
    outstanding["bucket"] = outstanding["days_to_maturity"].apply(bucket_for_days)

    total_outstanding = float(outstanding["total_accepted"].sum())

    # ---- headline summary ----
    weighted_avg_rate = weighted_avg(outstanding["total_accepted"], outstanding["effective_rate_pct"])
    weighted_avg_maturity = float(
        (outstanding["total_accepted"] * outstanding["years_to_maturity"]).sum() / total_outstanding
    )
    pct_maturing_12mo = float(
        outstanding.loc[outstanding["days_to_maturity"] <= 365, "total_accepted"].sum()
        / total_outstanding
        * 100
    )
    amount_maturing_90d = float(
        outstanding.loc[outstanding["days_to_maturity"] <= 90, "total_accepted"].sum()
    )
    pct_maturing_90d = amount_maturing_90d / total_outstanding * 100

    # ---- maturity wall: $ outstanding per bucket, split by type ----
    wall_pivot = outstanding.pivot_table(
        index="bucket", columns="type", values="total_accepted", aggfunc="sum", fill_value=0
    ).reindex(index=BUCKET_LABELS, columns=TYPE_ORDER, fill_value=0)
    by_type_amounts = {t: [float(v) for v in wall_pivot[t]] for t in TYPE_ORDER}

    # ---- weighted-average rate per bucket (separate chart -- see report_template.html for why not a second axis) ----
    outstanding["_rate_weight"] = outstanding["total_accepted"].where(outstanding["effective_rate_pct"].notna())
    outstanding["_rate_component"] = outstanding["total_accepted"] * outstanding["effective_rate_pct"]
    bucket_rates = (
        outstanding.groupby("bucket").agg(w=("_rate_weight", "sum"), wr=("_rate_component", "sum"))
        .reindex(BUCKET_LABELS)
    )
    rate_by_bucket = [
        round(float(wr / w), 3) if w else None
        for w, wr in zip(bucket_rates["w"], bucket_rates["wr"])
    ]

    # Annual interest cost implied by currently outstanding debt -- sum(face
    # value x rate) over each row, not total_outstanding x weighted_avg_rate
    # (that would silently misstate things if any outstanding row ever lacks
    # a rate). Uses cash_interest_rate_pct (coupon for Note/Bond/TIPS), NOT
    # effective_rate_pct (yield) -- this is a dollar "obligations" figure, so
    # it needs the actual cash payment, not the yield used for cost-of-
    # borrowing comparisons everywhere else. See compute_cash_interest_rate.
    outstanding["_cash_rate_component"] = outstanding["total_accepted"] * outstanding["cash_interest_rate_pct"]
    annual_interest_cost = float(outstanding["_cash_rate_component"].sum() / 100)

    # ---- debt composition over time: outstanding-by-type, one point-in-time
    # snapshot per calendar year (same issue_date<=D & maturity_date>D replay
    # used everywhere else in this app). The current year's snapshot is
    # "today" rather than Jan 1 of this year, so the most recent bar matches
    # the live numbers elsewhere on the page, without also needing a separate
    # "today" point that would otherwise duplicate the current year's label. ----
    composition_years = list(range(HISTORY_START.year, today_date.year))
    composition_dates = [pd.Timestamp(year=y, month=1, day=1) for y in composition_years] + [today]
    composition_periods = [str(y) for y in composition_years] + [str(today_date.year)]

    composition_by_type = {t: [] for t in TYPE_ORDER}
    for d in composition_dates:
        snap = df[(df["issue_date"] <= d) & (df["maturity_date"] > d)]
        totals = snap.groupby("type")["total_accepted"].sum()
        for t in TYPE_ORDER:
            composition_by_type[t].append(float(totals.get(t, 0.0)))

    composition_by_type_pct = {t: [] for t in TYPE_ORDER}
    for i in range(len(composition_periods)):
        year_total = sum(composition_by_type[t][i] for t in TYPE_ORDER)
        for t in TYPE_ORDER:
            pct = round(composition_by_type[t][i] / year_total * 100, 2) if year_total else 0.0
            composition_by_type_pct[t].append(pct)

    # ---- historical rate trend: weighted-avg rate at issuance, quarterly, full history ----
    hist = df[df["effective_rate_pct"].notna()].copy()
    hist["quarter"] = hist["auction_date"].dt.to_period("Q").astype(str)
    hist["_w"] = hist["total_accepted"]
    hist["_wr"] = hist["total_accepted"] * hist["effective_rate_pct"]
    q = hist.groupby("quarter").agg(w=("_w", "sum"), wr=("_wr", "sum")).sort_index()
    historical_periods = q.index.tolist()
    historical_rates = [round(float(wr / w), 3) for w, wr in zip(q["w"], q["wr"])]

    # ---- issuance mix: $ issued per year, by type, full history ----
    df["year"] = df["auction_date"].dt.year
    issuance_pivot = df.pivot_table(
        index="year", columns="type", values="total_accepted", aggfunc="sum", fill_value=0
    ).reindex(columns=TYPE_ORDER, fill_value=0)
    issuance_periods = [str(y) for y in issuance_pivot.index]
    issuance_by_type = {t: [float(v) for v in issuance_pivot[t]] for t in TYPE_ORDER}
    # Same breakdown as a % of that year's total issuance, for the chart's $/% toggle.
    issuance_pivot_pct = issuance_pivot.div(issuance_pivot.sum(axis=1), axis=0) * 100
    issuance_by_type_pct = {t: [round(float(v), 2) for v in issuance_pivot_pct[t]] for t in TYPE_ORDER}

    # ---- auction history by benchmark maturity: every individual auction
    # (raw, not aggregated) for each of the 7 durations, full history ----
    duration_series = {}
    for label, sec_type, sec_term in DURATION_DEFS:
        rows = df[(df["type"] == sec_type) & (df["term"] == sec_term)].sort_values("auction_date")
        duration_series[label] = [
            {
                "date": iso_date(row.auction_date),
                "rate": round(float(row.effective_rate_pct), 3) if pd.notna(row.effective_rate_pct) else None,
                "amount": float(row.total_accepted),
                "bid_to_cover": round(float(row.bid_to_cover_ratio), 2) if pd.notna(row.bid_to_cover_ratio) else None,
            }
            for row in rows.itertuples()
        ]

    # ---- refinancing cost impact: debt maturing in the next 12 months, its own
    # rate at issuance vs. the most recent auction of the same term ----
    # IMPORTANT: "term" alone isn't a unique key -- TIPS and FRN reuse the same
    # term labels as regular Notes/Bonds (e.g. both a 10-Year Note and a
    # 10-Year TIPS have term == "10-Year"), so grouping by term alone can match
    # a maturing nominal Note against a TIPS *real* yield (or vice versa) --
    # different instruments with rates that aren't comparable. Group by
    # (type, term) together so each security only ever gets compared to the
    # most recent auction of the exact same instrument.
    df["_type_term_key"] = df["type"] + " " + df["term"]
    latest_rate_by_type_term = (
        df.dropna(subset=["effective_rate_pct"])
        .sort_values("auction_date")
        .groupby("_type_term_key")["effective_rate_pct"]
        .last()
    )
    maturing_12mo = outstanding[outstanding["days_to_maturity"] <= 365].copy()
    maturing_12mo["current_term_rate"] = (maturing_12mo["type"] + " " + maturing_12mo["term"]).map(latest_rate_by_type_term)
    matched = maturing_12mo.dropna(subset=["effective_rate_pct", "current_term_rate"])
    matched_amount = float(matched["total_accepted"].sum())
    total_maturing_12mo = float(maturing_12mo["total_accepted"].sum())

    original_avg_rate = weighted_avg(matched["total_accepted"], matched["effective_rate_pct"])
    current_avg_rate = weighted_avg(matched["total_accepted"], matched["current_term_rate"])
    est_additional_annual_cost = (
        (current_avg_rate - original_avg_rate) / 100 * matched_amount
        if original_avg_rate is not None and current_avg_rate is not None
        else None
    )
    coverage_pct = round(matched_amount / total_maturing_12mo * 100, 1) if total_maturing_12mo else 0.0

    # ---- top individual maturities (reopenings of the same CUSIP combined) ----
    outstanding["_rate_weight2"] = outstanding["_rate_weight"]
    outstanding["_rate_component2"] = outstanding["_rate_component"]
    top_grp = outstanding.groupby(
        ["cusip", "type", "term", "security_term", "maturity_date"], as_index=False
    ).agg(
        amount=("total_accepted", "sum"),
        rate_weight=("_rate_weight2", "sum"),
        rate_component=("_rate_component2", "sum"),
    )
    top_grp["rate_pct"] = top_grp["rate_component"] / top_grp["rate_weight"]
    # Purely hypothetical "if this exact tranche were refinanced today at the
    # same term" rate -- the most recent auction of the same (type, term)
    # pair, same lookup used for the aggregate refinancing-cost-impact
    # section above (see the note there on why type is part of the key).
    top_grp["refi_rate_pct"] = (top_grp["type"] + " " + top_grp["term"]).map(latest_rate_by_type_term)
    top_grp = top_grp.sort_values("amount", ascending=False).head(20)
    top_maturities = [
        {
            "cusip": row.cusip,
            "type": row.type,
            "term": row.security_term,
            "maturity_date": iso_date(row.maturity_date),
            "amount": float(row.amount),
            "rate_pct": round(float(row.rate_pct), 3) if pd.notna(row.rate_pct) else None,
            "refi_rate_pct": round(float(row.refi_rate_pct), 3) if pd.notna(row.refi_rate_pct) else None,
        }
        for row in top_grp.itertuples()
    ]

    # ---- recent auctions: every individual auction event in the last 2
    # weeks (unlike top_maturities, reopenings are NOT combined here -- this
    # is a log of auction events, not a snapshot of outstanding debt) ----
    recent_cutoff = today - pd.Timedelta(days=14)
    recent = df[df["auction_date"] >= recent_cutoff].sort_values("auction_date", ascending=False)
    recent_auctions = [
        {
            "cusip": row.cusip,
            "type": row.type,
            "term": row.security_term,
            "auction_date": iso_date(row.auction_date),
            "maturity_date": iso_date(row.maturity_date),
            "amount": float(row.total_accepted),
            "rate_pct": round(float(row.effective_rate_pct), 3) if pd.notna(row.effective_rate_pct) else None,
            "bid_to_cover": round(float(row.bid_to_cover_ratio), 2) if pd.notna(row.bid_to_cover_ratio) else None,
        }
        for row in recent.itertuples()
    ]

    # ---- point-in-time history for the "Show history" tile sparklines ----
    hist_dates = month_start_dates(HISTORY_START, today_date.replace(day=1))
    if hist_dates[-1] != today:
        hist_dates.append(today)
    history = historical_snapshots(df, hist_dates)

    data = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "data_through": {
            "latest_auction_date": iso_date(df["auction_date"].max()),
            "latest_issue_date": iso_date(df["issue_date"].max()),
            "total_records": int(len(df)),
        },
        "summary": {
            "total_outstanding": total_outstanding,
            "weighted_avg_rate_pct": round(weighted_avg_rate, 3) if weighted_avg_rate is not None else None,
            "annual_interest_cost": annual_interest_cost,
            "weighted_avg_maturity_years": round(weighted_avg_maturity, 2),
            "amount_maturing_12mo": total_maturing_12mo,
            "pct_maturing_12mo": round(pct_maturing_12mo, 2),
            "amount_maturing_90d": amount_maturing_90d,
            "pct_maturing_90d": round(pct_maturing_90d, 2),
        },
        "history": history,
        "maturity_wall": {
            "buckets": BUCKET_LABELS,
            "by_type": by_type_amounts,
        },
        "rate_by_bucket": {
            "buckets": BUCKET_LABELS,
            "avg_rate_pct": rate_by_bucket,
        },
        "composition": {
            "periods": composition_periods,
            "by_type": composition_by_type,
            "by_type_pct": composition_by_type_pct,
        },
        "historical_rate_trend": {
            "periods": historical_periods,
            "avg_rate_pct": historical_rates,
        },
        "issuance_mix": {
            "periods": issuance_periods,
            "by_type": issuance_by_type,
            "by_type_pct": issuance_by_type_pct,
        },
        "duration_detail": {
            "durations": [label for label, _, _ in DURATION_DEFS],
            "series": duration_series,
        },
        "yield_curve": yield_curve,
        "refinancing_gap": {
            "amount_maturing_12mo": total_maturing_12mo,
            "coverage_pct": coverage_pct,
            "original_avg_rate_pct": round(original_avg_rate, 3) if original_avg_rate is not None else None,
            "current_avg_rate_pct": round(current_avg_rate, 3) if current_avg_rate is not None else None,
            "est_additional_annual_cost": (
                round(est_additional_annual_cost, 0) if est_additional_annual_cost is not None else None
            ),
        },
        "top_maturities": top_maturities,
        "recent_auctions": recent_auctions,
        "type_order": TYPE_ORDER,
        "caveats": [
            "Covers marketable, auctioned Treasury debt only — excludes savings bonds, SLGS, and "
            "intragovernmental holdings, so totals will not match the “total national debt” headline figure.",
            "Treasury's debt buyback program (started 2024) is not modeled; a small amount of repurchased "
            "debt may still appear as outstanding here.",
            "Rates are shown on a comparable annualized basis: bond-equivalent yield for Bills/CMBs, auction "
            "yield for Notes/Bonds/TIPS — except the “$X/yr in interest” figure, which uses each Note/Bond/"
            "TIPS's literal coupon rate instead of its yield, since it represents actual cash interest "
            "obligations rather than a cost-of-borrowing comparison.",
            "FRN rates are approximated as that auction's fixed spread plus the most recent 13-week Bill rate "
            "at the time — a proxy for a rate that actually floats weekly.",
            "“Refinancing cost impact” compares maturing debt's own rate at issuance to the most recent "
            "auction of the same term — a simplified proxy, not an official Treasury estimate. Covers "
            f"{coverage_pct:.1f}% of debt maturing in the next 12 months (the rest has no directly comparable "
            "recent auction of the same term).",
            "The “Refi. rate” column on the maturities table is the same kind of hypothetical, applied per "
            "tranche: what a new auction of that same term is yielding today, not a prediction of what will "
            "actually happen when that security is refinanced.",
            "The “Show history” sparklines replay each metric using only securities issued by that date. "
            "History before roughly 2010 understates totals slightly, since debt issued before 1980 (when "
            "this dataset starts) that was still outstanding in the 1980s–2000s isn't captured.",
        ],
    }

    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")
    template = TEMPLATE_PATH.read_text(encoding="utf-8")
    if DATA_PLACEHOLDER not in template:
        raise RuntimeError(f"Template is missing the {DATA_PLACEHOLDER} placeholder; can't inject data.")

    output = template.replace(DATA_PLACEHOLDER, json.dumps(data))
    OUTPUT_PATH.write_text(output, encoding="utf-8")
    print(
        f"Wrote {OUTPUT_PATH.name}: {len(outstanding):,} outstanding securities, "
        f"${total_outstanding:,.0f} total, as of {date.today()}."
    )


if __name__ == "__main__":
    main()
