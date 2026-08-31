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


def weighted_avg(amounts, rates):
    weight = amounts.where(rates.notna())
    total = weight.sum()
    if not total or pd.isna(total):
        return None
    return float((weight * rates).sum() / total)


def iso_date(value):
    return None if pd.isna(value) else value.strftime("%Y-%m-%d")


def main():
    df = load_auctions()
    df["effective_rate_pct"] = compute_effective_rate(df)

    today = pd.Timestamp(date.today())
    outstanding = df[df["maturity_date"] > today].copy()
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

    # ---- composition: total $ outstanding by type ----
    composition_amounts = (
        outstanding.groupby("type")["total_accepted"].sum().reindex(TYPE_ORDER, fill_value=0)
    )
    composition_amounts = [float(v) for v in composition_amounts]

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

    # ---- refinancing cost impact: debt maturing in the next 12 months, its own
    # rate at issuance vs. the most recent auction of the same term ----
    latest_rate_by_term = (
        df.dropna(subset=["effective_rate_pct"])
        .sort_values("auction_date")
        .groupby("term")["effective_rate_pct"]
        .last()
    )
    maturing_12mo = outstanding[outstanding["days_to_maturity"] <= 365].copy()
    maturing_12mo["current_term_rate"] = maturing_12mo["term"].map(latest_rate_by_term)
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
    top_grp = top_grp.sort_values("amount", ascending=False).head(20)
    top_maturities = [
        {
            "cusip": row.cusip,
            "type": row.type,
            "term": row.security_term,
            "maturity_date": iso_date(row.maturity_date),
            "amount": float(row.amount),
            "rate_pct": round(float(row.rate_pct), 3) if pd.notna(row.rate_pct) else None,
        }
        for row in top_grp.itertuples()
    ]

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
            "weighted_avg_maturity_years": round(weighted_avg_maturity, 2),
            "pct_maturing_12mo": round(pct_maturing_12mo, 2),
            "amount_maturing_90d": amount_maturing_90d,
        },
        "maturity_wall": {
            "buckets": BUCKET_LABELS,
            "by_type": by_type_amounts,
        },
        "rate_by_bucket": {
            "buckets": BUCKET_LABELS,
            "avg_rate_pct": rate_by_bucket,
        },
        "composition": {
            "types": TYPE_ORDER,
            "amounts": composition_amounts,
        },
        "historical_rate_trend": {
            "periods": historical_periods,
            "avg_rate_pct": historical_rates,
        },
        "issuance_mix": {
            "periods": issuance_periods,
            "by_type": issuance_by_type,
        },
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
        "type_order": TYPE_ORDER,
        "caveats": [
            "Covers marketable, auctioned Treasury debt only — excludes savings bonds, SLGS, and "
            "intragovernmental holdings, so totals will not match the “total national debt” headline figure.",
            "Treasury's debt buyback program (started 2024) is not modeled; a small amount of repurchased "
            "debt may still appear as outstanding here.",
            "Rates are shown on a comparable annualized basis: bond-equivalent yield for Bills/CMBs, auction "
            "yield for Notes/Bonds/TIPS.",
            "FRN rates are approximated as that auction's fixed spread plus the most recent 13-week Bill rate "
            "at the time — a proxy for a rate that actually floats weekly.",
            "The 4-week Bill auctioned 2001-09-11 has no rate on file in Treasury's own historical records.",
            "“Refinancing cost impact” compares maturing debt's own rate at issuance to the most recent "
            "auction of the same term — a simplified proxy, not an official Treasury estimate. Covers "
            f"{coverage_pct:.1f}% of debt maturing in the next 12 months (the rest has no directly comparable "
            "recent auction of the same term).",
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
