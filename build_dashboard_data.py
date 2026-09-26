"""Build the two flat files the Tableau workbook reads.

dashboard_reviews.csv  one row per review x theme x lens. A review with three themes gives three
                       rows, and a review with no theme gives one "(no complaint theme)" row.
                       - Share Weight = 1 / reviews of that app in that lens, so
                         SUM(Share Weight) for one theme and one app = share of that app's reviews.
                       - First Row = 1 on one row per review and lens, for review-level measures
                         such as average rating.
dashboard_market.csv   Bank Indonesia e-money statistics (annual and monthly) in long format.
"""
from pathlib import Path

import pandas as pd

HERE = Path(__file__).resolve().parent
SHORT = {
    "Failed / pending transfer or top-up": "Failed transfer",
    "Customer service": "Customer service",
    "Login, OTP & verification": "Login, OTP & KYC",
    "App performance & bugs": "App bugs & speed",
    "Balance lost / unauthorised deduction": "Balance lost",
    "Paylater, loans & debt collection": "Paylater & bills",
    "Fees & charges": "Fees & charges",
    "Hard to use": "Hard to use",
    "Account security & fraud": "Security & fraud",
    "Ads, notifications & promos": "Ads & promos",
}
NO_THEME = "(no complaint theme)"
METRICS = {"instruments": "E-money instruments (million)",
           "transaction_volume": "Transactions (million)",
           "transaction_value": "Transaction value (Rp trillion)"}


def snippet(text, n=220):
    t = " ".join(str(text).split())
    return t if len(t) <= n else t[:n].rsplit(" ", 1)[0] + " ..."


def reviews():
    df = pd.read_csv(HERE / "reviews_tagged.csv", parse_dates=["date"])
    df["theme"] = df["themes_str"].fillna("").map(lambda s: [SHORT[t] for t in s.split("; ") if t] or [NO_THEME])
    parts = []
    for flag, lens in (("in_recent", "Latest 500"), ("in_helpful", "Most helpful")):
        sub = df[df[flag]].copy()
        sub["lens"] = lens
        sub["share_weight"] = 1 / sub.groupby("app").id.transform("size")
        parts.append(sub)
    out = pd.concat(parts).explode("theme")
    out["first_row"] = (~out.duplicated(["id", "lens"])).astype(int)
    out["one_star"] = (out.rating == 1).astype(int)
    out["negative"] = out.is_negative.astype(int)
    out["peer_group"] = out.app.where(out.app == "DANA", "Peers")
    out["snippet"] = out.review.map(snippet)
    cols = {"id": "Review ID", "app": "App", "peer_group": "Peer Group", "lens": "Lens", "date": "Date",
            "rating": "Rating", "version": "Version", "theme": "Theme", "n_themes": "Theme Count",
            "share_weight": "Share Weight", "first_row": "First Row", "one_star": "One Star",
            "negative": "Negative", "votes": "Helpful Votes", "snippet": "Snippet"}
    out = out[list(cols)].rename(columns=cols)
    out.to_csv(HERE / "dashboard_reviews.csv", index=False, encoding="utf-8")
    check = out[(out.Lens == "Latest 500") & (out.Theme != NO_THEME)].groupby(["App", "Theme"])["Share Weight"].sum()
    print(f"dashboard_reviews.csv: {len(out)} rows | DANA paylater share = {check[('DANA', 'Paylater & bills')]:.3f}")


def market():
    bi = pd.read_csv(HERE / "bi_emoney.csv")
    bi["Period"] = pd.to_datetime(dict(year=bi.year, month=bi.month.fillna(1).astype(int), day=1))
    scale = {"instruments": 1, "transaction_volume": 1e-3, "transaction_value": 1e-3}
    bi["Value"] = bi.value * bi.metric.map(scale)
    bi["Metric"] = bi.metric.map(METRICS)
    bi["Period Type"] = bi.period.map({"year": "Annual", "month": "Monthly"})
    bi[["Metric", "Period Type", "Period", "Value"]].to_csv(HERE / "dashboard_market.csv", index=False)
    print(f"dashboard_market.csv: {len(bi)} rows")


if __name__ == "__main__":
    reviews()
    market()
