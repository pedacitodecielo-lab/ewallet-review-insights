"""Headline analysis for the memo: market context, app scorecard, complaint priority, DANA vs peers.

Design choice: Apple's feed returns the latest 500 reviews per app, and those 500 reviews cover
25 days for DANA but more than two years for LinkAja. So apps are compared as equal-size
snapshots ("of the last 500 reviews, what share mention X"), never as a common time trend.
Review velocity (reviews per day) is reported separately as an engagement signal.
"""
from pathlib import Path

import pandas as pd
from scipy.stats import chi2_contingency

HERE = Path(__file__).resolve().parent
OUT = HERE / "outputs"
OUT.mkdir(exist_ok=True)
MIN_VERSION_REVIEWS = 30


def load():
    df = pd.read_csv(HERE / "reviews_tagged.csv", parse_dates=["date"])
    df["themes"] = df["themes_str"].fillna("").map(lambda s: [t for t in s.split("; ") if t])
    df["one_star"] = df["rating"] == 1
    return df


def market():
    bi = pd.read_csv(HERE / "bi_emoney.csv")
    yr = bi[bi.period == "year"].pivot(index="year", columns="metric", values="value")
    first, last = yr.index.min(), yr.index.max()
    growth = (yr.loc[last] / yr.loc[first]).rename("multiple")
    cagr = (growth ** (1 / (last - first)) - 1).rename("cagr")
    yr["avg_ticket_rp"] = yr.transaction_value * 1e9 / (yr.transaction_volume * 1e3)
    print(f"\nMarket (Bank Indonesia, {first}-{last})")
    print(yr.round(0).to_string())
    print(pd.concat([growth, cagr], axis=1).round(3).to_string())
    return yr


def scorecard(df):
    rec = df[df.in_recent]
    g = rec.groupby("app").agg(reviews=("id", "size"), avg_rating=("rating", "mean"),
                               one_star=("one_star", "mean"), negative=("is_negative", "mean"),
                               first=("date", "min"), last=("date", "max"))
    g["days_covered"] = (g["last"] - g["first"]).dt.days + 1
    g["reviews_per_day"] = g.reviews / g.days_covered
    g["helpful_avg_rating"] = df[df.in_helpful].groupby("app").rating.mean()
    g = g.sort_values("avg_rating", ascending=False)
    print("\nApp scorecard (latest 500 reviews per app)")
    print(g.round(2).to_string())
    g.round(4).to_csv(OUT / "app_scorecard.csv")
    return g


def theme_table(df, lens="in_recent"):
    sub = df[df[lens]].explode("themes").dropna(subset=["themes"])
    n = df[df[lens]].groupby("app").size()
    t = sub.groupby(["app", "themes"]).agg(mentions=("id", "size"), one_star=("one_star", "mean"),
                                           avg_rating=("rating", "mean")).reset_index()
    t["share_of_reviews"] = t.mentions / t.app.map(n)
    return t.rename(columns={"themes": "theme"})


def priority(df):
    t = theme_table(df)
    # market view: each app weighted equally (equal-size snapshots)
    m = t.groupby("theme").agg(frequency=("share_of_reviews", "mean"), mentions=("mentions", "sum")).reset_index()
    sev = df[df.in_recent].explode("themes").groupby("themes").one_star.mean()
    m["severity_one_star"] = m.theme.map(sev)
    fx, sx = m.frequency.median(), m.severity_one_star.median()
    m["quadrant"] = [("Fix first" if f >= fx and s >= sx else "Frequent, milder" if f >= fx
                      else "Rare but severe" if s >= sx else "Monitor") for f, s in zip(m.frequency, m.severity_one_star)]
    m = m.sort_values("frequency", ascending=False)
    print("\nPriority matrix (market, equal-weighted across 5 apps)")
    print(m.round(3).to_string(index=False))
    m.round(4).to_csv(OUT / "priority_matrix.csv", index=False)
    return t, m


def dana_vs_peers(t):
    wide = t.pivot(index="theme", columns="app", values="share_of_reviews").fillna(0)
    peers = [c for c in wide.columns if c != "DANA"]
    out = pd.DataFrame({"DANA": wide["DANA"], "peer_avg": wide[peers].mean(axis=1),
                        "best_peer": wide[peers].min(axis=1), "best_peer_app": wide[peers].idxmin(axis=1)})
    out["gap_pp"] = (out.DANA - out.peer_avg) * 100
    out = out.sort_values("gap_pp", ascending=False)
    print("\nDANA vs peers: share of latest 500 reviews mentioning each theme")
    print(out.round(3).to_string())
    out.round(4).to_csv(OUT / "dana_vs_peers.csv")
    return out


def dana_versions(df):
    d = df[(df.app == "DANA") & df.in_recent]
    v = d.groupby("version").agg(reviews=("id", "size"), avg_rating=("rating", "mean"),
                                 one_star=("one_star", "mean"), first_seen=("date", "min")).query(
        "reviews >= @MIN_VERSION_REVIEWS").sort_values("first_seen")
    ex = d.explode("themes")
    share = (ex.groupby(["version", "themes"]).id.nunique() / d.groupby("version").size()).unstack().reindex(v.index)
    base = v.index[:-1]
    change = (share.loc[v.index[-1]] - share.loc[base].mean()).sort_values(ascending=False) * 100
    print("\nDANA by app version (latest-500 snapshot, versions with >= 30 reviews)")
    print(v.round(2).to_string())
    # is the latest version's 1-star share really higher? chi-square on a 2x2 table
    latest = d[d.version.isin(v.index)]
    table = pd.crosstab(latest.version == v.index[-1], latest.one_star)
    print(f"1-star share, {v.index[-1]} vs earlier versions: chi-square p = {chi2_contingency(table)[1]:.3f}")
    print(f"theme share change, {v.index[-1]} vs earlier versions (pp):")
    print(change.round(1).head(5).to_string())
    v.round(4).to_csv(OUT / "dana_versions.csv")
    return v, change


def lens_check(df):
    """Do the 'most helpful' reviews (voted up by other users) stress different themes?"""
    a = theme_table(df, "in_recent").groupby("theme").share_of_reviews.mean()
    b = theme_table(df, "in_helpful").groupby("theme").share_of_reviews.mean()
    out = pd.DataFrame({"recent": a, "helpful": b}).fillna(0)
    out["diff_pp"] = (out.helpful - out.recent) * 100
    print("\nLens check: most-helpful vs most-recent (equal-weighted share)")
    print(out.sort_values("diff_pp", ascending=False).round(3).to_string())
    return out


def main():
    df = load()
    market()
    scorecard(df)
    t, _ = priority(df)
    dana_vs_peers(t)
    dana_versions(df)
    lens_check(df)
    t.round(4).to_csv(OUT / "theme_by_app.csv", index=False)


if __name__ == "__main__":
    main()
