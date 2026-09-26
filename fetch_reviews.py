"""Collect public App Store reviews of Indonesian e-wallets from Apple's official review feed.

Source: https://itunes.apple.com/id/rss/customerreviews/page=N/id=APP_ID/sortby=SORT/json
Apple serves up to 10 pages x 50 reviews per sort order. We read both "mostrecent" and
"mosthelpful" and de-duplicate by review id.

Privacy: reviewer names are NOT stored. Phone numbers, e-mail addresses, long digit strings
(account / transaction ids) and @handles in the review text are masked before saving.
"""
import json
import re
import time
from pathlib import Path

import pandas as pd
import requests

HERE = Path(__file__).resolve().parent
RAW = HERE / "raw" / "reviews"
RAW.mkdir(parents=True, exist_ok=True)
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; personal research project)"}

APPS = {  # App Store ids, looked up with Apple's public iTunes Search API
    "DANA": 1437123008,
    "OVO": 1142114207,
    "GoPay": 6446321594,
    "ShopeePay": 6455990519,
    "LinkAja": 984436451,
}
SORTS = ["mostrecent", "mosthelpful"]

PII_PATTERNS = [
    (re.compile(r"[\w.+-]+@[\w-]+\.[\w.]+"), "[EMAIL]"),
    (re.compile(r"(?:\+?62|0)8[\d\s-]{7,13}\d"), "[PHONE]"),
    (re.compile(r"\b\d{9,}\b"), "[NUMBER]"),
    (re.compile(r"@\w{3,}"), "[HANDLE]"),
]


def mask(text):
    for pat, repl in PII_PATTERNS:
        text = pat.sub(repl, text)
    return text


def fetch(app_id, sort, page):
    cache = RAW / f"{app_id}_{sort}_{page}.json"
    if cache.exists():
        return json.loads(cache.read_text(encoding="utf-8"))
    url = f"https://itunes.apple.com/id/rss/customerreviews/page={page}/id={app_id}/sortby={sort}/json"
    for attempt in range(4):
        try:
            r = requests.get(url, headers=HEADERS, timeout=60)
            break
        except requests.RequestException:
            time.sleep(5 * (attempt + 1))
    else:
        raise RuntimeError(f"feed unreachable after retries: {url}")
    r.raise_for_status()
    data = r.json()
    # cache only the fields we use, and never the author block
    entries = data.get("feed", {}).get("entry", [])
    if isinstance(entries, dict):
        entries = [entries]
    slim = [{"id": e["id"]["label"], "updated": e["updated"]["label"], "rating": e["im:rating"]["label"],
             "version": e["im:version"]["label"], "title": mask(e["title"]["label"]),
             "text": mask(e["content"]["label"]), "votes": e.get("im:voteSum", {}).get("label", "0")}
            for e in entries if "im:rating" in e]
    cache.write_text(json.dumps(slim, ensure_ascii=False), encoding="utf-8")
    time.sleep(1.0)
    return slim


def main():
    rows = []
    for app, app_id in APPS.items():
        for sort in SORTS:
            for page in range(1, 11):
                got = fetch(app_id, sort, page)
                if not got:
                    break
                for e in got:
                    e.update(app=app, sort=sort)
                rows += got
        print(f"{app}: {sum(r['app'] == app for r in rows)} rows so far")
    df = pd.DataFrame(rows)
    df["rating"] = df["rating"].astype(int)
    df["votes"] = pd.to_numeric(df["votes"], errors="coerce").fillna(0).astype(int)
    df["date"] = pd.to_datetime(df["updated"], utc=True).dt.tz_convert("Asia/Jakarta").dt.date
    # A review can appear in both feeds; keep one row with a flag per feed.
    flags = df.pivot_table(index="id", columns="sort", values="rating", aggfunc="size", fill_value=0) > 0
    df = df.drop(columns=["updated", "sort"]).drop_duplicates("id").set_index("id")
    df["in_recent"] = flags.reindex(df.index)["mostrecent"].fillna(False).astype(bool)
    df["in_helpful"] = flags.reindex(df.index)["mosthelpful"].fillna(False).astype(bool)
    df = df.reset_index()
    df["review"] = (df["title"].fillna("") + ". " + df["text"].fillna("")).str.strip()
    out = df[["id", "app", "date", "rating", "version", "votes", "in_recent", "in_helpful", "title", "text", "review"]]
    out.to_csv(HERE / "reviews.csv", index=False)
    print(f"reviews.csv: {len(out)} unique reviews")
    print(out.groupby("app").agg(n=("id", "count"), avg_rating=("rating", "mean"),
                                 first=("date", "min"), last=("date", "max")).round(2).to_string())


if __name__ == "__main__":
    main()
