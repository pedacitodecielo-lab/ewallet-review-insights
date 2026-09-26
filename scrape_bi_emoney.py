"""Scrape Indonesia's e-money market statistics from Bank Indonesia's payment-system page.

Steps
1. Load the public statistics page and discover the links for "Tabel 5e. Uang Elektronik".
2. Try the Excel link first. On 2026-09-26 that link returned an HTML 404 page, so the
   script checks what it actually received and falls back to the PDF version.
3. Extract the rows we need from the PDF text (number of instruments, transaction volume,
   transaction value) for the annual columns and the monthly columns.

Politeness: two page requests plus one file download, cached in raw/; stops on 403/429/503.
"""
import re
import time
from pathlib import Path

import pandas as pd
import pdfplumber
import requests

HERE = Path(__file__).resolve().parent
RAW = HERE / "raw"
RAW.mkdir(exist_ok=True)
BASE = "https://www.bi.go.id"
PAGE = BASE + "/id/statistik/ekonomi-keuangan/ssp/default.aspx"
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; personal research project)"}

# row number in Tabel 5e -> (metric, unit)
ROWS = {1: ("instruments", "million units"),
        10: ("transaction_volume", "thousand transactions"),
        20: ("transaction_value", "billion rupiah")}
MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6, "jul": 7, "aug": 8,
          "sep": 9, "oct": 10, "nov": 11, "dec": 12}


def get(url):
    r = requests.get(url, headers=HEADERS, timeout=60)
    if r.status_code in (403, 429, 503):
        raise SystemExit(f"HTTP {r.status_code} from {url}: server asks us to stop")
    r.raise_for_status()
    time.sleep(1.5)
    return r


def discover_links():
    html = get(PAGE).text
    links = sorted(set(re.findall(r'href="([^"]*TABEL_5e\.(?:xlsx?|pdf))"', html, re.I)))
    print("found:", links)
    return {l.rsplit(".", 1)[1].lower(): BASE + l for l in links}


def download(links):
    """Prefer Excel; accept a file only if it really is that file type."""
    for kind in ("xls", "xlsx", "pdf"):
        if kind not in links:
            continue
        path = RAW / f"TABEL_5e.{kind}"
        if not path.exists():
            content = get(links[kind]).content
            is_pdf = content[:5] == b"%PDF-"
            is_xls = content[:4] in (b"\xd0\xcf\x11\xe0", b"PK\x03\x04")
            if (kind == "pdf" and not is_pdf) or (kind != "pdf" and not is_xls):
                print(f"{kind}: server returned something else (starts {content[:15]!r}); skipping")
                continue
            path.write_bytes(content)
        print("using", path.name)
        return path
    raise SystemExit("no usable Tabel 5e file")


def numbers(text):
    # 2,719,803 / 1,065.44 / 0 ; thousands use commas and decimals use dots in this table
    return [float(x.replace(",", "")) for x in re.findall(r"\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+\.\d+|\d+", text)]


def parse_pdf(path):
    with pdfplumber.open(path) as pdf:
        p1 = pdf.pages[0].extract_text().splitlines()
        p2 = pdf.pages[1].extract_text().splitlines()
    # Page 1: header lists annual years then month abbreviations; rows start with the row number.
    head = next(l for l in p1 if l.startswith("KOMPONEN"))
    years = [int(y) for y in re.findall(r"\b20\d\d\b", head)]
    month_line = next(l for l in p1 if re.fullmatch(r"(?:[A-Z][a-z]{2}\s*)+", l.strip()))
    p1_months = month_line.split()
    # Page 2: a month header line, and each row ends with its label and row number.
    p2_month_line = next(l for l in p2 if re.fullmatch(r"(?:[A-Z][a-z]{2}\s*)+", l.strip()))
    p2_months = p2_month_line.split()
    year_line = next(l for l in p2 if re.fullmatch(r"(?:20\d\d\s*)+", l.strip()))
    p2_years = [int(y) for y in year_line.split()]

    # assign calendar years to month columns: page 1 months belong to the last annual year,
    # page 2 months continue in order and roll into the next year after December
    last_year = years[-1]
    seq = [(last_year, MONTHS[m.lower()]) for m in p1_months]
    y = last_year if p2_years[0] == last_year else p2_years[0]
    prev = seq[-1][1]
    for m in p2_months:
        mm = MONTHS[m.lower()]
        if mm < prev:
            y += 1
        seq.append((y, mm))
        prev = mm

    records = []
    for rownum, (metric, unit) in ROWS.items():
        l1 = next(l for l in p1 if re.match(rf"^{rownum}\s", l))
        vals = numbers(re.sub(r"^\d+\s+[^\d]+", "", l1))  # drop row number and label text
        l2 = next(l for l in p2 if re.search(rf"\s{rownum}$", l.strip()))
        vals2 = numbers(re.sub(r"[A-Za-z(),/ ]+\d+$", "", l2.strip()))
        annual, monthly = vals[:len(years)], vals[len(years):] + vals2
        assert len(monthly) == len(seq), (metric, len(monthly), len(seq))
        for yr, v in zip(years, annual):
            records.append(dict(metric=metric, unit=unit, period="year", year=yr, month=None, value=v))
        for (yr, mm), v in zip(seq, monthly):
            records.append(dict(metric=metric, unit=unit, period="month", year=yr, month=mm, value=v))
    return pd.DataFrame(records)


def main():
    path = download(discover_links())
    df = parse_pdf(path) if path.suffix == ".pdf" else pd.read_excel(path)
    # sanity checks: the December monthly values must sum/match the annual figures
    ann = df[df.period == "year"].set_index(["metric", "year"])["value"]
    mon = df[df.period == "month"]
    for metric in ("transaction_volume", "transaction_value"):
        yr = int(mon.year.min())
        full = mon[(mon.metric == metric) & (mon.year == yr)]
        if len(full) == 12:
            assert abs(full.value.sum() / ann[(metric, yr)] - 1) < 0.01, metric
    dec = mon[(mon.metric == "instruments") & (mon.month == 12)]
    for _, r in dec.iterrows():
        if (("instruments", r.year) in ann.index):
            assert abs(r.value / ann[("instruments", r.year)] - 1) < 0.001, "Dec stock must equal annual"
    df.to_csv(HERE / "bi_emoney.csv", index=False)
    print(df[df.period == "year"].pivot(index="year", columns="metric", values="value").to_string())
    print(f"monthly rows: {len(mon)} ({mon.year.min()}-{int(mon[mon.year == mon.year.min()].month.min()):02d} "
          f"to {mon.year.max()}-{int(mon[mon.year == mon.year.max()].month.max()):02d})")


if __name__ == "__main__":
    main()
