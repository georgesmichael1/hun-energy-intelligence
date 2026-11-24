"""MNB (Magyar Nemzeti Bank) official daily exchange rates: EUR/HUF and USD/HUF.

The HUF floats; there is no peg. The English page prints the publication date
("Napi árfolyamok: 24 November 2026") and one row per currency
(code, name, unit, value). Strict checks: the date must parse, and each rate
must sit inside a wide sanity band. A failed day never erases saved rates.
"""
import csv
import datetime as dt
import io
import math
import re
from pathlib import Path

import requests
from pipeline_status import atomic_text, record, utc_now

SOURCE = "https://www.mnb.hu/en/arfolyamok/"
COLUMNS = ["Rate_Date", "Currency", "Unit", "HUF_Per_Unit", "Source", "Retrieved_UTC"]
WATCHED = ("EUR", "USD")
HISTORY_DAYS = 62
MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]

HEADERS = {"User-Agent": (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")}


def parse_page(html, today):
    m = re.search(r"Napi árfolyamok: (\d{1,2}) (\w+) (\d{4})", html)
    if not m or m.group(2) not in MONTHS:
        raise ValueError("MNB publication date missing or unparseable")
    rate_date = dt.date(int(m.group(3)), MONTHS.index(m.group(2)) + 1, int(m.group(1)))
    if rate_date > today or (today - rate_date).days > 10:
        raise ValueError("MNB publication date implausible")
    rows = []
    for currency in WATCHED:
        row = re.search(
            r'<td class="fw-b">\s*' + currency + r'\s*</td>(.*?)</tr>', html, re.S)
        if not row:
            raise ValueError("MNB row missing for " + currency)
        cells = re.findall(r"<td[^>]*>(.*?)</td>", row.group(1), re.S)
        cells = [re.sub(r"<[^>]+>", "", c).strip() for c in cells]
        if len(cells) < 3:
            raise ValueError("MNB row shape changed for " + currency)
        unit = int(cells[1].replace(" ", ""))
        value = float(cells[2].replace(" ", "").replace(",", "."))
        if unit < 1 or unit > 10000 or not math.isfinite(value) or not 0.0001 < value < 100000:
            raise ValueError("MNB value outside sanity band for " + currency)
        rows.append({"Rate_Date": rate_date.isoformat(), "Currency": currency,
                     "Unit": unit, "HUF_Per_Unit": round(value, 4)})
    return rows


def update_mnb_fx(path="mnb_fx.csv", session=None, now=None):
    now = now or dt.datetime.now(dt.timezone.utc)
    stamp = utc_now()
    own = session is None
    session = session or requests.Session()
    try:
        response = session.get(SOURCE, headers=HEADERS, timeout=(10, 30))
        if response.status_code != 200:
            raise ValueError("HTTP_" + str(response.status_code))
        rows = parse_page(response.text, now.date())
        for r in rows:
            r["Source"] = SOURCE
            r["Retrieved_UTC"] = stamp
        old = []
        p = Path(path)
        if p.exists():
            with p.open(newline="", encoding="utf-8") as handle:
                old = list(csv.DictReader(handle))
        merged = {(r["Rate_Date"], r["Currency"]): r for r in old}
        for r in rows:
            prior = merged.get((r["Rate_Date"], r["Currency"]))
            if prior and prior["HUF_Per_Unit"] != str(r["HUF_Per_Unit"]):
                raise ValueError("MNB restated a published rate; keeping both out until reviewed")
            merged[(r["Rate_Date"], r["Currency"])] = r
        cutoff = (now.date() - dt.timedelta(days=HISTORY_DAYS)).isoformat()
        kept = [merged[k] for k in sorted(merged) if k[0] >= cutoff]
        out = io.StringIO(newline="")
        writer = csv.DictWriter(out, fieldnames=COLUMNS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(kept)
        atomic_text(p, out.getvalue())
        eur = next(r["HUF_Per_Unit"] for r in rows if r["Currency"] == "EUR")
        usd = next(r["HUF_Per_Unit"] for r in rows if r["Currency"] == "USD")
        record("fx", "OK", detail="Official MNB daily rates read with publication date; HUF floats, no peg",
               source=SOURCE, source_date=rows[0]["Rate_Date"], eur_huf=eur, usd_huf=usd)
        return rows
    except Exception as exc:
        detail = str(exc) if isinstance(exc, ValueError) else type(exc).__name__
        record("fx", "FAILED", detail="Saved dated rates retained; " + detail, source=SOURCE)
        return None
    finally:
        if own:
            session.close()


if __name__ == "__main__":
    update_mnb_fx()
