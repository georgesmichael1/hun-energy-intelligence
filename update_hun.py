"""Hungary energy intelligence master updater (token-free sources).

Runs every 30 minutes from GitHub Actions. Fetches MAVIR operational feeds,
HUPX-zone day-ahead prices, MNB FX, IMF GDP, World Bank macro, Eurostat
HICP/tariffs/population and the MAVIR balance responsible party register, then
syncs the live rows of hun_data.csv and writes status.json.

ENTSO-E token feeds (load forecast, physical flows, official A44 prices) run in
their own daily workflows, exactly as in the Bosnia pipeline.

No dataset failure is ever hidden: each is recorded in status.json and the
workflow gate turns the run red. Saved dated data is always retained on
failure. Nothing here writes typed-in fallback values.
"""
import csv
import datetime as dt
import io
import math
import os
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from pipeline_status import DATASETS, atomic_text, record, utc_now, write_status
import mavir_adapter
import hupx_prices
import mnb_fx
import imf_gdp
import eurostat_feeds
import brp_feed

TZ = ZoneInfo("Europe/Budapest")
UTC = dt.timezone.utc
MASTER_CSV = "hun_data.csv"
STATUS_JSON = "status.json"

EUROSTAT_POP_URL = ("https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"
                    "demo_pjan?geo=HU&sex=T&age=TOTAL&unit=NR&lang=EN&lastTimePeriod=6")


def make_session():
    session = requests.Session()
    session.mount("https://", HTTPAdapter(max_retries=Retry(
        total=2, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"], raise_on_status=False)))
    session.headers.update({"User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")})
    return session


def fetch_eurostat_population(session):
    r = session.get(EUROSTAT_POP_URL, timeout=(10, 45))
    if r.status_code != 200:
        raise ValueError("HTTP_" + str(r.status_code))
    payload = r.json()
    if "Population on 1 January" not in str(payload.get("label", "")):
        raise ValueError("Eurostat population dataset label changed")
    periods = payload["dimension"]["time"]["category"]["index"]
    values = payload.get("value", {})
    rows = []
    for period, idx in sorted(periods.items(), key=lambda kv: kv[1]):
        v = values.get(str(idx))
        if v is None:
            continue
        if isinstance(v, bool) or not isinstance(v, (int, float)) or not math.isfinite(v) or not 5e6 < v < 20e6:
            raise ValueError("Eurostat population outside Hungary sanity band")
        rows.append((int(period), float(v)))
    if not rows:
        raise ValueError("No Eurostat population values")
    year, value = rows[-1]
    record("population", "OK",
           detail="Official Eurostat population on 1 January, annual, published with a lag",
           source=EUROSTAT_POP_URL, latest_year=year, latest_value=int(value))
    return year, value


def fetch_worldbank_hun_macro(session, start_year=2020, end_year=2025):
    indicators = {
        "SP.POP.TOTL": "WB_Population_Headcount",
        "NY.GDP.MKTP.CD": "GDP_USD_Million",
        "NE.EXP.GNFS.CD": "Exports_USD_Million",
        "NE.IMP.GNFS.CD": "Imports_USD_Million",
    }
    records = {}
    parsed_codes = set()
    for code, col_name in indicators.items():
        api_url = f"https://api.worldbank.org/v2/country/HUN/indicator/{code}?date={start_year}:{end_year}&format=json"
        r = session.get(api_url, timeout=(10, 20))
        if r.status_code != 200:
            continue
        data = r.json()
        if isinstance(data, list) and len(data) > 1 and isinstance(data[-1], list):
            for item in data[-1]:
                if isinstance(item, dict):
                    yr = int(item.get("date", 0))
                    val = item.get("value")
                    if val is not None and yr > 0:
                        scale = 1.0 if code == "SP.POP.TOTL" else 1e6
                        records.setdefault(yr, {})[col_name] = val / scale
                        parsed_codes.add(code)
    if len(parsed_codes) != len(indicators):
        record("worldbank", "FAILED",
               detail="World Bank macro incomplete; no typed-in fallback is ever substituted",
               indicators_parsed=sorted(parsed_codes),
               requested_years=[start_year, end_year])
        return {}
    record("worldbank", "OK", detail="Historical values, not today's observation",
           indicators_parsed=sorted(parsed_codes), requested_years=[start_year, end_year])
    return records


def _latest_complete_price_day(csv_path="hupx_day_ahead_prices.csv"):
    path = Path(csv_path)
    if not path.exists():
        return None
    with path.open(newline="", encoding="utf-8") as handle:
        rows = list(csv.DictReader(handle))
    if not rows:
        return None
    day = max(r["Delivery_Date"] for r in rows)
    day_rows = [r for r in rows if r["Delivery_Date"] == day]
    if len(day_rows) < 92:  # 96 quarter-hours, 92 on DST switch-on day
        return None
    prices = [float(r["Price_EUR_MWh"]) for r in day_rows]
    return day, round(sum(prices) / len(prices), 2), round(max(prices), 2), round(min(prices), 2)


def sync_master_csv(master_path=MASTER_CSV, status_path=STATUS_JSON):
    if not os.path.exists(master_path):
        record("master", "FAILED", detail="Master CSV missing; sync skipped")
        return
    with open(master_path, newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        columns, rows = reader.fieldnames, list(reader)

    def set_row(record_id, **fields):
        matched = [r for r in rows if r["Record_ID"] == record_id]
        if not matched:
            return
        row = matched[0]
        for key, value in fields.items():
            if key in row:
                row[key] = str(value)

    pop = DATASETS.get("population", {})
    if pop.get("state") == "OK":
        set_row("HUN-0001",
                Primary_Value=f"{pop['latest_value']:,} (Eurostat, 1 Jan {pop['latest_year']})",
                Period_Year=str(pop["latest_year"]),
                Verification_Status_Source_URL=pop.get("source", ""))
    fx = DATASETS.get("fx", {})
    if fx.get("state") == "OK":
        set_row("HUN-0004",
                Primary_Value=f"EUR/HUF {fx['eur_huf']} | USD/HUF {fx['usd_huf']}",
                Secondary_Value=f"MNB official daily rates, publication date {fx['source_date']}; HUF floats, no peg",
                Period_Year=fx["source_date"][:4],
                Verification_Status_Source_URL=fx.get("source", ""))
    gdp = DATASETS.get("current_gdp", {})
    if gdp.get("state") == "OK":
        set_row("HUN-0002",
                Primary_Value=f"{gdp['value_usd_billions']:.1f} ({gdp['classification']}, {gdp['vintage']})",
                Period_Year=str(gdp["year"]),
                Verification_Status_Source_URL=gdp.get("source", ""))
    hicp = DATASETS.get("eurostat_hicp", {})
    if hicp.get("state") == "OK":
        set_row("HUN-0003",
                Primary_Value=f"{hicp['latest_value_pct']}% ({hicp['latest_period']})",
                Period_Year=hicp["latest_period"][:4],
                Verification_Status_Source_URL=hicp.get("source", ""))
    price = DATASETS.get("price_hu", {})
    if price.get("state") == "OK":
        summary = _latest_complete_price_day()
        if summary:
            day, avg, peak, trough = summary
            set_row("HUN-0005",
                    Primary_Value=avg,
                    Secondary_Value=f"Peak: {peak:.2f} EUR/MWh | Trough: {trough:.2f} EUR/MWh | delivery day {day}",
                    Tariff_Price_Rate_Value=f"{avg:.2f} EUR/MWh (HUPX day-ahead, wholesale spot - not a customer tariff)",
                    Period_Year=day[:4],
                    Verification_Status_Source_URL=price.get("source", ""))
    th = DATASETS.get("eurostat_tariff_household", {})
    if th.get("state") == "OK":
        set_row("HUN-0006",
                Primary_Value=th["latest_value_eur_per_kwh"],
                Tariff_Price_Rate_Value=f"{th['latest_value_eur_per_kwh']:.4f} EUR/kWh ({th['latest_period']}, all taxes included)",
                Period_Year=th["latest_period"][:4],
                Verification_Status_Source_URL=th.get("source", ""))
    tn = DATASETS.get("eurostat_tariff_nonhousehold", {})
    if tn.get("state") == "OK":
        set_row("HUN-0007",
                Primary_Value=tn["latest_value_eur_per_kwh"],
                Tariff_Price_Rate_Value=f"{tn['latest_value_eur_per_kwh']:.4f} EUR/kWh ({tn['latest_period']}, all taxes included)",
                Period_Year=tn["latest_period"][:4],
                Verification_Status_Source_URL=tn.get("source", ""))
    agg = DATASETS.get("brp_register", {})
    if agg.get("state") == "OK":
        set_row("HUN-0008",
                Primary_Value=agg["party_count"],
                Period_Year=str(dt.datetime.now(UTC).year),
                Verification_Status_Source_URL=agg.get("source", ""))

    out = io.StringIO(newline="")
    writer = csv.DictWriter(out, fieldnames=columns, lineterminator="\n")
    writer.writeheader()
    writer.writerows(rows)
    atomic_text(master_path, out.getvalue())
    write_status(status_path, rows=len(rows))
    print(f"   [SUCCESS] Updated {master_path} ({len(rows)} rows) and {status_path}.")


def main():
    session = make_session()

    def call(name, func, default=None):
        try:
            return func()
        except Exception as exc:
            if name not in DATASETS:
                record(name, "FAILED", detail=type(exc).__name__)
            return default

    call("mavir", lambda: mavir_adapter.update_all(session=session))
    call("price_hu", lambda: hupx_prices.update_hupx_prices(session=session))
    call("fx", lambda: mnb_fx.update_mnb_fx(session=session))
    call("current_gdp", lambda: imf_gdp.update_current_gdp())  # own curl_cffi session; IMF blocks plain requests
    call("eurostat", lambda: eurostat_feeds.update_eurostat(session=session))
    call("brp_register", lambda: brp_feed.update_brp_register(session=session))
    call("population", lambda: fetch_eurostat_population(session))
    call("worldbank", lambda: fetch_worldbank_hun_macro(session))

    try:
        sync_master_csv()
    except Exception as exc:
        record("master", "FAILED", detail=type(exc).__name__)
        write_status(STATUS_JSON)
    print("Hungary CSV update finished. See per-dataset status; this is not full-master verification.")


if __name__ == "__main__":
    main()
