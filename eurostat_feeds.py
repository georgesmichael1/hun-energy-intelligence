"""Eurostat JSON-stat feeds for Hungary: HICP inflation and customer electricity tariffs.

Wholesale market prices live in hupx_day_ahead_prices.csv / entsoe_day_ahead_prices.csv.
Everything here is a customer-side official statistic, published with a lag
(HICP monthly, tariffs semi-annual). Tariff values are final customer prices
per kWh; they are never mixed with wholesale EUR/MWh market prices.
"""
import csv
import datetime as dt
import io
import math
from pathlib import Path

import requests
from pipeline_status import atomic_text, record, utc_now

BASE = "https://ec.europa.eu/eurostat/api/dissemination/statistics/1.0/data/"

DATASETS = {
    "hicp": {
        "dataset": "prc_hicp_manr",
        "params": {"geo": "HU", "coicop": "CP00", "unit": "RCH_A", "lang": "EN"},
        "csv": "eurostat_hicp.csv",
        "columns": ["Period", "HICP_Annual_Rate_Pct", "Source", "Retrieved_UTC"],
        "expected_label": "HICP - monthly data (annual rate of change)",
        "period_key": "time",
    },
    "tariff_household": {
        "dataset": "nrg_pc_204",
        "params": {"geo": "HU", "currency": "EUR", "nrg_cons": "KWH2500-4999",
                   "unit": "KWH", "tax": "I_TAX", "lang": "EN"},
        "csv": "eurostat_tariffs.csv",
        "columns": ["Period", "Consumer_Group", "Consumption_Band", "Tax_Level",
                    "Price_EUR_Per_kWh", "Source", "Retrieved_UTC"],
        "expected_label": "Electricity prices for household consumers",
        "period_key": "time",
        "consumer_group": "Household",
        "band": "2500-4999 kWh (band DC)",
    },
    "tariff_nonhousehold": {
        "dataset": "nrg_pc_205",
        "params": {"geo": "HU", "currency": "EUR", "nrg_cons": "MWH500-1999",
                   "unit": "KWH", "tax": "I_TAX", "lang": "EN"},
        "csv": "eurostat_tariffs.csv",
        "columns": ["Period", "Consumer_Group", "Consumption_Band", "Tax_Level",
                    "Price_EUR_Per_kWh", "Source", "Retrieved_UTC"],
        "expected_label": "Electricity prices for non-household consumers",
        "period_key": "time",
        "consumer_group": "Non-household",
        "band": "500-1999 MWh (band IC)",
    },
}


def fetch_jsonstat(session, name, spec):
    url = BASE + spec["dataset"]
    response = session.get(url, params={**spec["params"], "lastTimePeriod": 12}, timeout=(10, 45))
    if response.status_code != 200:
        raise ValueError("HTTP_" + str(response.status_code))
    payload = response.json()
    if spec["expected_label"] not in str(payload.get("label", "")):
        raise ValueError("Eurostat dataset label changed for " + name)
    periods = payload["dimension"][spec["period_key"]]["category"]["index"]
    values = payload.get("value", {})
    rows = []
    for period, idx in sorted(periods.items(), key=lambda kv: kv[1]):
        value = values.get(str(idx))
        if value is None:
            continue
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
            raise ValueError("Invalid Eurostat value for " + name)
        rows.append((period, float(value)))
    if not rows:
        raise ValueError("No Eurostat values returned for " + name)
    return url, rows


def update_eurostat(session=None, now=None):
    now = now or dt.datetime.now(dt.timezone.utc)
    stamp = utc_now()
    own = session is None
    session = session or requests.Session()
    tariffs = []
    try:
        for name, spec in DATASETS.items():
            try:
                url, rows = fetch_jsonstat(session, name, spec)
                if name == "hicp":
                    out_rows = [{"Period": p, "HICP_Annual_Rate_Pct": v,
                                 "Source": url, "Retrieved_UTC": stamp} for p, v in rows]
                    out = io.StringIO(newline="")
                    writer = csv.DictWriter(out, fieldnames=spec["columns"], lineterminator="\n")
                    writer.writeheader()
                    writer.writerows(out_rows)
                    atomic_text(spec["csv"], out.getvalue())
                    latest_period, latest_value = rows[-1]
                    record("eurostat_hicp", "OK",
                           detail="Official HICP annual rate of change, monthly, published with a lag",
                           source=url, latest_period=latest_period, latest_value_pct=latest_value)
                else:
                    for p, v in rows:
                        tariffs.append({"Period": p, "Consumer_Group": spec["consumer_group"],
                                        "Consumption_Band": spec["band"], "Tax_Level": "All taxes and levies included",
                                        "Price_EUR_Per_kWh": v, "Source": url, "Retrieved_UTC": stamp})
                    latest_period, latest_value = rows[-1]
                    record("eurostat_" + name, "OK",
                           detail="Official semi-annual final customer price per kWh; NOT a wholesale market price",
                           source=url, latest_period=latest_period, latest_value_eur_per_kwh=latest_value,
                           consumer_group=spec["consumer_group"], band=spec["band"])
            except Exception as exc:
                detail = str(exc) if isinstance(exc, ValueError) else type(exc).__name__
                record("eurostat_" + name, "FAILED", detail="Saved values retained; " + detail,
                       source=BASE + spec["dataset"])
        if tariffs:
            # Full replace: every run re-reads the same trailing windows, and
            # Eurostat may revise recent periods.
            out = io.StringIO(newline="")
            writer = csv.DictWriter(out, fieldnames=DATASETS["tariff_household"]["columns"], lineterminator="\n")
            writer.writeheader()
            writer.writerows(tariffs)
            atomic_text(DATASETS["tariff_household"]["csv"], out.getvalue())
        return tariffs
    finally:
        if own:
            session.close()


if __name__ == "__main__":
    update_eurostat()
