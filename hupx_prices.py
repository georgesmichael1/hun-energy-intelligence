"""Hungarian day-ahead wholesale spot prices by published interval (HUPX zone).

Primary source: HUPX's own Labs publication, DAM aggregated trading data
(15-minute MTU) - the market operator's official day-ahead results, fetched
from the public labs.hupx.hu data endpoint; every row is required to carry
Status "final". Fallback source: Fraunhofer ISE energy-charts (bzn=HU), whose
data is CC BY 4.0 from Bundesnetzagentur | SMARD.de; the licence string is
validated on every run and output stops if it changes. With an ENTSO-E token,
the official Transparency Platform A44 auction for the HU bidding zone
(primary owner: HUPX) is fetched by entsoe_adapter.py instead, in the token
workflow.

These are wholesale market clearing prices. They are not end-user tariffs;
customer tariffs live separately in eurostat_tariffs.csv.

Validation mirrors the Bosnia pipeline: a delivery day must be complete
(including 23/25-hour DST days), interval-contiguous, in EUR/MWh, and every
price finite (negative prices are valid). A failed day never erases saved days.
"""
import csv
import datetime as dt
import io
import math
import json
from pathlib import Path
from zoneinfo import ZoneInfo

import requests
from pipeline_status import atomic_text, record, utc_now

TZ = ZoneInfo("Europe/Budapest")
UTC = dt.timezone.utc
HU_SOURCE = "https://api.energy-charts.info/price?bzn=HU"
HUPX_SOURCE = "https://labs.hupx.hu/data/v1/dam_aggregated_trading_data_15min"
HUPX_ATTRIBUTION = "HUPX Hungarian Power Exchange - Labs DAM aggregated trading data, 15-minute MTU"
HUPX_LICENSE_URL = "https://labs.hupx.hu/about_us/terms_of_use"
ATTRIBUTION = "Bundesnetzagentur | SMARD.de, via Fraunhofer ISE Energy-Charts; underlying market: HUPX day-ahead"
LICENSE_URL = "https://creativecommons.org/licenses/by/4.0/"
COLUMNS = ["Market", "Country", "Product", "Delivery_Date", "Interval_Start_UTC",
           "Interval_End_UTC", "Price_EUR_MWh", "Source", "Attribution", "License", "Fetched_At"]
HISTORY_DAYS = 31  # run-local today plus the previous 30 delivery days


def parse_hu_day(payload, delivery_day):
    """Require one complete local delivery day, including 23/25-hour DST days."""
    license_info = str(payload.get("license_info", ""))
    if not ("CC BY 4.0" in license_info and "SMARD.de" in license_info):
        raise ValueError("HU response no longer carries the verified CC BY/SMARD licence")
    if payload.get("deprecated") is True:
        raise ValueError("HU source is deprecated")
    if payload.get("unit") not in ("EUR / MWh", "EUR/MWh"):
        raise ValueError("Unexpected HU price unit")
    times, prices = payload.get("unix_seconds"), payload.get("price")
    if not isinstance(times, list) or not isinstance(prices, list) or len(times) != len(prices):
        raise ValueError("Mismatched price and timestamp arrays")
    start = dt.datetime.combine(delivery_day, dt.time(), TZ).astimezone(UTC)
    end = dt.datetime.combine(delivery_day + dt.timedelta(days=1), dt.time(), TZ).astimezone(UTC)
    interval = 900  # Hungary's native day-ahead market time unit since October 2025.
    expected = list(range(int(start.timestamp()), int(end.timestamp()), interval))
    if any(isinstance(x, bool) or not isinstance(x, int) for x in times) or times != expected:
        raise ValueError("HU delivery day has gaps, duplicates, wrong order or wrong boundaries")
    checked = []
    for value in prices:
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
            raise ValueError("Missing or invalid HU price")
        checked.append(float(value))  # Negative and zero prices are valid.
    return [{"Market": "HU", "Country": "Hungary", "Product": "DAY_AHEAD",
             "Delivery_Date": delivery_day.isoformat(),
             "Interval_Start_UTC": dt.datetime.fromtimestamp(t, UTC).isoformat(),
             "Interval_End_UTC": dt.datetime.fromtimestamp(t + interval, UTC).isoformat(),
             "Price_EUR_MWh": p, "Source": HU_SOURCE, "Attribution": ATTRIBUTION,
             "License": LICENSE_URL}
            for t, p in zip(times, checked)]


def expected_count(day):
    """Complete local delivery days hold 92, 96 or 100 quarter-hour intervals (DST)."""
    s = dt.datetime.combine(day, dt.time(), TZ).astimezone(UTC)
    e = dt.datetime.combine(day + dt.timedelta(days=1), dt.time(), TZ).astimezone(UTC)
    return int((e - s).total_seconds() // 900)


def parse_hupx_day(rows, delivery_day):
    """Strictly validated local delivery day from HUPX Labs (quarter-hour MTU).

    Mirrors parse_hu_day: complete 1..N quarter-hour sequence (92/96/100 for
    23/24/25-hour days), Region HU only, every row Status final, every price
    finite (negative prices are valid). One bad row rejects the whole day."""
    expected = expected_count(delivery_day)
    if len(rows) != expected:
        raise ValueError("HUPX day holds " + str(len(rows)) + " rows, expected " + str(expected))
    for row in rows:
        if row.get("Region") != "HU":
            raise ValueError("Unexpected region in HUPX rows")
        if str(row.get("Status", "")).lower() != "final":
            raise ValueError("HUPX day holds non-final rows")
    try:
        order = [int(row["ProductQH"]) for row in rows]
    except (KeyError, TypeError, ValueError):
        raise ValueError("HUPX quarter-hour labels unreadable") from None
    if sorted(order) != list(range(1, expected + 1)):
        raise ValueError("HUPX quarter-hours are not a complete 1..N sequence")
    by_qh = {int(row["ProductQH"]): row for row in rows}
    start = dt.datetime.combine(delivery_day, dt.time(), TZ).astimezone(UTC)
    out = []
    for qh in range(1, expected + 1):
        value = by_qh[qh].get("Price")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(value):
            raise ValueError("Missing or invalid HUPX price")
        t0 = start + dt.timedelta(seconds=900 * (qh - 1))
        out.append({"Market": "HU", "Country": "Hungary", "Product": "DAY_AHEAD",
                    "Delivery_Date": delivery_day.isoformat(),
                    "Interval_Start_UTC": t0.isoformat(),
                    "Interval_End_UTC": (t0 + dt.timedelta(seconds=900)).isoformat(),
                    "Price_EUR_MWh": float(value), "Source": HUPX_SOURCE,
                    "Attribution": HUPX_ATTRIBUTION, "License": HUPX_LICENSE_URL})
    return out


def fetch_hupx_days(session, oldest, today):
    """One range request to HUPX Labs covering the whole window.

    Returns {date: validated rows}; incomplete or non-final days are dropped
    whole (retained history below decides), never partially substituted."""
    params = {"filter": "DeliveryDay__gte__" + oldest.isoformat()
                        + ",DeliveryDay__lte__" + today.isoformat() + ",Region__in__HU"}
    response = session.get(HUPX_SOURCE, params=params, timeout=(10, 60))
    if response.status_code != 200:
        raise ValueError("HTTP_" + str(response.status_code))
    payload = response.json()
    rows = payload.get("data")
    if not isinstance(rows, list):
        raise ValueError("HUPX response has no data array")
    grouped = {}
    for row in rows:
        try:
            day = dt.date.fromisoformat(str(row.get("DeliveryDay", ""))[:10])
        except ValueError:
            continue
        grouped.setdefault(day, []).append(row)
    fresh = {}
    for day, day_rows in grouped.items():
        try:
            fresh[day] = parse_hupx_day(day_rows, day)
        except ValueError:
            pass  # source-gapped day: retained history below decides
    return fresh


def update_hupx_prices(csv_path="hupx_day_ahead_prices.csv", session=None, now=None):
    """Incremental design over the dynamic 31-day window (run-local
    today + previous 30 days): refresh today and yesterday, refill any missing or
    incomplete older day, and retain complete saved history. One efficient range
    request covers the whole window; saved good data survives a failed fetch."""
    now = now or dt.datetime.now(UTC)
    if now.tzinfo is None:
        raise ValueError("now must be timezone-aware")
    today = now.astimezone(TZ).date()
    oldest = today - dt.timedelta(days=HISTORY_DAYS - 1)
    stamp = utc_now()
    own_session = session is None
    session = session or requests.Session()
    try:
        saved = []
        path = Path(csv_path)
        if path.exists():
            with path.open(newline="", encoding="utf-8") as handle:
                reader = csv.DictReader(handle)
                if reader.fieldnames != COLUMNS:
                    raise ValueError("Existing HUPX CSV schema mismatch")
                saved = list(reader)
        saved_by_day = {}
        for r in saved:
            saved_by_day.setdefault(r.get("Delivery_Date", ""), []).append(r)
        complete_saved = {d for d, rs in saved_by_day.items() if len(rs) == expected_count(dt.date.fromisoformat(d))}

        fresh_by_day = {}
        hupx_error = None
        fetch_error = None
        try:
            fresh_by_day = fetch_hupx_days(session, oldest, today)
        except (requests.RequestException, ValueError, KeyError, TypeError) as exc:
            hupx_error = type(exc).__name__
        try:
            response = session.get(HU_SOURCE + "&start=" + oldest.isoformat() + "&end=" + today.isoformat(),
                                   timeout=(10, 60))
            if response.status_code != 200:
                raise ValueError("HTTP_" + str(response.status_code))
            payload = response.json()
            license_info = str(payload.get("license_info", ""))
            times, prices = payload.get("unix_seconds"), payload.get("price")
            if not isinstance(times, list) or not isinstance(prices, list) or len(times) != len(prices):
                raise ValueError("Mismatched price and timestamp arrays")
            grouped = {}
            for t, p in zip(times, prices):
                grouped.setdefault(dt.datetime.fromtimestamp(t, TZ).date(), []).append((t, p))
            for day in sorted(grouped):
                if day in fresh_by_day:
                    continue  # HUPX-primary day already validated
                sub = {"license_info": license_info, "unit": payload.get("unit"),
                       "unix_seconds": [t for t, _ in grouped[day]],
                       "price": [p for _, p in grouped[day]]}
                try:
                    fresh_by_day[day] = parse_hu_day(sub, day)
                except ValueError:
                    pass  # source-gapped day: retained history below decides
        except (requests.RequestException, ValueError, KeyError, TypeError) as exc:
            fetch_error = type(exc).__name__

        refresh_from = today - dt.timedelta(days=1)
        rows, missing = [], []
        day = oldest
        while day <= today:
            key = day.isoformat()
            fresh = fresh_by_day.get(day)
            if fresh is not None and (day >= refresh_from or key not in complete_saved):
                rows.extend(fresh)                      # refresh recent / fill gaps
            elif key in complete_saved:
                rows.extend(saved_by_day[key])          # retain complete saved history
            elif fresh is not None:
                rows.extend(fresh)
            else:
                missing.append(key)
            day += dt.timedelta(days=1)

        if not rows:
            raise ValueError("No data: HUPX fetch failed (" + str(hupx_error) + "), fallback fetch failed ("
                             + str(fetch_error) + ") and no saved history")
        for row in rows:
            row["Fetched_At"] = stamp
        out = io.StringIO(newline="")
        writer = csv.DictWriter(out, fieldnames=COLUMNS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(sorted(rows, key=lambda r: r["Interval_Start_UTC"]))
        atomic_text(path, out.getvalue())

        source_note = ("primary HUPX Labs ok" if hupx_error is None else "primary HUPX Labs failed: " + hupx_error) \
            + "; " + ("fallback energy-charts ok" if fetch_error is None else "fallback energy-charts failed: " + fetch_error)
        base = dict(market="HU", product="DAY_AHEAD",
                    source=HUPX_SOURCE + " (primary), " + HU_SOURCE + " (fallback)",
                    delivery_days=[oldest.isoformat(), today.isoformat()],
                    intervals=len(rows), fetched_at=stamp)
        if hupx_error and fetch_error:
            record("price_hu", "FAILED", **base,
                   detail="Both sources failed (" + source_note + "); complete saved history retained, "
                          "missing days stay empty, never zero")
        elif missing:
            record("price_hu", "SOURCE_PARTIAL", **base, last_successful_fetch_at=stamp,
                   detail="No valid prices for " + ",".join(missing)
                          + " from either source - day omitted whole, no retained history substituted; "
                          + source_note + "; wholesale spot, not a customer tariff")
        else:
            record("price_hu", "OK", **base, last_successful_fetch_at=stamp,
                   detail="Incremental refresh of today/yesterday plus missing or incomplete days; "
                          "complete saved history retained; " + source_note
                          + "; wholesale spot, not a customer tariff")
        return rows
    except (ValueError, OSError, KeyError, TypeError) as exc:
        record("price_hu", "FAILED", market="HU", source=HU_SOURCE,
               detail="Update failed: " + type(exc).__name__ + "; saved window retained")
        return None
    finally:
        if own_session:
            session.close()


def main():
    from pipeline_status import DATASETS, write_status
    path = Path("status.json")
    old = json.loads(path.read_text(encoding="utf-8")) if path.exists() else {}
    DATASETS.clear()
    DATASETS.update(old.get("dataset_details", {}))
    update_hupx_prices()
    reserved = {"schema_version", "run_finished_at", "last_verified_date",
                "last_verified_date_is_legacy", "verification_note", "rows",
                "datasets", "dataset_details"}
    legacy = {k: v for k, v in old.items() if k not in reserved}
    write_status(path, rows=old.get("rows"), legacy=legacy)


if __name__ == "__main__":
    main()
