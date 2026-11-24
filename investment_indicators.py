"""Derived investment-screening indicators from data this repo already fetches.

Two daily datasets, computed from hupx_day_ahead_prices.csv and
mavir_generation_15min.csv only (no new external sources):

- solar_capture_daily.csv: for each complete Budapest delivery day, the
  solar-generation-weighted average day-ahead price, the plain day average
  price, and their ratio (solar capture indicator; below 1 means solar
  output concentrated in cheaper hours). National level, not a site's
  revenue. The ratio is left blank when the day average price is zero or
  negative, or when measured solar generation is zero (meaningless there,
  never forced).
- bess_spread_daily.csv: for each complete price day and each scenario
  window (1, 2 and 4 hours), the best chronological charge window before a
  discharge window of the same length, as a raw day-ahead price spread in
  EUR/MWh. Gross screening indicator only: no efficiency, degradation,
  fees, capacity or opportunity-cost assumptions, so it is not profit. A
  day whose best chronological spread is negative is reported as such (no
  profitable window that day).

Days are attributed by Budapest local delivery date derived from UTC
interval starts (92/96/100 intervals, so DST switch days are handled); a
day computes only when its source series is complete for it, otherwise it
is left out whole, never estimated. Rows already stored for days that have
rolled out of the source window are retained unchanged.
"""
import csv
import datetime as dt
import io
from pathlib import Path
from zoneinfo import ZoneInfo

from pipeline_status import DATASETS, atomic_text, record

TZ = ZoneInfo("Europe/Budapest")
UTC = dt.timezone.utc
QH = dt.timedelta(minutes=15)
WINDOW_HOURS = (1, 2, 4)   # scenario charge/discharge window lengths

PRICE_CSV = "hupx_day_ahead_prices.csv"
GEN_CSV = "mavir_generation_15min.csv"
SOLAR_CSV = "solar_capture_daily.csv"
BESS_CSV = "bess_spread_daily.csv"

SOLAR_HEADER = ["Delivery_Date", "Day_Avg_Price_EUR_MWh", "Solar_Weighted_Avg_Price_EUR_MWh",
                "Solar_Capture_Ratio", "Solar_Generation_MWh", "Note"]
BESS_HEADER = ["Delivery_Date", "Window_Hours", "Charge_Window_Start_Local",
               "Charge_Window_End_Local", "Charge_Avg_Price_EUR_MWh",
               "Discharge_Window_Start_Local", "Discharge_Window_End_Local",
               "Discharge_Avg_Price_EUR_MWh", "Spread_EUR_MWh", "Note", "Day_Window_Key"]


def _utc(text):
    return dt.datetime.fromisoformat(text.replace("Z", "+00:00")).astimezone(UTC)


def _local_day(ts_utc):
    return ts_utc.astimezone(TZ).date()


def _local_text(ts_utc):
    return ts_utc.astimezone(TZ).strftime("%Y-%m-%d %H:%M")


def _expected_intervals(day):
    start = dt.datetime.combine(day, dt.time(), tzinfo=TZ).astimezone(UTC)
    end = dt.datetime.combine(day + dt.timedelta(days=1), dt.time(), tzinfo=TZ).astimezone(UTC)
    return int((end - start).total_seconds() // 900)


def _read(path):
    path = Path(path)
    if not path.exists():
        return []
    with path.open(newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def _series_by_day(rows, value_key):
    """delivery day -> {utc_start: float value}; unparsable rows dropped."""
    days = {}
    for r in rows:
        try:
            ts = _utc(r["Interval_Start_UTC"])
            value = float(r[value_key])
        except (KeyError, ValueError):
            continue
        days.setdefault(_local_day(ts), {})[ts] = value
    return days


def _complete(day_map, day):
    got = day_map.get(day)
    if not got:
        return False
    n = _expected_intervals(day)
    if len(got) != n:
        return False
    start = min(got)
    return all((start + QH * i) in got for i in range(n))


def _merge_rows(path, header, new_rows, key="Delivery_Date"):
    old = {tuple(r[k] for k in key) if isinstance(key, tuple) else r[key]: r for r in _read(path)}
    old.update(new_rows)
    out = io.StringIO(newline="")
    w = csv.DictWriter(out, fieldnames=header, lineterminator="\n")
    w.writeheader()
    for k in sorted(old):
        w.writerow(old[k])
    atomic_text(Path(path), out.getvalue())
    return len(old)


def update_solar_capture(output_dir="."):
    """Solar-capture indicator per complete Budapest delivery day."""
    prices = _series_by_day(_read(Path(output_dir) / PRICE_CSV), "Price_EUR_MWh")
    solar = _series_by_day(_read(Path(output_dir) / GEN_CSV), "Solar_PV_Measured_MW")
    rows, skipped = {}, []
    for day in sorted(set(prices) & set(solar)):
        if not (_complete(prices, day) and _complete(solar, day)):
            skipped.append(str(day))
            continue
        ts_order = sorted(prices[day])
        day_avg = sum(prices[day][t] for t in ts_order) / len(ts_order)
        solar_mwh = sum(solar[day][t] * 0.25 for t in ts_order)
        solar_mw_sum = sum(solar[day][t] for t in ts_order)
        note = []
        if solar_mw_sum > 0:
            sw_avg = sum(prices[day][t] * solar[day][t] for t in ts_order) / solar_mw_sum
            sw_text = f"{sw_avg:.4f}"
        else:
            sw_avg = None
            sw_text = ""
            note.append("no measured solar generation this day")
        if sw_avg is not None and day_avg > 0:
            ratio_text = f"{sw_avg / day_avg:.4f}"
        else:
            if day_avg <= 0:
                note.append("day average price zero/negative; ratio undefined")
            ratio_text = ""
        rows[str(day)] = {"Delivery_Date": str(day),
                          "Day_Avg_Price_EUR_MWh": f"{day_avg:.2f}",
                          "Solar_Weighted_Avg_Price_EUR_MWh": sw_text,
                          "Solar_Capture_Ratio": ratio_text,
                          "Solar_Generation_MWh": f"{solar_mwh:.1f}",
                          "Note": "; ".join(note)}
    n = _merge_rows(Path(output_dir) / SOLAR_CSV, SOLAR_HEADER, rows)
    record("solar_capture_daily", "OK",
           detail="Solar-generation-weighted day-ahead price vs plain day average, complete Budapest "
                  "delivery days only; ratio blank when the day average is zero/negative or solar output "
                  "is zero. National level, not site revenue. Sources: HUPX/energy-charts DAM + MAVIR "
                  "measured solar.",
           days_computed=len(rows), days_skipped_incomplete=skipped,
           latest_day=max(rows) if rows else "")
    return n


def update_bess_spread(output_dir="."):
    """Best chronological charge->discharge raw price spread per complete day and window length."""
    prices = _series_by_day(_read(Path(output_dir) / PRICE_CSV), "Price_EUR_MWh")
    rows, skipped = {}, []
    for day in sorted(prices):
        if not _complete(prices, day):
            skipped.append(str(day))
            continue
        ts_order = sorted(prices[day])
        vals = [prices[day][t] for t in ts_order]
        for hours in WINDOW_HOURS:
            wq = hours * 4
            wins = [(ts_order[i], sum(vals[i:i + wq]) / wq)
                    for i in range(len(vals) - wq + 1)]
            prefix_min = []
            for ts, avg in wins:
                prefix_min.append((ts, avg) if not prefix_min or avg < prefix_min[-1][1] else prefix_min[-1])
            best = None
            for j, (dstart, davg) in enumerate(wins):
                last_charge = j - wq   # charge window must end before discharge starts
                if last_charge < 0:
                    continue
                cstart, cavg = prefix_min[last_charge]
                spread = davg - cavg
                if best is None or spread > best[0]:
                    best = (spread, cstart, dstart, cavg, davg)
            if best is None:
                continue
            spread, cstart, dstart, cavg, davg = best
            rows[(str(day), str(hours))] = {
                "Delivery_Date": str(day), "Window_Hours": str(hours),
                "Charge_Window_Start_Local": _local_text(cstart),
                "Charge_Window_End_Local": _local_text(cstart + QH * wq),
                "Charge_Avg_Price_EUR_MWh": f"{cavg:.2f}",
                "Discharge_Window_Start_Local": _local_text(dstart),
                "Discharge_Window_End_Local": _local_text(dstart + QH * wq),
                "Discharge_Avg_Price_EUR_MWh": f"{davg:.2f}",
                "Spread_EUR_MWh": f"{spread:.2f}",
                "Note": "" if spread > 0 else "no profitable chronological window this day",
                "Day_Window_Key": f"{day}|{hours}"}
    n = _merge_rows(Path(output_dir) / BESS_CSV, BESS_HEADER, rows, key=("Delivery_Date", "Window_Hours"))
    days = sorted({k[0] for k in rows})
    record("bess_spread_daily", "OK",
           detail="Best chronological charge window before a same-length discharge window per complete "
                  "Budapest delivery day, for 1h/2h/4h scenarios; raw day-ahead price spread in EUR/MWh. "
                  "Gross screening indicator: no efficiency, degradation, fees, capacity or "
                  "opportunity-cost assumptions; not profit.",
           days_computed=len(days), days_skipped_incomplete=skipped,
           latest_day=max(days) if days else "")
    return n
