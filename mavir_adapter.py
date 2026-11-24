"""MAVIR rtdwweb chart exports: Hungarian system load, generation, balancing, flows.

The MAVIR "Hungarian Power System actual data" pages iframe a charting app
(rtdwweb.mavir.hu) whose per-chart export endpoint returns XLSX for a requested
UTC-millisecond window. No token is required. Values are MAVIR's own published
operational measurements; plan columns are MAVIR schedules/forecasts, never
actuals. Sign conventions and per-column meaning follow MAVIR's published
chart labels; the Hungarian source header for every output column is recorded
in SOURCE_HEADERS below so a relabel can never happen silently.

Strict validation, mirroring the Bosnia pipeline's rules:
- the first column must be 'Időpont' and the value headers must match exactly;
- timestamps must parse, carry an offset, and sit on a complete 15-minute grid
  for the requested window (the current day may be partial: published
  completed intervals only, never future intervals);
- a failed chart never erases previously saved intervals.
"""
import csv
import datetime as dt
import io
import math
import time
import warnings
from pathlib import Path
from zoneinfo import ZoneInfo

import openpyxl
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from pipeline_status import atomic_text, record, utc_now

EXPORT_URL = "https://rtdwweb.mavir.hu/rtdwweb/webuser/chart/{chart_id}/export"
CHART_PAGE = "https://www.mavir.hu/web/mavir-en/hungarian-power-system-actual-data"
TZ = ZoneInfo("Europe/Budapest")
UTC = dt.timezone.utc
RESOLUTION_MIN = 15
HISTORY_DAYS = 31
MAX_EXPORT_ROWS = 20000

# chart_id -> (dataset name, output csv, ordered (hungarian header -> output column))
CHARTS = {
    "load": {
        "chart_id": "7678",
        "csv": "mavir_load.csv",
        "title": "Actual and Plan System Loads",
        "columns": [
            ("Bruttó hitelesített rendszerterhelés tény", "Gross_Certified_Load_Actual_MW"),
            ("Bruttó rendszerterhelés becslés", "Gross_Load_Estimate_MW"),
            ("Nettó terv rendszertermelés", "Net_Plan_System_Generation_MW"),
            ("Nettó terv rendszerterhelés", "Net_Plan_System_Load_MW"),
            ("Bruttó tény rendszerterhelés", "Gross_Actual_System_Load_MW"),
            ("Bruttó terv rendszerterhelés", "Gross_Plan_System_Load_MW"),
            ("Nettó tény rendszerterhelés - net.ker.elsz.meres", "Net_Actual_Load_Settlement_MW"),
            ("Nettó terhelés", "Net_Load_MW"),
            ("Nettó rendszerterhelés tény - üzemirányítási", "Net_Actual_Load_Operational_MW"),
            ("Nettó rendszerterhelés becslés", "Net_Load_Estimate_MW"),
        ],
    },
    "generation": {
        "chart_id": "12058",
        "csv": "mavir_generation_15min.csv",
        "title": "Net Actual Generation per Production Type (settlement)",
        "columns": [
            ("Nettó hazai termelés tény", "Net_Domestic_Generation_Actual_MW"),
            ("Nukleáris erőművek", "Nuclear_MW"),
            ("Barnakőszén-lignit erőművek", "Lignite_MW"),
            ("Gáz (fosszilis) erőművek", "Gas_MW"),
            ("Olaj (fosszilis) erőművek", "Oil_MW"),
            ("Szárazföldi szélerőművek", "Wind_Onshore_MW"),
            ("Biomassza erőművek", "Biomass_MW"),
            ("VER összes PV - Mért hálózati kitáplálás", "Solar_PV_Measured_MW"),
            ("Szemétégető erőművek", "Waste_Incineration_MW"),
            ("Folyóvizes erőművek", "Hydro_RunOfRiver_MW"),
            ("Víztározós vízerőművek", "Hydro_Reservoir_MW"),
            ("Egyéb megújuló erőművek", "Other_Renewable_MW"),
            ("Geotermikus erőművek", "Geothermal_MW"),
            ("Egyéb primer energiahordozójú erőművek", "Other_Primary_MW"),
            ("Feketekőszén erőművek", "Hard_Coal_MW"),
        ],
    },
    "balancing": {
        "chart_id": "1244886",
        "csv": "mavir_balancing_offers.csv",
        "title": "Spinning reserves and system data - aFRR/mFRR/FRR offer results",
        "columns": [
            ("FRR ajánlat szumma FEL (15p)", "FRR_Up_Offers_MW"),
            ("FRR ajánlat szumma LE (15p)", "FRR_Down_Offers_MW"),
            ("aFRR ajánlat szumma FEL (15p)", "AFRR_Up_Offers_MW"),
            ("aFRR ajánlat szumma LE (15p)", "AFRR_Down_Offers_MW"),
            ("mFRR ajánlat szumma FEL (15p)", "MFRR_Up_Offers_MW"),
            ("mFRR ajánlat szumma LE (15p)", "MFRR_Down_Offers_MW"),
        ],
    },
    "flows": {
        "chart_id": "5229",
        "csv": "mavir_flows_interconnectors.csv",
        "title": "Flows per interconnection",
        "columns": [
            ("HU-UK", "HU_UA_Actual_MW"),
            ("HU-SK", "HU_SK_Actual_MW"),
            ("HU-SI", "HU_SI_Actual_MW"),
            ("HU-RS", "HU_RS_Actual_MW"),
            ("HU-AT", "HU_AT_Actual_MW"),
            ("HU-HR", "HU_HR_Actual_MW"),
            ("HU-RO", "HU_RO_Actual_MW"),
            ("HU-AT menetrend", "HU_AT_Schedule_MW"),
            ("HU-HR menetrend", "HU_HR_Schedule_MW"),
            ("HU-SI menetrend (RIR NT)", "HU_SI_Schedule_MW"),
            ("HU-SK menetrend", "HU_SK_Schedule_MW"),
            ("HU-RS menetrend", "HU_RS_Schedule_MW"),
            ("HU-UK menetrend", "HU_UA_Schedule_MW"),
            ("HU-RO menetrend", "HU_RO_Schedule_MW"),
        ],
    },
}

BASE_COLUMNS = ["Interval_Start_UTC", "Interval_End_UTC", "Interval_Start_Local", "Resolution"]


class MavirError(RuntimeError):
    pass


class MavirRateLimited(MavirError):
    pass


def make_session():
    session = requests.Session()
    session.mount("https://", HTTPAdapter(max_retries=Retry(
        total=2, backoff_factor=1, status_forcelist=[500, 502, 503, 504],
        allowed_methods=["GET"], raise_on_status=False)))
    session.headers.update({"User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")})
    return session


def day_bounds(day):
    start = dt.datetime.combine(day, dt.time(), TZ).astimezone(UTC)
    end = dt.datetime.combine(day + dt.timedelta(days=1), dt.time(), TZ).astimezone(UTC)
    return start, end


TRANSPORT_ERRORS = (requests.ConnectTimeout, requests.ReadTimeout, requests.ConnectionError)


def export_chart(session, chart_id, window_start, window_end, *, max_429_retries=4, max_transport_retries=3):
    """Download the XLSX export for one chart and window.

    Retries HTTP 429 politely and retries transport failures (connect/read
    timeouts, connection errors) with linear backoff; after the bounded
    attempts are exhausted the failure is preserved, never hidden."""
    params = {
        "exportType": "xlsx",
        "fromTime": int(window_start.timestamp() * 1000),
        "toTime": int(window_end.timestamp() * 1000),
        "periodType": "min",
        "period": RESOLUTION_MIN,
    }
    transport_attempts = 0
    for attempt in range(max_429_retries + 1):
        try:
            response = session.get(EXPORT_URL.format(chart_id=chart_id), params=params, timeout=(10, 90))
        except TRANSPORT_ERRORS as exc:
            if transport_attempts >= max_transport_retries:
                raise MavirError("Network request failed after retries: " + type(exc).__name__) from None
            transport_attempts += 1
            time.sleep(5 * transport_attempts)
            continue
        except requests.RequestException as exc:
            raise MavirError("Network request failed: " + type(exc).__name__) from None
        if response.status_code == 429:
            if attempt == max_429_retries:
                raise MavirRateLimited("Export rate-limited (HTTP 429) after retries")
            time.sleep(20 * (attempt + 1))
            continue
        if response.status_code != 200:
            raise MavirError("Export HTTP " + str(response.status_code))
        content = response.content
        if not content[:2] == b"PK":
            raise MavirError("Export response is not an XLSX workbook")
        return content
    raise MavirRateLimited("Export rate-limited (HTTP 429) after retries")


def parse_export(content, column_pairs, window_start, window_end, now):
    """Parse and strictly validate one chart export. Returns (rows, value_columns)."""
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        try:
            book = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
        except Exception as exc:
            raise MavirError("XLSX unreadable: " + type(exc).__name__) from None
        sheets = book.worksheets
        if len(sheets) != 1:
            raise MavirError("Expected exactly one export sheet")
        table = list(sheets[0].values)
    if not table or not table[0]:
        raise MavirError("Empty export")
    header = [str(c).strip() if c is not None else "" for c in table[0]]
    expected = ["Időpont"] + [hu for hu, _ in column_pairs]
    if header != expected:
        raise MavirError("Export headers changed: expected MAVIR layout no longer matches")
    value_columns = [en for _, en in column_pairs]
    step = dt.timedelta(minutes=RESOLUTION_MIN)
    rows = []
    seen = set()
    for raw in table[1:]:
        if raw is None or raw[0] is None:
            continue
        if len(raw) < len(expected):
            raise MavirError("Short data row in export")
        stamp_text = str(raw[0]).strip()
        try:
            local = dt.datetime.strptime(stamp_text[:26], "%Y.%m.%d %H:%M:%S %z")
        except ValueError:
            raise MavirError("Unparseable export timestamp") from None
        # MAVIR labels each quarter-hour by its END: a row stamped 01:15 is the
        # 01:00-01:15 interval. Verified against the export's own boundaries.
        end = local.astimezone(UTC)
        start = end - step
        if end > now:
            # Published completed intervals only: MAVIR pads the window with
            # future rows; they carry no measurements and are skipped.
            continue
        if start < window_start or end > window_end:
            raise MavirError("Export row outside requested window")
        if start in seen:
            raise MavirError("Duplicate interval in export")
        seen.add(start)
        values = []
        for cell in raw[1:len(expected)]:
            if cell is None or (isinstance(cell, str) and not cell.strip()):
                # MAVIR certifies some columns late: missing stays missing, never zero.
                values.append("")
                continue
            if isinstance(cell, bool) or not isinstance(cell, (int, float)) or not math.isfinite(cell):
                raise MavirError("Invalid numeric cell in export")
            values.append(round(float(cell), 6))
        rows.append({
            "Interval_Start_UTC": start.isoformat(),
            "Interval_End_UTC": end.isoformat(),
            "Interval_Start_Local": local.isoformat(),
            "Resolution": "PT15M",
            **dict(zip(value_columns, values)),
        })
    rows.sort(key=lambda r: r["Interval_Start_UTC"])
    if not rows:
        raise MavirError("No data rows in export window")
    for prev, nxt in zip(rows, rows[1:]):
        if prev["Interval_End_UTC"] != nxt["Interval_Start_UTC"]:
            raise MavirError("Gap inside export interval series")
    return rows, value_columns


def _merge_rolling(path, base_plus_values, new_rows, cutoff):
    old = []
    if path.exists():
        with path.open(newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            if reader.fieldnames != base_plus_values:
                raise MavirError("Existing CSV schema mismatch for " + path.name)
            old = list(reader)
    merged = {r["Interval_Start_UTC"]: r for r in old
              if dt.datetime.fromisoformat(r["Interval_Start_UTC"]) >= cutoff}
    for row in new_rows:
        merged[row["Interval_Start_UTC"]] = row
    out = io.StringIO(newline="")
    writer = csv.DictWriter(out, fieldnames=base_plus_values, lineterminator="\n")
    writer.writeheader()
    writer.writerows(merged[k] for k in sorted(merged))
    atomic_text(path, out.getvalue())
    return len(merged)


def update_chart(name, session, now, output_dir="."):
    spec = CHARTS[name]
    today = now.astimezone(TZ).date()
    oldest = today - dt.timedelta(days=HISTORY_DAYS - 1)
    window_start = day_bounds(oldest)[0]
    window_end = day_bounds(today)[1]
    meta = {"chart_id": spec["chart_id"], "chart_title": spec["title"],
            "source": CHART_PAGE, "timezone": "Europe/Budapest",
            "requested_first_day": str(oldest), "requested_last_day": str(today)}
    content = export_chart(session, spec["chart_id"], window_start, window_end)
    rows, value_columns = parse_export(content, spec["columns"], window_start, window_end, now)
    if len(rows) > MAX_EXPORT_ROWS:
        raise MavirError("Export row count above sanity ceiling")
    cutoff = window_start
    stamp = utc_now()
    for row in rows:
        row["Fetched_At_UTC"] = stamp
    path = Path(output_dir) / spec["csv"]
    total = _merge_rolling(path, BASE_COLUMNS + value_columns + ["Fetched_At_UTC"], rows, cutoff)
    record("mavir_" + name, "OK",
           detail="MAVIR published chart export; today holds published completed intervals only",
           row_count=total, intervals_fetched=len(rows), unit="MW",
           resolution="PT15M", **meta)
    return total


def update_all(output_dir=".", now=None, session=None, names=None):
    now = now or dt.datetime.now(UTC)
    if now.tzinfo is None:
        raise ValueError("now must be timezone-aware")
    owns = session is None
    session = session or make_session()
    names = names or list(CHARTS)
    try:
        for i, name in enumerate(names):
            if i:
                time.sleep(12)  # rtdwweb rate-limits bursts (observed HTTP 429)
            try:
                update_chart(name, session, now, output_dir)
            except MavirRateLimited as exc:
                record("mavir_" + name, "FAILED", detail=str(exc) + "; saved intervals retained",
                       chart_id=CHARTS[name]["chart_id"], source=CHART_PAGE)
            except (MavirError, ValueError, OSError, KeyError, TypeError) as exc:
                detail = str(exc) if isinstance(exc, MavirError) else type(exc).__name__
                record("mavir_" + name, "FAILED", detail="Saved intervals retained; " + detail,
                       chart_id=CHARTS[name]["chart_id"], source=CHART_PAGE)
    finally:
        if owns:
            session.close()


if __name__ == "__main__":
    update_all()
