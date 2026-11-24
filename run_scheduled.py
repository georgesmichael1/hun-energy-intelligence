"""Single orchestrated entrypoint for the Hungary update.

One workflow, one schedule (every 30 minutes) and one manual "Run workflow"
button. Each run always refreshes the token-free feeds. The ENTSO-E token feeds
(previous-day load forecast, official A44 prices, and the rolling 31-day
physical-flow history) run once per day, after the previous Budapest day is
complete, and only when the ENTSOE_API_TOKEN secret is configured. Without the
secret they record SKIPPED (a warning, not a red run); with it, any failure
turns the run red through the workflow gate.

The token-free update and the daily ENTSO-E batch run CONCURRENTLY: disjoint
hosts, sessions, dataset keys and CSV files, and pipeline_status serializes
record()/write_status() internally, so neither side can clobber the other.
The HUPX-omitted-day gap fill runs only after both sides join, because it
needs this run's fresh price_hu record. A manual dispatch can set
ENTSOE_FORCE_DAILY=1 to rerun the daily batch regardless of the once-a-day
gate (used for honest speed measurements and operator-initiated recovery);
scheduled runs always keep the gate.
"""
import datetime as dt
import json
import os
import threading
from pathlib import Path

from pipeline_status import DATASETS, record, write_status

ENTSOE_DATASETS = ("entsoe_load_forecast", "entsoe_prices", "entsoe_physical_flows_31d")
DAILY_AFTER_LOCAL_HOUR = 3   # run the daily ENTSO-E batch from 03:30 Budapest
DAILY_AFTER_LOCAL_MINUTE = 30


def _scheduled_today(now_budapest):
    return now_budapest.replace(hour=DAILY_AFTER_LOCAL_HOUR, minute=DAILY_AFTER_LOCAL_MINUTE,
                                second=0, microsecond=0)


def entsoe_due_by_schedule(now_budapest, status_path="status.json"):
    """Time-and-status part of the daily gate: after 03:30 Budapest, when the
    status file is unreadable, a hard physical-flow failure needs a retry, or
    the last successful batch predates today's boundary. The HUPX-omitted-day
    trigger is deliberately NOT checked here: it needs this run's fresh
    price_hu record, which only exists after the token-free update."""
    if now_budapest < _scheduled_today(now_budapest):
        return False
    try:
        old = json.loads(Path(status_path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return True
    details = old.get("dataset_details", {})
    # Recover: a hard physical-flow failure (transport/validation/auth/unexplained gap) is retried
    # on the next scheduled run. SOURCE_PARTIAL (upstream publishes nothing) is not retried daily.
    if (details.get("entsoe_physical_flows_31d", {}) or {}).get("state") in ("PARTIAL", "FAILED"):
        return True
    last = (details.get("entsoe_load_forecast", {}) or {}).get("last_success_at")
    if not last:
        return True
    try:
        last_dt = dt.datetime.fromisoformat(last)
    except ValueError:
        return True
    return last_dt < _scheduled_today(now_budapest).astimezone(dt.timezone.utc)


def entsoe_due(now_budapest, status_path="status.json"):
    """Full daily gate; the omitted-day trigger is valid only once this run's
    price_hu record exists (i.e. after the token-free update)."""
    if now_budapest < _scheduled_today(now_budapest):
        return False
    if entsoe_due_by_schedule(now_budapest, status_path):
        return True
    try:
        old = json.loads(Path(status_path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return True
    details = old.get("dataset_details", {})
    if omitted_price_days() and not details.get("entsoe_a44_gapfill"):
        return True   # first sight of a HUPX-omitted day: read A44 once, in the separate dataset
    return False


def omitted_price_days():
    """Delivery days the HUPX/energy-charts feed omitted (read from this run's price_hu record)."""
    import re
    detail = DATASETS.get("price_hu", {}).get("detail", "")
    if DATASETS.get("price_hu", {}).get("state") != "SOURCE_PARTIAL":
        return []
    return sorted(set(re.findall(r"\d{4}-\d{2}-\d{2}", detail.split(" - ")[0])))


def fill_omitted_price_days(entsoe_adapter):
    """Read the official ENTSO-E A44 curve for days the HUPX feed omitted.

    Stored ONLY in entsoe_day_ahead_prices.csv, a separate labelled dataset. The
    HUPX series (hupx_day_ahead_prices.csv) is never filled from it.
    """
    days = omitted_price_days()
    if not days:
        return
    ok, missing = [], []
    for day in days:
        rows = entsoe_adapter.update_day_ahead_prices(day=dt.date.fromisoformat(day))
        (ok if rows else missing).append(day)
    state = "OK" if not missing else ("SOURCE_PARTIAL" if ok else "SOURCE_PARTIAL")
    record("entsoe_a44_gapfill", state,
           detail="ENTSO-E A44 read for HUPX-omitted days. Filled in the separate ENTSO-E dataset: "
                  + (",".join(ok) or "none") + ". Not available from ENTSO-E: " + (",".join(missing) or "none")
                  + ". HUPX series left unfilled.")


def run_entsoe_batch():
    """Daily ENTSO-E batch: previous-day load forecast, official A44 day-ahead
    prices and the rolling 31-day physical-flow history. Safe to run
    concurrently with update_hun.main(): disjoint hosts, sessions, dataset
    keys and CSV files."""
    import entsoe_adapter
    import physical_flow_history
    entsoe_adapter.update_entsoe(include_physical=False)
    entsoe_adapter.update_day_ahead_prices()
    status = physical_flow_history.update()
    record("entsoe_physical_flows_31d", status["state"],
           detail="31 Budapest delivery days x 14 directions; see physical_flows_status.json",
           row_count=status["row_count"],
           historical_incomplete_pairs=len(status["historical_incomplete_pairs"]),
           request_failures=len(status["request_failures"]))


def main():
    import update_hun
    import entsoe_adapter
    from zoneinfo import ZoneInfo
    now_budapest = dt.datetime.now(ZoneInfo("Europe/Budapest"))
    token = os.environ.get("ENTSOE_API_TOKEN", "").strip()
    force = os.environ.get("ENTSOE_FORCE_DAILY", "").strip().lower() in ("1", "true", "yes")

    if not token:
        update_hun.main()
        for name in ENTSOE_DATASETS:
            record(name, "SKIPPED", detail="ENTSOE_API_TOKEN repository secret not set; "
                                           "token-free feeds unaffected. Set it once under "
                                           "Settings > Secrets and variables > Actions.")
    elif force or entsoe_due_by_schedule(now_budapest):
        # Batch due: run it concurrently with the token-free update. An
        # exception on either side is re-raised after both join, so the
        # workflow step fails exactly as before and saved data is retained.
        errors = {}

        def run(name, func):
            try:
                func()
            except Exception as exc:
                errors[name] = exc

        threads = [threading.Thread(target=run, args=("update_hun", update_hun.main)),
                   threading.Thread(target=run, args=("entsoe_batch", run_entsoe_batch))]
        for thread in threads:
            thread.start()
        for thread in threads:
            thread.join()
        for exc in errors.values():
            raise exc
        # Needs this run's fresh price_hu record, so only after both sides join.
        fill_omitted_price_days(entsoe_adapter)
    else:
        update_hun.main()
        if entsoe_due(now_budapest):   # only the HUPX-omitted-day trigger can still fire
            run_entsoe_batch()
            fill_omitted_price_days(entsoe_adapter)
        else:
            for name in ENTSOE_DATASETS:
                record(name, "SKIPPED", detail="Daily ENTSO-E batch already succeeded after today's "
                                               "03:30 Europe/Budapest boundary; next run tomorrow.")
    write_status("status.json")


if __name__ == "__main__":
    main()
