"""Single orchestrated entrypoint for the Hungary update.

One workflow, one schedule (every 30 minutes) and one manual "Run workflow"
button. Each run always refreshes the token-free feeds. The ENTSO-E token feeds
(previous-day load forecast, official A44 prices, and the rolling 31-day
physical-flow history) run once per day, after the previous Budapest day is
complete, and only when the ENTSOE_API_TOKEN secret is configured. Without the
secret they record SKIPPED (a warning, not a red run); with it, any failure
turns the run red through the workflow gate.
"""
import datetime as dt
import json
import os
import sys
from pathlib import Path

from pipeline_status import DATASETS, record, write_status

ENTSOE_DATASETS = ("entsoe_load_forecast", "entsoe_prices", "entsoe_physical_flows_31d")
DAILY_AFTER_LOCAL_HOUR = 3   # run the daily ENTSO-E batch from 03:30 Budapest
DAILY_AFTER_LOCAL_MINUTE = 30


def entsoe_due(now_budapest, status_path="status.json"):
    scheduled_today = now_budapest.replace(hour=DAILY_AFTER_LOCAL_HOUR, minute=DAILY_AFTER_LOCAL_MINUTE,
                                           second=0, microsecond=0)
    if now_budapest < scheduled_today:
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
    return last_dt < scheduled_today.astimezone(dt.timezone.utc)


def main():
    import update_hun
    update_hun.main()

    import entsoe_adapter
    import physical_flow_history
    from zoneinfo import ZoneInfo
    now_budapest = dt.datetime.now(ZoneInfo("Europe/Budapest"))
    token = os.environ.get("ENTSOE_API_TOKEN", "").strip()
    if not token:
        for name in ENTSOE_DATASETS:
            record(name, "SKIPPED", detail="ENTSOE_API_TOKEN repository secret not set; "
                                           "token-free feeds unaffected. Set it once under "
                                           "Settings > Secrets and variables > Actions.")
    elif not entsoe_due(now_budapest):
        for name in ENTSOE_DATASETS:
            record(name, "SKIPPED", detail="Daily ENTSO-E batch already succeeded after today's "
                                           "03:30 Europe/Budapest boundary; next run tomorrow.")
    else:
        entsoe_adapter.update_entsoe(include_physical=False)
        entsoe_adapter.update_day_ahead_prices()
        status = physical_flow_history.update()
        record("entsoe_physical_flows_31d", status["state"],
               detail="31 Budapest delivery days x 14 directions; see physical_flows_status.json",
               row_count=status["row_count"],
               historical_incomplete_pairs=len(status["historical_incomplete_pairs"]),
               request_failures=len(status["request_failures"]))
    write_status("status.json")


if __name__ == "__main__":
    main()
