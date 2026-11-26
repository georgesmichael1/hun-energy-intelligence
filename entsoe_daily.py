"""Daily ENTSO-E job for Hungary (previous Budapest day, token feeds).

Mirrors the Bosnia daily ENTSO-E workflow: load forecast (A65) and the official
A44 day-ahead curve for yesterday, plus the A44 read for any day the HUPX feed
omitted (stored only in the separate entsoe_day_ahead_prices.csv, never merged
into the HUPX series). Physical flows have their own job (physical_flow_history.py).

State goes to entsoe_status.json, owned by this job, so the 30-minute job's
status.json is never touched here. Any dataset that is not OK turns the job red;
SOURCE_PARTIAL and STATIC_REFERENCE stay warnings.
"""
import datetime as dt
import json
import re
import sys
from pathlib import Path

from pipeline_status import DATASETS, record, write_status

ENTSOE_STATUS = "entsoe_status.json"


def omitted_price_days(status_path="status.json"):
    """Delivery days the HUPX feed omitted, read from the 30-minute job's status.json."""
    try:
        old = json.loads(Path(status_path).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    item = old.get("dataset_details", {}).get("price_hu", {}) or {}
    if item.get("state") != "SOURCE_PARTIAL":
        return []
    detail = item.get("detail", "")
    return sorted(set(re.findall(r"\d{4}-\d{2}-\d{2}", detail.split(" - ")[0])))


def fill_omitted_price_days(entsoe_adapter, status_path="status.json"):
    days = omitted_price_days(status_path)
    if not days:
        return
    ok, missing = [], []
    for day in days:
        rows = entsoe_adapter.update_day_ahead_prices(day=dt.date.fromisoformat(day))
        (ok if rows else missing).append(day)
    record("entsoe_a44_gapfill", "OK" if not missing else "SOURCE_PARTIAL",
           detail="ENTSO-E A44 read for HUPX-omitted days. Filled in the separate ENTSO-E dataset: "
                  + (",".join(ok) or "none") + ". Not available from ENTSO-E: " + (",".join(missing) or "none")
                  + ". HUPX series left unfilled.")


def main():
    import entsoe_adapter
    entsoe_adapter.update_entsoe(include_physical=False)
    entsoe_adapter.update_day_ahead_prices()
    fill_omitted_price_days(entsoe_adapter)
    write_status(ENTSOE_STATUS)
    return 0


if __name__ == "__main__":
    main()
    from quality_gate import issues
    status = json.loads(Path(ENTSOE_STATUS).read_text(encoding="utf-8"))
    problems = issues(status)
    for p in problems:
        print(f"::error::{p}")
    sys.exit(1 if problems else 0)
