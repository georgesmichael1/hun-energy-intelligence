"""Tests for run_scheduled orchestration: the token-free update and the daily
ENTSO-E batch run concurrently when due, the once-a-day gate is preserved,
ENTSOE_FORCE_DAILY overrides it for manual runs, a batch failure still fails
the run, and the HUPX-omitted-day gap fill sees this run's price_hu record."""
import json
import os
import sys
import tempfile
import time
import types
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import run_scheduled
from pipeline_status import DATASETS


class _FakeUpdateHun(types.SimpleNamespace):
    def __init__(self, marks, sleep_s=0.0, price_state="OK", price_detail="mock"):
        super().__init__()
        def main():
            time.sleep(sleep_s)
            marks.append(("update_hun", time.monotonic()))
            run_scheduled.record("price_hu", price_state, detail=price_detail)
            run_scheduled.record("fx", "OK", detail="mock")
        self.main = main


class _FakeEntsoe(types.SimpleNamespace):
    def __init__(self, marks, sleep_s=0.0):
        super().__init__()
        self.gapfill_days = []
        def update_entsoe(include_physical=False):
            time.sleep(sleep_s)
            marks.append(("update_entsoe", time.monotonic()))
            run_scheduled.record("entsoe_load_forecast", "OK", detail="mock")
        def update_day_ahead_prices(day=None):
            if day is not None:
                self.gapfill_days.append(str(day))
                return ["row"]
            marks.append(("update_day_ahead_prices", time.monotonic()))
            run_scheduled.record("entsoe_prices", "OK", detail="mock")
            return ["row"]
        self.update_entsoe = update_entsoe
        self.update_day_ahead_prices = update_day_ahead_prices


class _FakePhysical(types.SimpleNamespace):
    def __init__(self, marks, sleep_s=0.0, error=None):
        super().__init__()
        def update():
            time.sleep(sleep_s)
            marks.append(("physical_flows", time.monotonic()))
            if error:
                raise error
            return {"state": "OK", "row_count": 10,
                    "historical_incomplete_pairs": [], "request_failures": []}
        self.update = update


class RunScheduledTest(unittest.TestCase):
    def setUp(self):
        DATASETS.clear()
        self.tmp = tempfile.TemporaryDirectory()
        self.old_cwd = os.getcwd()
        os.chdir(self.tmp.name)
        self.marks = []
        self.fake_hun = _FakeUpdateHun(self.marks)
        self.fake_entsoe = _FakeEntsoe(self.marks)
        self.fake_physical = _FakePhysical(self.marks)
        self.modules = patch.dict(sys.modules, {
            "update_hun": self.fake_hun,
            "entsoe_adapter": self.fake_entsoe,
            "physical_flow_history": self.fake_physical,
        })
        self.modules.start()
        self.addCleanup(self.modules.stop)
        self.env = patch.dict(os.environ, {"ENTSOE_API_TOKEN": "token", "ENTSOE_FORCE_DAILY": ""})
        self.env.start()
        self.addCleanup(self.env.stop)
        self.addCleanup(lambda: os.chdir(self.old_cwd))
        self.addCleanup(self.tmp.cleanup)
        self.addCleanup(DATASETS.clear)

    def _write_old_status(self, last_success_at, physical_state="OK", gapfill=None):
        details = {"entsoe_load_forecast": {"state": "OK", "last_success_at": last_success_at},
                   "entsoe_physical_flows_31d": {"state": physical_state}}
        if gapfill:
            details["entsoe_a44_gapfill"] = {"state": "OK"}
        Path("status.json").write_text(json.dumps({"dataset_details": details}), encoding="utf-8")

    def test_concurrent_when_due_and_all_keys_in_status(self):
        self._write_old_status("2020-01-01T00:00:00+00:00")
        self.fake_hun.main.__defaults__ = None
        self.fake_hun = _FakeUpdateHun(self.marks, sleep_s=0.4)
        self.fake_entsoe = _FakeEntsoe(self.marks, sleep_s=0.4)
        with patch.dict(sys.modules, {"update_hun": self.fake_hun, "entsoe_adapter": self.fake_entsoe}):
            start = time.monotonic()
            run_scheduled.main()
            elapsed = time.monotonic() - start
        # Serial would take at least 0.8s; concurrency must finish well under that.
        self.assertLess(elapsed, 0.75)
        for key in ("price_hu", "fx", "entsoe_load_forecast", "entsoe_prices", "entsoe_physical_flows_31d"):
            self.assertEqual(DATASETS[key]["state"], "OK", key)
        on_disk = json.loads(Path("status.json").read_text(encoding="utf-8"))
        for key in ("price_hu", "fx", "entsoe_load_forecast", "entsoe_prices", "entsoe_physical_flows_31d"):
            self.assertIn(key, on_disk["dataset_details"], key)

    def test_gate_skip_when_already_succeeded_today(self):
        self._write_old_status(run_scheduled.dt.datetime.now(run_scheduled.dt.timezone.utc)
                               .isoformat(timespec="seconds"))
        def forbidden(**kwargs):
            raise AssertionError("ENTSO-E batch must not run")
        self.fake_entsoe.update_entsoe = forbidden
        run_scheduled.main()
        for key in run_scheduled.ENTSOE_DATASETS:
            self.assertEqual(DATASETS[key]["state"], "SKIPPED", key)
        self.assertEqual(DATASETS["price_hu"]["state"], "OK")

    def test_force_env_overrides_gate(self):
        self._write_old_status(run_scheduled.dt.datetime.now(run_scheduled.dt.timezone.utc)
                               .isoformat(timespec="seconds"))
        os.environ["ENTSOE_FORCE_DAILY"] = "1"
        run_scheduled.main()
        self.assertEqual(DATASETS["entsoe_load_forecast"]["state"], "OK")
        self.assertEqual(DATASETS["entsoe_physical_flows_31d"]["state"], "OK")

    def test_batch_failure_propagates_and_token_free_side_kept(self):
        self._write_old_status("2020-01-01T00:00:00+00:00")
        self.fake_physical = _FakePhysical(self.marks, error=RuntimeError("boom"))
        with patch.dict(sys.modules, {"physical_flow_history": self.fake_physical}):
            with self.assertRaises(RuntimeError):
                run_scheduled.main()
        self.assertEqual(DATASETS["price_hu"]["state"], "OK")
        self.assertEqual(DATASETS["entsoe_load_forecast"]["state"], "OK")
        self.assertNotIn("entsoe_physical_flows_31d", DATASETS)

    def test_omitted_day_trigger_runs_batch_and_gapfill_after_update_hun(self):
        # Not due by schedule, but this run's price_hu reports a HUPX-omitted day.
        self._write_old_status(run_scheduled.dt.datetime.now(run_scheduled.dt.timezone.utc)
                               .isoformat(timespec="seconds"))
        self.fake_hun = _FakeUpdateHun(self.marks, price_state="SOURCE_PARTIAL",
                                       price_detail="2026-11-20 - omitted by source")
        with patch.dict(sys.modules, {"update_hun": self.fake_hun}):
            run_scheduled.main()
        self.assertEqual(DATASETS["entsoe_load_forecast"]["state"], "OK")
        self.assertEqual(self.fake_entsoe.gapfill_days, ["2026-11-20"])
        self.assertEqual(DATASETS["entsoe_a44_gapfill"]["state"], "OK")

    def test_no_token_records_skipped(self):
        os.environ["ENTSOE_API_TOKEN"] = ""
        run_scheduled.main()
        for key in run_scheduled.ENTSOE_DATASETS:
            self.assertEqual(DATASETS[key]["state"], "SKIPPED", key)
        self.assertEqual(DATASETS["price_hu"]["state"], "OK")


if __name__ == "__main__":
    unittest.main()
