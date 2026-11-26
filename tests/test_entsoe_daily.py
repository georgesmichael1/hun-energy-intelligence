"""Tests for the split Hungary jobs: the daily ENTSO-E job and the shared quality gate."""
import json
import os
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import entsoe_daily
import quality_gate
from pipeline_status import DATASETS


class _FakeEntsoe(types.SimpleNamespace):
    def __init__(self):
        super().__init__()
        self.gapfill_days = []
        def update_entsoe(include_physical=False):
            assert include_physical is False   # physical flows have their own job
            entsoe_daily.record("entsoe_load_forecast", "OK", detail="mock")
        def update_day_ahead_prices(day=None):
            if day is not None:
                self.gapfill_days.append(str(day))
                return ["row"]
            entsoe_daily.record("entsoe_prices", "OK", detail="mock")
            return ["row"]
        self.update_entsoe = update_entsoe
        self.update_day_ahead_prices = update_day_ahead_prices


class DailyJob(unittest.TestCase):
    def setUp(self):
        DATASETS.clear()
        self.tmp = tempfile.TemporaryDirectory()
        self.old = os.getcwd()
        os.chdir(self.tmp.name)
        self.fake = _FakeEntsoe()
        patcher = patch.dict(sys.modules, {"entsoe_adapter": self.fake})
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(lambda: os.chdir(self.old))
        self.addCleanup(self.tmp.cleanup)
        self.addCleanup(DATASETS.clear)

    def test_writes_own_status_file_and_leaves_status_json_alone(self):
        Path("status.json").write_text('{"keep": true}', encoding="utf-8")
        entsoe_daily.main()
        on_disk = json.loads(Path("entsoe_status.json").read_text(encoding="utf-8"))
        self.assertEqual(on_disk["dataset_details"]["entsoe_load_forecast"]["state"], "OK")
        self.assertEqual(on_disk["dataset_details"]["entsoe_prices"]["state"], "OK")
        self.assertEqual(json.loads(Path("status.json").read_text(encoding="utf-8")), {"keep": True})

    def test_gapfill_uses_price_hu_from_status_json(self):
        Path("status.json").write_text(json.dumps({"dataset_details": {"price_hu": {
            "state": "SOURCE_PARTIAL", "detail": "2026-11-20 - omitted by source"}}}), encoding="utf-8")
        entsoe_daily.main()
        self.assertEqual(self.fake.gapfill_days, ["2026-11-20"])
        self.assertEqual(DATASETS["entsoe_a44_gapfill"]["state"], "OK")


class Gate(unittest.TestCase):
    def test_states(self):
        ok = {"dataset_details": {"a": {"state": "OK"}, "b": {"state": "SOURCE_PARTIAL", "detail": "x"},
                                  "c": {"state": "STATIC_REFERENCE"}}}
        self.assertEqual(quality_gate.issues(ok), [])
        bad = {"dataset_details": {"a": {"state": "FAILED", "detail": "x"}, "b": {"state": "SKIPPED"}}}
        self.assertEqual(len(quality_gate.issues(bad)), 2)
        self.assertEqual(quality_gate.issues({}), ["Dataset state missing"])


if __name__ == "__main__":
    unittest.main()
