import threading
import time
import unittest
from unittest.mock import patch

import update_hun
from pipeline_status import DATASETS

TASKS = ["mavir", "brp_register", "substation_capacity", "mavir_connection_requests",
         "mavir_installed_capacity", "eurostat", "population", "price_hu", "fx",
         "current_gdp", "mekh_register", "worldbank"]


class ParallelFetchTest(unittest.TestCase):
    def setUp(self):
        DATASETS.clear()

    def _patched(self, side_effects):
        """Patch every feed entry point. side_effects: name -> callable(session) or exception."""
        def make(name):
            def run(*args, **kwargs):
                effect = side_effects.get(name)
                if isinstance(effect, Exception):
                    raise effect
                if callable(effect):
                    effect(name)
                update_hun.record(name, "OK", detail="mock")
                return None
            return run
        return [
            patch.object(update_hun.mavir_adapter, "update_all", make("mavir")),
            patch.object(update_hun.brp_feed, "update_brp_register", make("brp_register")),
            patch.object(update_hun.substation_capacity_feed, "update_substation_capacity", make("substation_capacity")),
            patch.object(update_hun.mavir_connection_requests_feed, "update", make("mavir_connection_requests")),
            patch.object(update_hun.mavir_installed_capacity_feed, "update_installed_capacity", make("mavir_installed_capacity")),
            patch.object(update_hun.eurostat_feeds, "update_eurostat", make("eurostat")),
            patch.object(update_hun, "fetch_eurostat_population", make("population")),
            patch.object(update_hun.hupx_prices, "update_hupx_prices", make("price_hu")),
            patch.object(update_hun.mnb_fx, "update_mnb_fx", make("fx")),
            patch.object(update_hun.imf_gdp, "update_current_gdp", make("current_gdp")),
            patch.object(update_hun.mekh_register_feed, "update_mekh_register", make("mekh_register")),
            patch.object(update_hun, "fetch_worldbank_hun_macro", make("worldbank")),
            patch.object(update_hun, "make_session", lambda: _DummySession()),
            patch.object(update_hun, "sync_master_csv", lambda: None),
            patch.object(update_hun.investment_indicators, "update_solar_capture", make("solar_capture_daily")),
            patch.object(update_hun.investment_indicators, "update_bess_spread", make("bess_spread_daily")),
            patch.object(update_hun.solar_bess_model, "update_solar_bess", make("solar_bess_backtest")),
        ]

    def test_all_datasets_recorded_once_and_failure_preserved(self):
        marks = []
        for patcher in self._patched({"fx": RuntimeError("boom")}):
            patcher.start()
        self.addCleanup(patch.stopall)
        update_hun.main()
        for name in TASKS:
            self.assertIn(name, DATASETS, name)
        self.assertEqual(DATASETS["fx"]["state"], "FAILED")
        self.assertEqual(DATASETS["fx"]["detail"], "RuntimeError")

    def test_same_host_group_never_concurrent(self):
        mavir_group = {"mavir", "brp_register", "substation_capacity",
                       "mavir_connection_requests", "mavir_installed_capacity"}
        active = {"n": 0}
        lock = threading.Lock()
        overlap = []

        def slow(name):
            if name in mavir_group:
                with lock:
                    active["n"] += 1
                    if active["n"] > 1:
                        overlap.append(name)
                time.sleep(0.05)
                with lock:
                    active["n"] -= 1

        for patcher in self._patched({t: slow for t in mavir_group}):
            patcher.start()
        self.addCleanup(patch.stopall)
        update_hun.main()
        self.assertEqual(overlap, [])
        for name in mavir_group:
            self.assertIn(name, DATASETS)

    def test_groups_run_concurrently(self):
        def slow(name):
            time.sleep(0.2)

        for patcher in self._patched({t: slow for t in TASKS}):
            patcher.start()
        self.addCleanup(patch.stopall)
        start = time.monotonic()
        update_hun.main()
        elapsed = time.monotonic() - start
        # 5 MAVIR tasks x 0.2s = 1.0s sequential floor for the longest group;
        # fully sequential would take 12 x 0.2s = 2.4s.
        self.assertLess(elapsed, 2.0)


class _DummySession:
    def close(self):
        pass


if __name__ == "__main__":
    unittest.main()
