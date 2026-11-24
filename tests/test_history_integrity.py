"""User-requested integrity tests for the saved-history mechanism (Hungary):
retention of complete saved days, rolling 31-day pruning including DST days,
duplicate rejection, missing days staying empty (never zero), and saved good
data surviving a failed fetch with an honest FAILED status."""
import csv
import datetime as dt
import tempfile
import unittest
from pathlib import Path
from zoneinfo import ZoneInfo

import hupx_prices
from pipeline_status import DATASETS

TZ = ZoneInfo("Europe/Budapest")
UTC = dt.timezone.utc
NOW = dt.datetime(2026, 11, 24, 12, 0, tzinfo=UTC)
TODAY = NOW.astimezone(TZ).date()
OLDEST = TODAY - dt.timedelta(days=30)


def day_series(day, base=100.0, null=False):
    start = dt.datetime.combine(day, dt.time(), TZ).astimezone(UTC)
    end = dt.datetime.combine(day + dt.timedelta(days=1), dt.time(), TZ).astimezone(UTC)
    times = list(range(int(start.timestamp()), int(end.timestamp()), 900))
    prices = [None if null else base + (i % 5) for i in range(len(times))]
    return times, prices


def window_payload(skip=(), null_days=()):
    times, prices = [], []
    day = OLDEST
    while day <= TODAY:
        if day not in skip:
            t, p = day_series(day, null=day in null_days)
            times += t
            prices += p
        day += dt.timedelta(days=1)
    return {"license_info": "CC BY 4.0 (x) from Bundesnetzagentur | SMARD.de",
            "unit": "EUR / MWh", "unix_seconds": times, "price": prices}


class FakeResponse:
    status_code = 200

    def __init__(self, payload):
        self._payload = payload

    def json(self):
        return self._payload


class FakeSession:
    def __init__(self, payload=None, status=200):
        self._payload = payload
        self._status = status

    def get(self, url, timeout=None):
        if self._status != 200:
            r = FakeResponse(None)
            r.status_code = self._status
            return r
        return FakeResponse(self._payload)

    def close(self):
        pass


def read_csv(path):
    with open(path, newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def write_day(path, day, price):
    times, _ = day_series(day)
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.DictWriter(handle, fieldnames=hupx_prices.COLUMNS)
        writer.writeheader()
        for t in times:
            writer.writerow({"Market": "HU", "Country": "Hungary", "Product": "DAY_AHEAD",
                             "Delivery_Date": day.isoformat(),
                             "Interval_Start_UTC": dt.datetime.fromtimestamp(t, UTC).isoformat(),
                             "Interval_End_UTC": dt.datetime.fromtimestamp(t + 900, UTC).isoformat(),
                             "Price_EUR_MWh": price, "Source": "saved", "Attribution": "saved",
                             "License": "saved", "Fetched_At": "2026-11-20T00:00:00Z"})


class HistoryIntegrityTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = str(Path(self.tmp.name) / "hupx.csv")

    def tearDown(self):
        self.tmp.cleanup()
        DATASETS.pop("price_hu", None)

    def update(self, session):
        return hupx_prices.update_hupx_prices(self.path, session=session, now=NOW)

    def test_retains_complete_saved_days_and_refreshes_recent(self):
        old_day = TODAY - dt.timedelta(days=5)
        write_day(self.path, old_day, 777.0)
        rows = self.update(FakeSession(window_payload()))
        by_day = {}
        for r in rows:
            by_day.setdefault(r["Delivery_Date"], []).append(float(r["Price_EUR_MWh"]))
        self.assertEqual(set(by_day[old_day.isoformat()]), {777.0})          # retained
        self.assertTrue(all(p < 200.0 for p in by_day[TODAY.isoformat()]))   # refreshed
        self.assertIn((TODAY - dt.timedelta(days=1)).isoformat(), by_day)    # yesterday refreshed
        self.assertEqual(DATASETS["price_hu"]["state"], "OK")

    def test_refills_incomplete_saved_day(self):
        gap_day = TODAY - dt.timedelta(days=5)
        write_day(self.path, gap_day, 777.0)
        rows = read_csv(self.path)[:10]  # truncate the saved day: incomplete
        with open(self.path, "w", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(handle, fieldnames=hupx_prices.COLUMNS)
            writer.writeheader()
            writer.writerows(rows)
        out = self.update(FakeSession(window_payload()))
        day_rows = [r for r in out if r["Delivery_Date"] == gap_day.isoformat()]
        self.assertEqual(len(day_rows), hupx_prices.expected_count(gap_day))
        self.assertTrue(all(float(r["Price_EUR_MWh"]) < 200.0 for r in day_rows))  # fresh, not saved

    def test_rolling_window_prunes_day_31(self):
        stale = TODAY - dt.timedelta(days=31)
        write_day(self.path, stale, 555.0)
        rows = self.update(FakeSession(window_payload()))
        self.assertNotIn(stale.isoformat(), {r["Delivery_Date"] for r in rows})
        self.assertEqual(min(r["Delivery_Date"] for r in rows), OLDEST.isoformat())

    def test_dst_fall_back_day_complete_in_window(self):
        dst_fall = dt.date(2026, 10, 25)  # 100 intervals inside this window
        rows = self.update(FakeSession(window_payload()))
        day_rows = [r for r in rows if r["Delivery_Date"] == dst_fall.isoformat()]
        self.assertEqual(len(day_rows), 100)

    def test_no_duplicate_intervals(self):
        rows = self.update(FakeSession(window_payload()))
        starts = [r["Interval_Start_UTC"] for r in rows]
        self.assertEqual(len(starts), len(set(starts)))

    def test_missing_day_stays_empty_never_zero(self):
        gap = TODAY - dt.timedelta(days=3)
        rows = self.update(FakeSession(window_payload(null_days=(gap,))))
        self.assertEqual([r for r in rows if r["Delivery_Date"] == gap.isoformat()], [])
        self.assertEqual(DATASETS["price_hu"]["state"], "SOURCE_PARTIAL")
        self.assertIn(gap.isoformat(), DATASETS["price_hu"]["detail"])

    def test_failed_fetch_preserves_saved_and_reports_failed(self):
        write_day(self.path, TODAY - dt.timedelta(days=2), 555.0)
        rows = self.update(FakeSession(status=500))
        self.assertIsNotNone(rows)  # saved history still served
        self.assertTrue(all(float(r["Price_EUR_MWh"]) == 555.0 for r in rows))
        self.assertEqual(DATASETS["price_hu"]["state"], "FAILED")
        self.assertIn("saved", DATASETS["price_hu"]["detail"])

    def test_failed_fetch_with_no_saved_history_is_honest_failure(self):
        rows = self.update(FakeSession(status=500))
        self.assertIsNone(rows)
        self.assertFalse(Path(self.path).exists())
        self.assertEqual(DATASETS["price_hu"]["state"], "FAILED")


if __name__ == "__main__":
    unittest.main()


class FailureClassificationTests(unittest.TestCase):
    def test_categories_separate_source_from_transport(self):
        import physical_flow_history as pf
        import entsoe_adapter as ea
        self.assertEqual(pf.classify_failure(ea.DataError('API acknowledgement, code 999: No matching data'))['category'], 'SOURCE_NO_DATA')
        self.assertEqual(pf.classify_failure(ea.APIError('API HTTP 500'))['category'], 'TRANSPORT')
        self.assertEqual(pf.classify_failure(ea.DataError('Unexpected unit, MW required'))['category'], 'VALIDATION')
