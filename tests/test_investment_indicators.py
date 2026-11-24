"""Tests for investment_indicators: solar capture ratio and BESS spread computed
from synthetic complete/incomplete price and generation days."""
import csv
import datetime as dt
import io
import sys
import tempfile
import unittest
from pathlib import Path
from zoneinfo import ZoneInfo

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import investment_indicators as inv
from pipeline_status import DATASETS

TZ = ZoneInfo("Europe/Budapest")
UTC = dt.timezone.utc
QH = dt.timedelta(minutes=15)


def make_day_rows(day, price_fn, solar_fn=None):
    start = dt.datetime.combine(day, dt.time(), tzinfo=TZ).astimezone(UTC)
    n = inv._expected_intervals(day)
    price_rows = [{"Delivery_Date": str(day),
                   "Interval_Start_UTC": (start + QH * i).isoformat(),
                   "Price_EUR_MWh": str(price_fn(i))} for i in range(n)]
    solar_rows = [{"Interval_Start_UTC": (start + QH * i).isoformat(),
                   "Solar_PV_Measured_MW": str(solar_fn(i) if solar_fn else 0)} for i in range(n)]
    return price_rows, solar_rows


def write_csv(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=header, lineterminator="\n")
        w.writeheader()
        w.writerows(rows)


class InvestmentIndicatorsTest(unittest.TestCase):
    def setUp(self):
        DATASETS.clear()
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.addCleanup(self.tmp.cleanup)
        self.addCleanup(DATASETS.clear)

    def _setup(self, price_rows, solar_rows):
        write_csv(self.dir / inv.PRICE_CSV,
                  ["Market", "Country", "Product", "Delivery_Date", "Interval_Start_UTC",
                   "Interval_End_UTC", "Price_EUR_MWh", "Source", "Attribution", "License", "Fetched_At"],
                  [{**{k: "" for k in ("Market", "Country", "Product", "Interval_End_UTC",
                                       "Source", "Attribution", "License", "Fetched_At")}, **r}
                   for r in price_rows])
        write_csv(self.dir / inv.GEN_CSV,
                  ["Interval_Start_UTC", "Interval_End_UTC", "Interval_Start_Local", "Resolution",
                   "Solar_PV_Measured_MW"],
                  [{**{"Interval_End_UTC": "", "Interval_Start_Local": "", "Resolution": "PT15M"}, **r}
                   for r in solar_rows])

    def _read_out(self, name):
        with open(self.dir / name, newline="", encoding="utf-8") as f:
            return list(csv.DictReader(f))

    def test_capture_ratio_known_values(self):
        day = dt.date(2026, 11, 20)   # 96 intervals
        price_rows, solar_rows = make_day_rows(
            day, lambda i: 50.0 if 40 <= i < 48 else 100.0,
            lambda i: 100.0 if 40 <= i < 48 else 0.0)
        self._setup(price_rows, solar_rows)
        inv.update_solar_capture(self.dir)
        rows = self._read_out(inv.SOLAR_CSV)
        self.assertEqual(len(rows), 1)
        r = rows[0]
        day_avg = (88 * 100 + 8 * 50) / 96
        self.assertEqual(r["Day_Avg_Price_EUR_MWh"], f"{day_avg:.2f}")
        self.assertEqual(r["Solar_Weighted_Avg_Price_EUR_MWh"], "50.0000")
        self.assertEqual(r["Solar_Capture_Ratio"], f"{50 / day_avg:.4f}")
        self.assertEqual(r["Solar_Generation_MWh"], "200.0")
        self.assertEqual(DATASETS["solar_capture_daily"]["state"], "OK")

    def test_capture_ratio_blank_when_day_avg_not_positive(self):
        day = dt.date(2026, 11, 20)
        price_rows, solar_rows = make_day_rows(day, lambda i: -10.0, lambda i: 100.0)
        self._setup(price_rows, solar_rows)
        inv.update_solar_capture(self.dir)
        r = self._read_out(inv.SOLAR_CSV)[0]
        self.assertEqual(r["Solar_Capture_Ratio"], "")
        self.assertIn("zero/negative", r["Note"])

    def test_capture_ratio_blank_when_no_solar(self):
        day = dt.date(2026, 11, 20)
        price_rows, solar_rows = make_day_rows(day, lambda i: 100.0, lambda i: 0.0)
        self._setup(price_rows, solar_rows)
        inv.update_solar_capture(self.dir)
        r = self._read_out(inv.SOLAR_CSV)[0]
        self.assertEqual(r["Solar_Capture_Ratio"], "")
        self.assertEqual(r["Solar_Weighted_Avg_Price_EUR_MWh"], "")
        self.assertIn("no measured solar", r["Note"])

    def test_incomplete_day_excluded_whole(self):
        day = dt.date(2026, 11, 20)
        price_rows, solar_rows = make_day_rows(day, lambda i: 100.0, lambda i: 100.0)
        price_rows = price_rows[:50]   # incomplete price day
        self._setup(price_rows, solar_rows)
        inv.update_solar_capture(self.dir)
        self.assertEqual(self._read_out(inv.SOLAR_CSV), [])
        inv.update_bess_spread(self.dir)
        self.assertEqual(self._read_out(inv.BESS_CSV), [])

    def test_dst_hundred_interval_day_computes(self):
        day = dt.date(2026, 10, 25)   # DST switch-off: 100 quarter-hours
        self.assertEqual(inv._expected_intervals(day), 100)
        price_rows, solar_rows = make_day_rows(day, lambda i: 80.0, lambda i: 10.0)
        self._setup(price_rows, solar_rows)
        inv.update_solar_capture(self.dir)
        inv.update_bess_spread(self.dir)
        self.assertEqual(len(self._read_out(inv.SOLAR_CSV)), 1)
        self.assertEqual(len(self._read_out(inv.BESS_CSV)), len(inv.WINDOW_HOURS))

    def test_bess_spread_known_windows_all_durations(self):
        day = dt.date(2026, 11, 20)
        # cheap first 4 hours, expensive last 4 hours, 50 elsewhere
        price_rows, solar_rows = make_day_rows(
            day, lambda i: 10.0 if i < 16 else (110.0 if i >= 80 else 50.0))
        self._setup(price_rows, solar_rows)
        inv.update_bess_spread(self.dir)
        rows = {r["Window_Hours"]: r for r in self._read_out(inv.BESS_CSV)}
        self.assertEqual(set(rows), {"1", "2", "4"})
        for hours, r in rows.items():
            self.assertEqual(r["Spread_EUR_MWh"], "100.00")
            self.assertEqual(r["Charge_Avg_Price_EUR_MWh"], "10.00")
            self.assertEqual(r["Discharge_Avg_Price_EUR_MWh"], "110.00")
            self.assertEqual(r["Charge_Window_Start_Local"], f"{day} 00:00")
            expected_end = (dt.datetime.combine(day, dt.time()) + dt.timedelta(hours=int(hours)))
            self.assertEqual(r["Charge_Window_End_Local"], expected_end.strftime("%Y-%m-%d %H:%M"))

    def test_bess_chronology_enforced_and_negative_reported(self):
        day = dt.date(2026, 11, 20)
        price_rows, solar_rows = make_day_rows(day, lambda i: 200.0 - i)   # strictly decreasing
        self._setup(price_rows, solar_rows)
        inv.update_bess_spread(self.dir)
        for r in self._read_out(inv.BESS_CSV):
            c_start = dt.datetime.strptime(r["Charge_Window_Start_Local"], "%Y-%m-%d %H:%M")
            d_start = dt.datetime.strptime(r["Discharge_Window_Start_Local"], "%Y-%m-%d %H:%M")
            self.assertLessEqual(c_start + dt.timedelta(hours=int(r["Window_Hours"])), d_start)
            self.assertLess(float(r["Spread_EUR_MWh"]), 0)
            self.assertIn("no profitable", r["Note"])

    def test_bess_joint_optimum_not_greedy_cheapest(self):
        # globally cheapest charge window (00:00-02:00, avg 10) is followed by a discharge
        # at 20:00 (avg 100) -> spread 90; the optimal pair charges at 04:00 (avg 20) and
        # discharges at 18:00 (avg 130) -> spread 110. Greedy cheapest-first would miss it.
        day = dt.date(2026, 11, 20)

        def price(i):
            if 16 <= i < 24:
                return 20.0      # moderate early charge, precedes the real peak
            if 72 <= i < 80:
                return 130.0     # best peak, ends 20:00
            if 80 <= i < 88:
                return 10.0      # globally cheapest window, but after the peak
            return 60.0
        price_rows, solar_rows = make_day_rows(day, price)
        self._setup(price_rows, solar_rows)
        inv.update_bess_spread(self.dir)
        rows = {r["Window_Hours"]: r for r in self._read_out(inv.BESS_CSV)}
        r = rows["2"]
        self.assertEqual(r["Charge_Window_Start_Local"], f"{day} 04:00")
        self.assertEqual(r["Discharge_Window_Start_Local"], f"{day} 18:00")
        self.assertEqual(r["Spread_EUR_MWh"], "110.00")

    def test_bess_matches_brute_force_oracle(self):
        import random
        rng = random.Random(20261124)
        for trial in range(30):
            day = dt.date(2026, 11, 20) + dt.timedelta(days=trial % 3)
            vals = [rng.uniform(-30, 250) for _ in range(96)]
            price_rows, solar_rows = make_day_rows(day, lambda i: vals[i])
            self._setup(price_rows, solar_rows)
            inv.update_bess_spread(self.dir)
            got = {r["Window_Hours"]: r for r in self._read_out(inv.BESS_CSV)
                   if r["Delivery_Date"] == str(day)}
            for hours in (1, 2, 4):
                wq = hours * 4
                avgs = [sum(vals[i:i + wq]) / wq for i in range(96 - wq + 1)]
                best, bi, bj = None, None, None
                for j in range(wq, len(avgs)):
                    for i in range(0, j - wq + 1):
                        s = avgs[j] - avgs[i]
                        if best is None or s > best + 1e-12:
                            best, bi, bj = s, i, j
                r = got[str(hours)]
                self.assertAlmostEqual(float(r["Spread_EUR_MWh"]), best, places=2)
                self.assertEqual(r["Charge_Window_Start_Local"],
                                 (dt.datetime.combine(day, dt.time()) + dt.timedelta(minutes=15 * bi)).strftime("%Y-%m-%d %H:%M"))
                self.assertEqual(r["Discharge_Window_Start_Local"],
                                 (dt.datetime.combine(day, dt.time()) + dt.timedelta(minutes=15 * bj)).strftime("%Y-%m-%d %H:%M"))
                # charge window ends no later than discharge starts
                self.assertLessEqual(bi + wq, bj)

    def test_bess_negative_prices_and_ties(self):
        day = dt.date(2026, 11, 20)
        # negative overnight plateau (tie across 8 quarter-hours), flat peak plateau
        price_rows, solar_rows = make_day_rows(
            day, lambda i: -20.0 if 4 <= i < 12 else (150.0 if 72 <= i < 80 else 40.0))
        self._setup(price_rows, solar_rows)
        inv.update_bess_spread(self.dir)
        rows = {r["Window_Hours"]: r for r in self._read_out(inv.BESS_CSV)}
        r = rows["2"]
        self.assertEqual(r["Charge_Avg_Price_EUR_MWh"], "-20.00")
        self.assertEqual(r["Discharge_Avg_Price_EUR_MWh"], "150.00")
        self.assertEqual(r["Spread_EUR_MWh"], "170.00")
        # ties resolve to the earliest optimal window
        self.assertEqual(r["Charge_Window_Start_Local"], f"{day} 01:00")
        self.assertEqual(r["Discharge_Window_Start_Local"], f"{day} 18:00")

    def test_existing_history_rows_retained(self):
        day = dt.date(2026, 11, 20)
        price_rows, solar_rows = make_day_rows(day, lambda i: 100.0, lambda i: 100.0)
        self._setup(price_rows, solar_rows)
        old = [{"Delivery_Date": "2026-10-01", "Day_Avg_Price_EUR_MWh": "70.00",
                "Solar_Weighted_Avg_Price_EUR_MWh": "60.0000", "Solar_Capture_Ratio": "0.8571",
                "Solar_Generation_MWh": "300.0", "Note": ""}]
        write_csv(self.dir / inv.SOLAR_CSV, inv.SOLAR_HEADER, old)
        inv.update_solar_capture(self.dir)
        days = [r["Delivery_Date"] for r in self._read_out(inv.SOLAR_CSV)]
        self.assertEqual(days, ["2026-10-01", "2026-11-20"])


if __name__ == "__main__":
    unittest.main()
