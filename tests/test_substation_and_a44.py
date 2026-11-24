import csv
import datetime as dt
import tempfile
import unittest
from pathlib import Path

import entsoe_adapter
import run_scheduled
import substation_capacity_feed as feed
from pipeline_status import DATASETS

FIX = Path(__file__).parent / "fixtures" / "eon_capacity.txt"


class SubstationParser(unittest.TestCase):
    def test_parses_published_table(self):
        rows = feed.parse_text(FIX.read_text(encoding="utf-8"), "t")
        self.assertEqual(len(rows), 51)
        first = rows[0]
        self.assertEqual((first["Substation"], first["Voltage_kV"], first["Releasable_Offtake_MVA"],
                          first["Probable_5y_Offtake_MVA"]), ("Albertfalva", 11, 6, 2))
        zeros = [r for r in rows if r["Substation"] == "Biatorbágy"][0]
        self.assertEqual(zeros["Releasable_Offtake_MVA"], 0)   # a published 0 stays 0
        self.assertTrue(any(r["Substation"] == "Vác" and r["Voltage_kV"] == 35 for r in rows))
        self.assertTrue(any(r["Substation"] == "Vác" and r["Voltage_kV"] == 22 for r in rows))

    def test_rejects_layout_change_and_duplicates(self):
        with self.assertRaises(ValueError):
            feed.parse_text("Grid hosting capacity - EEL\nnothing here\n", "t")
        text = FIX.read_text(encoding="utf-8")
        dup = text.replace("Albertfalva                               11                                 6                                2",
                           "Albertfalva                               11                                 6                                2\n"
                           "Albertfalva                               11                                 6                                2")
        with self.assertRaises(ValueError):
            feed.parse_text(dup, "t")
        bad = text.replace("Albertfalva                               11 ", "Albertfalva                               17 ")
        with self.assertRaises(ValueError):
            feed.parse_text(bad, "t")


class OmittedDayFill(unittest.TestCase):
    def setUp(self):
        DATASETS.clear()

    def test_omitted_days_only_from_source_partial(self):
        DATASETS["price_hu"] = {"state": "SOURCE_PARTIAL",
                                "detail": "Source returned no valid prices for 2026-11-21,2026-11-18 - day omitted whole"}
        self.assertEqual(run_scheduled.omitted_price_days(), ["2026-11-18", "2026-11-21"])
        DATASETS["price_hu"] = {"state": "OK", "detail": "x 2026-11-21"}
        self.assertEqual(run_scheduled.omitted_price_days(), [])

    def test_multi_day_a44_persist_keeps_days_sorted_and_separate(self):
        def day_rows(day):
            base = dt.datetime.fromisoformat(day + "T00:00:00+00:00")
            return [{c: "" for c in entsoe_adapter.PRICE_COLUMNS} | {
                "Delivery_Date": day, "Interval_Start_UTC": (base + dt.timedelta(minutes=15 * i)).isoformat(),
                "Price_EUR_MWh": 100 + i} for i in range(4)]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "p.csv"
            for day in ("2026-11-23", "2026-11-21", "2026-11-23"):   # fill out of order, then refresh latest
                entsoe_adapter.persist_prices(path, day_rows(day), "2026-10-25")
            with path.open(newline="", encoding="utf-8") as f:
                rows = list(csv.DictReader(f))
        self.assertEqual([r["Delivery_Date"] for r in rows], ["2026-11-21"] * 4 + ["2026-11-23"] * 4)
        self.assertEqual([r["Interval_Start_UTC"] for r in rows], sorted(r["Interval_Start_UTC"] for r in rows))
        self.assertEqual(len(rows), 8)   # refresh replaced, did not duplicate


if __name__ == "__main__":
    unittest.main()


class ConnectionRequests(unittest.TestCase):
    @staticmethod
    def _book(n, dup=False, layout_ok=True):
        import io
        import openpyxl
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append([])
        head = [None] * 28
        head[0] = "Hálózati engedélyes" if layout_ok else "x"
        head[13] = "Alállomás/kapcsolóállomás megnevezése"
        ws.append(head)
        ws.append([])
        ws.append([None, None, "SOLARPHOTOVO", "WINDONSHORE", "BATTERYSTRG"])
        for i in range(n):
            ws.append(["E.ON Zrt.\xa0", "ID-%d" % (0 if dup and i == 1 else i), 1, 0, 2.5, 0, 0, 0, 0, 0,
                       "Új igény", 3.5, 0, "Alfa", 22, "nem", ""])
        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    def test_parse_values_and_blank_handling(self):
        import mavir_connection_requests_feed as m
        rows = m.parse(self._book(120), "t")
        self.assertEqual(len(rows), 120)
        self.assertEqual((rows[0]["Licensee"], rows[0]["Solar_MVA"], rows[0]["Storage_MVA"], rows[0]["Voltage_kV"],
                          rows[0]["Rejected"]), ("E.ON Zrt.", 1, 2.5, "22", "nem"))

    def test_rejects_duplicates_layout_and_short(self):
        import mavir_connection_requests_feed as m
        for book in (self._book(120, dup=True), self._book(120, layout_ok=False), self._book(10)):
            with self.assertRaises(ValueError):
                m.parse(book, "t")


class InstalledCapacity(unittest.TestCase):
    def test_parse_fixture(self):
        import mavir_installed_capacity_feed as m
        text = (Path(__file__).parent / "fixtures" / "mavir_bt.txt").read_text(encoding="utf-8")
        rows = m.parse_text(text, dt.date(2026, 11, 1), "u", "s")
        latest = {r["Fuel_Group"]: r["Gross_Installed_MW"] for r in rows if r["Period"] == "2026-11-01"}
        self.assertEqual(latest["Solar_PV_Industrial"], 5046.7)
        self.assertEqual(latest["Battery_Storage_Market_plus_MAVIR"], 1074.9)
        self.assertEqual(latest["TOTAL_excl_HMKE_SCTE"], 13291.1)
        self.assertEqual({r["Period"] for r in rows if r["Fuel_Group"] == "Wind"}.__len__(), 12)

    def test_rejects_changed_layout_and_bad_sum(self):
        import mavir_installed_capacity_feed as m
        text = (Path(__file__).parent / "fixtures" / "mavir_bt.txt").read_text(encoding="utf-8")
        with self.assertRaises(ValueError):
            m.parse_text("nothing", dt.date(2026, 11, 1), "u", "s")
        with self.assertRaises(ValueError):
            m.parse_text(text.replace("5046,7", "6046,7"), dt.date(2026, 11, 1), "u", "s")
