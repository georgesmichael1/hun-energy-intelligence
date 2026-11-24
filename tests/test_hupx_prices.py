import datetime as dt
import unittest
from zoneinfo import ZoneInfo

import hupx_prices

TZ = ZoneInfo("Europe/Budapest")


def payload_for(day, interval=900, unit="EUR / MWh", license_info="CC BY 4.0 (x) from Bundesnetzagentur | SMARD.de"):
    start = dt.datetime.combine(day, dt.time(), TZ).astimezone(dt.timezone.utc)
    end = dt.datetime.combine(day + dt.timedelta(days=1), dt.time(), TZ).astimezone(dt.timezone.utc)
    times = list(range(int(start.timestamp()), int(end.timestamp()), interval))
    return {"license_info": license_info, "unit": unit,
            "unix_seconds": times, "price": [100.0 + (i % 5) for i in range(len(times))]}


class HupxPriceTest(unittest.TestCase):
    def test_complete_day(self):
        day = dt.date(2026, 11, 23)
        rows = hupx_prices.parse_hu_day(payload_for(day), day)
        self.assertEqual(len(rows), 96)
        self.assertEqual(rows[0]["Market"], "HU")
        self.assertEqual(rows[0]["License"], hupx_prices.LICENSE_URL)
        self.assertEqual(rows[0]["Interval_End_UTC"], rows[1]["Interval_Start_UTC"])

    def test_dst_spring_forward_day(self):
        day = dt.date(2026, 3, 29)  # 23-hour Budapest day
        rows = hupx_prices.parse_hu_day(payload_for(day), day)
        self.assertEqual(len(rows), 92)

    def test_dst_fall_back_day(self):
        day = dt.date(2026, 10, 25)  # 25-hour Budapest day
        rows = hupx_prices.parse_hu_day(payload_for(day), day)
        self.assertEqual(len(rows), 100)

    def test_license_change_rejected(self):
        with self.assertRaises(ValueError):
            hupx_prices.parse_hu_day(payload_for(dt.date(2026, 11, 23), license_info="proprietary"), dt.date(2026, 11, 23))

    def test_gap_rejected(self):
        p = payload_for(dt.date(2026, 11, 23))
        del p["unix_seconds"][40]
        del p["price"][40]
        with self.assertRaises(ValueError):
            hupx_prices.parse_hu_day(p, dt.date(2026, 11, 23))

    def test_negative_prices_valid(self):
        p = payload_for(dt.date(2026, 11, 23))
        p["price"][10] = -12.5
        rows = hupx_prices.parse_hu_day(p, dt.date(2026, 11, 23))
        self.assertEqual(rows[10]["Price_EUR_MWh"], -12.5)


if __name__ == "__main__":
    unittest.main()
