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




def hupx_rows_for(day, count=None, status="final"):
    n = count if count is not None else hupx_prices.expected_count(day)
    return [{"BaseloadPrice": 150.0, "DeliveryDay": day.isoformat() + "T00:00:00Z",
             "Price": 100.0 + (i % 5), "ProductQH": str(i + 1), "Region": "HU",
             "Status": status, "Volume": 1000.0} for i in range(n)]


class _FakeResp:
    def __init__(self, payload, status_code=200):
        self._payload = payload
        self.status_code = status_code

    def json(self):
        return self._payload


class _FakeSession:
    def __init__(self, payload):
        self.payload = payload
        self.calls = []

    def get(self, url, **kwargs):
        self.calls.append(url)
        return _FakeResp(self.payload)


class HupxLabsTest(unittest.TestCase):
    def test_complete_day(self):
        day = dt.date(2026, 11, 21)
        rows = hupx_prices.parse_hupx_day(hupx_rows_for(day), day)
        self.assertEqual(len(rows), 96)
        self.assertEqual(rows[0]["Source"], hupx_prices.HUPX_SOURCE)
        self.assertEqual(rows[0]["Attribution"], hupx_prices.HUPX_ATTRIBUTION)
        self.assertEqual(rows[0]["License"], hupx_prices.HUPX_LICENSE_URL)
        self.assertEqual(rows[0]["Interval_End_UTC"], rows[1]["Interval_Start_UTC"])
        start = dt.datetime.combine(day, dt.time(), TZ).astimezone(dt.timezone.utc)
        self.assertEqual(rows[0]["Interval_Start_UTC"], start.isoformat())

    def test_dst_fall_back_day(self):
        day = dt.date(2026, 10, 25)  # 25-hour Budapest day = 100 quarter-hours
        rows = hupx_prices.parse_hupx_day(hupx_rows_for(day), day)
        self.assertEqual(len(rows), 100)

    def test_dst_spring_forward_day(self):
        day = dt.date(2026, 3, 29)  # 23-hour Budapest day = 92 quarter-hours
        rows = hupx_prices.parse_hupx_day(hupx_rows_for(day), day)
        self.assertEqual(len(rows), 92)

    def test_non_final_rejected(self):
        day = dt.date(2026, 11, 21)
        with self.assertRaises(ValueError):
            hupx_prices.parse_hupx_day(hupx_rows_for(day, status="preliminary"), day)

    def test_incomplete_day_rejected(self):
        day = dt.date(2026, 11, 21)
        with self.assertRaises(ValueError):
            hupx_prices.parse_hupx_day(hupx_rows_for(day, count=95), day)

    def test_wrong_region_rejected(self):
        day = dt.date(2026, 11, 21)
        rows = hupx_rows_for(day)
        rows[3] = dict(rows[3], Region="RS")
        with self.assertRaises(ValueError):
            hupx_prices.parse_hupx_day(rows, day)

    def test_fetch_groups_valid_days_and_drops_gapped_ones(self):
        good = dt.date(2026, 11, 20)
        gapped = dt.date(2026, 11, 21)
        payload = {"data": hupx_rows_for(good) + hupx_rows_for(gapped, count=90)}
        session = _FakeSession(payload)
        fresh = hupx_prices.fetch_hupx_days(session, good, gapped)
        self.assertIn(good, fresh)
        self.assertNotIn(gapped, fresh)
        self.assertEqual(len(fresh[good]), 96)


if __name__ == "__main__":
    unittest.main()
