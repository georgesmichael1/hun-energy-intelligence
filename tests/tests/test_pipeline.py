import datetime as dt
import unittest
from zoneinfo import ZoneInfo

import entsoe_adapter as e
import physical_flow_history as pfh


class AreaTest(unittest.TestCase):
    def test_eight_verified_eics(self):
        self.assertEqual(e.AREAS["HU"], "10YHU-MAVIR----U")
        self.assertEqual(len(e.AREAS), 8)
        self.assertEqual(set(e.NEIGHBORS), {"AT", "SK", "SI", "RO", "HR", "RS", "UA"})

    def test_fourteen_directions(self):
        self.assertEqual(len(pfh.DIRECTIONS), 14)
        for a, b in pfh.DIRECTIONS:
            self.assertTrue(a == "HU" or b == "HU")

    def test_day_bounds_budapest(self):
        start, end = e.day_bounds(dt.date(2026, 11, 23))
        self.assertEqual(start, dt.datetime(2026, 11, 22, 23, 0, tzinfo=dt.timezone.utc))
        self.assertEqual(end, dt.datetime(2026, 11, 23, 23, 0, tzinfo=dt.timezone.utc))

    def test_no_token_in_module_constants(self):
        import inspect
        for module in (e, pfh):
            for name, value in inspect.getmembers(module):
                if isinstance(value, str) and len(value) > 20 and name.isupper():
                    self.assertNotRegex(value, r"[0-9a-f]{8}-[0-9a-f]{4}-", name)


if __name__ == "__main__":
    unittest.main()
