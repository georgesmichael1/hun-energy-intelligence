import datetime as dt
import io
import unittest

import openpyxl

import mnb_fx
import brp_feed as aggregators_feed
import eurostat_feeds


MNB_HTML = """
<html><body><h2>Napi árfolyamok: 24 November 2026</h2>
<table>
<tr><td class="fw-b">EUR</td><td>Euro</td><td>1</td><td>369.57</td></tr>
<tr><td class="fw-b">USD</td><td>US Dollar</td><td>1</td><td>337.97</td></tr>
</table></body></html>
"""


class MnbTest(unittest.TestCase):
    def test_parse(self):
        rows = mnb_fx.parse_page(MNB_HTML, dt.date(2026, 11, 24))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["Rate_Date"], "2026-11-24")
        self.assertEqual(rows[0]["HUF_Per_Unit"], 369.57)
        self.assertEqual(rows[1]["Currency"], "USD")

    def test_stale_date_rejected(self):
        with self.assertRaises(ValueError):
            mnb_fx.parse_page(MNB_HTML, dt.date(2027, 1, 1))

    def test_missing_currency_rejected(self):
        with self.assertRaises(ValueError):
            mnb_fx.parse_page("<html>Napi árfolyamok: 24 November 2026</html>", dt.date(2026, 11, 24))


class AggregatorTest(unittest.TestCase):
    def _book(self, rows):
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.append(list(aggregators_feed.EXPECTED_HEADERS))
        for r in rows:
            sheet.append(r)
        buf = io.BytesIO()
        book.save(buf)
        return buf.getvalue()

    def test_parse(self):
        content = self._book([[f"Trader {i:03d} Kft.", f"15X{i:012d}X",
                               dt.datetime(2018, 6, 1)] for i in range(60)])
        rows = aggregators_feed.parse_register(content, "2026-11-24T00:00:00+00:00", "http://x")
        self.assertEqual(len(rows), 60)
        self.assertEqual(rows[0]["Valid_From"], "2018-06-01")
        self.assertEqual(rows[0]["EIC_Code"], "15X000000000000X")

    def test_short_register_rejected(self):
        content = self._book([["A Kft.", "15X-ALTEOEKER--Y", dt.datetime(2018, 6, 1)]])
        with self.assertRaises(ValueError):
            aggregators_feed.parse_register(content, "s", "http://x")

    def test_bad_eic_rejected(self):
        content = self._book([[f"Trader {i:03d} Kft.", "not-an-eic", dt.datetime(2018, 6, 1)] for i in range(60)])
        with self.assertRaises(ValueError):
            aggregators_feed.parse_register(content, "s", "http://x")


class EurostatTest(unittest.TestCase):
    def test_label_check(self):
        spec = eurostat_feeds.DATASETS["hicp"]
        self.assertIn("prc_hicp_manr", spec["dataset"])
        self.assertEqual(spec["params"]["geo"], "HU")
        for name in ("tariff_household", "tariff_nonhousehold"):
            self.assertEqual(eurostat_feeds.DATASETS[name]["params"]["currency"], "EUR")


if __name__ == "__main__":
    unittest.main()
