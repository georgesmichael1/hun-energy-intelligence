import datetime as dt
import io
import unittest
from zoneinfo import ZoneInfo

import openpyxl

import mavir_adapter as m

TZ = ZoneInfo("Europe/Budapest")


def build_xlsx(headers, rows):
    book = openpyxl.Workbook()
    sheet = book.active
    sheet.append(headers)
    for r in rows:
        sheet.append(r)
    buf = io.BytesIO()
    book.save(buf)
    return buf.getvalue()


def day_export(day, columns, rows_per_day=4):
    spec = m.CHARTS["load"]
    headers = ["Időpont"] + [hu for hu, _ in spec["columns"]]
    start = dt.datetime.combine(day, dt.time(), TZ)
    rows = []
    for i in range(rows_per_day):
        stamp = start + dt.timedelta(minutes=15 * i)
        rows.append([(stamp + dt.timedelta(minutes=15)).strftime("%Y.%m.%d %H:%M:%S %z")] + [100.0 + i] * len(spec["columns"]))
    return build_xlsx(headers, rows), spec


class ExportParseTest(unittest.TestCase):
    def test_valid_partial_day(self):
        day = dt.date(2026, 11, 23)
        content, spec = day_export(day, None)
        start = dt.datetime.combine(day, dt.time(), TZ).astimezone(dt.timezone.utc)
        end = start + dt.timedelta(days=1)
        now = dt.datetime.now(dt.timezone.utc)
        rows, cols = m.parse_export(content, spec["columns"], start, end, now)
        self.assertEqual(len(rows), 4)
        self.assertEqual(rows[0]["Resolution"], "PT15M")
        self.assertEqual(rows[0]["Gross_Actual_System_Load_MW"], 100.0)
        self.assertTrue(rows[0]["Interval_Start_UTC"].endswith("+00:00"))

    def test_header_change_rejected(self):
        day = dt.date(2026, 11, 23)
        spec = m.CHARTS["load"]
        headers = ["Idopont"] + [hu for hu, _ in spec["columns"]]  # renamed first column
        stamp = (dt.datetime.combine(day, dt.time(), TZ) + dt.timedelta(minutes=15)).strftime("%Y.%m.%d %H:%M:%S %z")
        content = build_xlsx(headers, [[stamp] + [1.0] * len(spec["columns"])])
        start = dt.datetime.combine(day, dt.time(), TZ).astimezone(dt.timezone.utc)
        with self.assertRaises(m.MavirError):
            m.parse_export(content, spec["columns"], start, start + dt.timedelta(days=1),
                           dt.datetime.now(dt.timezone.utc))

    def test_duplicate_interval_rejected(self):
        day = dt.date(2026, 11, 23)
        spec = m.CHARTS["load"]
        headers = ["Időpont"] + [hu for hu, _ in spec["columns"]]
        start_local = dt.datetime.combine(day, dt.time(), TZ)
        stamp = start_local.strftime("%Y.%m.%d %H:%M:%S %z")
        content = build_xlsx(headers, [[stamp] + [1.0] * len(spec["columns"])],
                             )
        content2 = build_xlsx(headers, [[stamp] + [1.0] * len(spec["columns"]),
                                        [stamp] + [2.0] * len(spec["columns"])])
        start = start_local.astimezone(dt.timezone.utc)
        with self.assertRaises(m.MavirError):
            m.parse_export(content2, spec["columns"], start, start + dt.timedelta(days=1),
                           dt.datetime.now(dt.timezone.utc))

    def test_future_interval_skipped(self):
        spec = m.CHARTS["load"]
        headers = ["Időpont"] + [hu for hu, _ in spec["columns"]]
        now = dt.datetime.now(dt.timezone.utc)
        past = (now - dt.timedelta(minutes=30)).astimezone(TZ)
        past = past.replace(second=0, microsecond=0)
        past = past - dt.timedelta(minutes=past.minute % 15)
        future = past + dt.timedelta(days=1)
        content = build_xlsx(headers, [
            [(past + dt.timedelta(minutes=15)).strftime("%Y.%m.%d %H:%M:%S %z")] + [1.0] * len(spec["columns"]),
            [(future + dt.timedelta(minutes=15)).strftime("%Y.%m.%d %H:%M:%S %z")] + [2.0] * len(spec["columns"])])
        start = (past - dt.timedelta(days=1)).astimezone(dt.timezone.utc)
        end = (future + dt.timedelta(days=2)).astimezone(dt.timezone.utc)
        rows, _ = m.parse_export(content, spec["columns"], start, end, now)
        self.assertEqual(len(rows), 1)

    def test_empty_cell_stays_empty(self):
        spec = m.CHARTS["load"]
        headers = ["Időpont"] + [hu for hu, _ in spec["columns"]]
        now = dt.datetime.now(dt.timezone.utc)
        past = (now - dt.timedelta(minutes=30)).astimezone(TZ).replace(second=0, microsecond=0)
        past = past - dt.timedelta(minutes=past.minute % 15)
        content = build_xlsx(headers, [
            [(past + dt.timedelta(minutes=15)).strftime("%Y.%m.%d %H:%M:%S %z"), None] + [1.0] * (len(spec["columns"]) - 1)])
        start = (past - dt.timedelta(days=1)).astimezone(dt.timezone.utc)
        end = (past + dt.timedelta(days=1)).astimezone(dt.timezone.utc)
        rows, _ = m.parse_export(content, spec["columns"], start, end, now)
        self.assertEqual(rows[0]["Gross_Certified_Load_Actual_MW"], "")


if __name__ == "__main__":
    unittest.main()
