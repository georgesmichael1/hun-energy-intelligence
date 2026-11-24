"""MAVIR connection-request publication (2nd publication procedure, closing publication).

One XLSX lists every grid-connection request in the procedure, by licensee, with installed MVA per
technology, requested feed-in and offtake MVA, substation, voltage level and whether it was rejected.
It is a CLOSED PROCEDURE published in September 2024 (file timestamp), not a live queue. It is saved once
as a static reference and is not re-downloaded on every run: the saved CSV is kept and the dataset is
labelled STATIC_REFERENCE. Rows are copied as published; no MW conversion and no inference. MVA is MVA.
"""
import csv
import io
import re
from pathlib import Path

import openpyxl
import requests
from pipeline_status import atomic_text, record, utc_now

URL = ("https://www.mavir.hu/documents/10258/246912995/II_kozzetetel_lezaro_publikacio.xlsx/"
       "47d0431f-88ad-a8c4-44c6-9da75d6a9ee0?t=1726553202714")
PAGE = "https://www.mavir.hu/web/mavir/masodik-kozzeteteli-eljaras"
CSV_NAME = "hun_connection_requests.csv"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"}
TECH = ["Solar_MVA", "Wind_MVA", "Storage_MVA", "Hydro_MVA", "Other_RES_MVA", "Oil_MVA", "Gas_CCGT_MVA", "Gas_Other_MVA"]
COLUMNS = (["Licensee", "Request_ID"] + TECH +
           ["Request_Type", "Requested_Feed_In_MVA", "Requested_Offtake_MVA", "Substation", "Voltage_kV",
            "Rejected", "Rejection_Reason", "Source", "Retrieved_UTC"])


def _clean(v):
    return re.sub(r"\s+", " ", str(v).replace("\xa0", " ")).strip() if v not in (None, "") else ""


def _num(v):
    if v in (None, ""):
        return ""
    if isinstance(v, bool) or not isinstance(v, (int, float)) or v < 0:
        raise ValueError("Unexpected numeric value in connection-request sheet")
    return v


def parse(content, stamp):
    book = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    rows = list(book[book.sheetnames[0]].iter_rows(values_only=True))
    if len(rows) < 100 or _clean(rows[1][0]) != "Hálózati engedélyes" or _clean(rows[3][2]) != "SOLARPHOTOVO" \
            or "Alállomás" not in _clean(rows[1][13]):
        raise ValueError("Connection-request sheet layout changed")
    out, seen = [], set()
    for r in rows[4:]:
        if not r or not _clean(r[0]):
            continue
        rid = _clean(r[1])
        if rid in seen:
            raise ValueError("Duplicate request id")
        seen.add(rid)
        rej = _clean(r[15]).lower()
        if rej not in ("igen", "nem"):
            raise ValueError("Unexpected rejected flag")
        kv = r[14]
        out.append(dict(zip(COLUMNS, [_clean(r[0]), rid] + [_num(r[i]) for i in range(2, 10)] +
                            [_clean(r[10]), _num(r[11]), _num(r[12]), _clean(r[13]),
                             _clean(kv) if kv not in (None, "") else "", rej, _clean(r[16]), URL, stamp])))
    if len(out) < 100:
        raise ValueError("Request list implausibly short")
    return out


def _csv(rows):
    buf = io.StringIO(newline="")
    w = csv.DictWriter(buf, fieldnames=COLUMNS, lineterminator="\n")
    w.writeheader()
    w.writerows(rows)
    return buf.getvalue()


def update(output_dir=".", session=None, force=False):
    path = Path(output_dir) / CSV_NAME
    if path.exists() and not force:
        with path.open(newline="", encoding="utf-8") as f:
            n = sum(1 for _ in csv.DictReader(f))
        record("mavir_connection_requests", "STATIC_REFERENCE", source=URL, publication_page=PAGE, row_count=n,
               detail="MAVIR 2nd publication procedure, closing publication of September 2024 (closed procedure, "
                      "not a live queue). Saved once; not re-downloaded each run.")
        return None
    own = session is None
    session = session or requests.Session()
    try:
        resp = session.get(URL, headers=HEADERS, timeout=(10, 60))
        if resp.status_code != 200 or resp.content[:2] != b"PK":
            raise ValueError("Download failed or not XLSX (HTTP " + str(resp.status_code) + ")")
        rows = parse(resp.content, utc_now())
        atomic_text(str(path), _csv(rows))
        record("mavir_connection_requests", "STATIC_REFERENCE", source=URL, publication_page=PAGE, row_count=len(rows),
               detail="MAVIR 2nd publication procedure, closing publication of September 2024 (closed procedure, "
                      "not a live queue). Saved once; not re-downloaded each run.")
        return rows
    except Exception as exc:
        detail = str(exc) if isinstance(exc, ValueError) else type(exc).__name__
        record("mavir_connection_requests", "FAILED", detail="No CSV saved; " + detail, source=PAGE)
        return None
    finally:
        if own:
            session.close()


if __name__ == "__main__":
    update(force=True)
