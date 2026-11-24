"""MAVIR balance responsible party (BRP) register: Hungary's licensed market actors.

Every Hungarian market participant belongs to a balance group (merlegkor)
whose balance responsible party settles imbalances with MAVIR. MAVIR publishes
the current register as an XLSX linked from its scheduling page. A BRP is an imbalance-settlement
party, not automatically an aggregator. The aggregator roster comes from MEKH
(mekh_register_feed.py). Names, EIC codes and validity dates are as published.
"""
import csv
import datetime as dt
import io
import re
import warnings
from pathlib import Path
from urllib.parse import urljoin, urlparse

import openpyxl
import requests
from pipeline_status import atomic_text, record, utc_now

PAGE = "https://www.mavir.hu/web/mavir/menetrendkezeles1"
COLUMNS = ["Balance_Responsible_Party", "EIC_Code", "Valid_From", "Source", "Retrieved_UTC"]
HEADERS = {"User-Agent": (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")}
EXPECTED_HEADERS = ("Mérlegkör-felelős/Balance Responsible Party",
                    "EIC kód/EIC code",
                    "Érvényesség kezdete/ Valid from")


def find_register_url(session):
    response = session.get(PAGE, headers=HEADERS, timeout=(10, 30))
    if response.status_code != 200:
        raise ValueError("HTTP_" + str(response.status_code))
    links = re.findall(r'href="(/documents/[^"]*EIC_\d{6}\.xlsx/[^"]*)"', response.text)
    if not links:
        raise ValueError("No BRP register link on MAVIR scheduling page")
    url = urljoin(PAGE, links[0].replace("&amp;", "&"))
    if urlparse(url).hostname != "www.mavir.hu":
        raise ValueError("Unexpected register host")
    return url


def parse_register(content, stamp, source):
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", UserWarning)
        book = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
    sheets = book.worksheets
    if len(sheets) != 1:
        raise ValueError("Register workbook sheet count changed")
    table = list(sheets[0].values)
    if not table or len(table) < 10:
        raise ValueError("Register workbook too small")
    header = tuple(str(c).strip() if c is not None else "" for c in table[0][:3])
    if header != EXPECTED_HEADERS:
        raise ValueError("Register headers changed")
    rows = []
    seen = set()
    for raw in table[1:]:
        if raw is None or raw[0] is None:
            continue
        name = str(raw[0]).strip()
        eic = str(raw[1]).strip() if raw[1] is not None else ""
        valid = raw[2]
        if not name or not eic:
            raise ValueError("Register row missing name or EIC")
        # Names can repeat (e.g. HUPX holds two registrations); the EIC is unique.
        if eic in seen:
            raise ValueError("Duplicate register EIC")
        seen.add(eic)
        if not re.fullmatch(r"[A-Z0-9\-]{16}", eic):
            raise ValueError("Unexpected EIC format in register")
        if isinstance(valid, dt.datetime):
            valid_from = valid.date().isoformat()
        elif isinstance(valid, dt.date):
            valid_from = valid.isoformat()
        else:
            raise ValueError("Unexpected validity date format in register")
        rows.append({"Balance_Responsible_Party": name, "EIC_Code": eic,
                     "Valid_From": valid_from, "Source": source, "Retrieved_UTC": stamp})
    if len(rows) < 50:
        raise ValueError("Register implausibly short")
    return rows


def update_brp_register(path="balance_responsible_parties.csv", session=None, now=None):
    now = now or dt.datetime.now(dt.timezone.utc)
    stamp = utc_now()
    own = session is None
    session = session or requests.Session()
    try:
        url = find_register_url(session)
        response = session.get(url, headers=HEADERS, timeout=(10, 60))
        if response.status_code != 200 or not response.content[:2] == b"PK":
            raise ValueError("Register download failed or not XLSX")
        rows = parse_register(response.content, stamp, url)
        out = io.StringIO(newline="")
        writer = csv.DictWriter(out, fieldnames=COLUMNS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
        atomic_text(path, out.getvalue())
        record("brp_register", "OK",
               detail="MAVIR balance responsible party register as published; BRPs settle imbalances and are "
                      "not all aggregators or traders - MAVIR accredits aggregators separately; no inferred roles or volumes",
               source=url, publication_page=PAGE, party_count=len(rows))
        return rows
    except Exception as exc:
        detail = str(exc) if isinstance(exc, ValueError) else type(exc).__name__
        record("brp_register", "FAILED", detail="Saved register retained; " + detail, source=PAGE)
        return None
    finally:
        if own:
            session.close()


if __name__ == "__main__":
    update_brp_register()
