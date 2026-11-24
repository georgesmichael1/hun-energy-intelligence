"""MEKH licensed-entity register: Hungary's registered independent aggregators and storage licensees.

The regulator (MEKH) publishes one XLSX with a sheet per licence type. Two sheets matter here:
  Aggregator - entities registered as aggregators (Hungary's real aggregator roster)
  Tarolok    - electricity storage licences with nominal power (MW) and storage capacity (MWh)
Rows are copied as published. No volumes, roles or ownership are inferred. A MAVIR balance
responsible party (brp_feed.py) is a different thing and is never labelled an aggregator here.
"""
import csv
import datetime as dt
import io
from pathlib import Path

import openpyxl
import requests
from pipeline_status import atomic_text, record, utc_now

PAGE = "https://www.mekh.hu/villamosenergia-ipari-engedelyesek-listaja"
URL = "https://www.mekh.hu/download/fix/Villamosenergia_engedelyesek"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"}
AGG_COLUMNS = ["Aggregator", "Registration_Decision_Type", "Registration_Decision_No", "Amendments",
               "Country", "Postcode", "City", "Suspension_Start", "Source", "Retrieved_UTC"]
STO_COLUMNS = ["Licensee", "Licence_Type", "Licence_No", "Issued_Date", "Valid_Until", "Site",
               "Nominal_Power_MW", "Storage_Capacity_MWh", "Revoked_Date", "Source", "Retrieved_UTC"]


def _sheet_rows(book, name, first_header):
    if name not in book.sheetnames:
        raise ValueError("Sheet missing: " + name)
    rows = [r for r in book[name].iter_rows(values_only=True) if r and r[0] is not None]
    for i, r in enumerate(rows):
        if str(r[0]).startswith(first_header):
            return rows[i], rows[i + 1:]
    raise ValueError("Header row not found in sheet " + name)


def _date(v):
    if isinstance(v, dt.datetime):
        return v.date().isoformat()
    return str(v).strip() if v not in (None, "") else ""


def _num(v):
    if v in (None, ""):
        return ""
    if isinstance(v, bool) or not isinstance(v, (int, float)) or v < 0:
        raise ValueError("Unexpected numeric value in storage sheet")
    return v


def parse(content, stamp):
    book = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
    head, data = _sheet_rows(book, "Aggregator", "Engedélyes")
    if "Aggregátor" not in str(data[0][1]) if data else True:
        raise ValueError("Aggregator sheet layout changed")
    agg, seen = [], set()
    for r in data:
        key = (str(r[0]).strip(), str(r[2]).strip())
        if key in seen:
            raise ValueError("Duplicate aggregator registration")
        seen.add(key)
        agg.append(dict(zip(AGG_COLUMNS, [str(r[0]).strip(), str(r[1]).strip(), str(r[2]).strip(), r[3],
                                          str(r[5] or ""), str(r[6] or ""), str(r[7] or ""),
                                          _date(r[11]), URL, stamp])))
    head, data = _sheet_rows(book, "Tarolok", "Engedélyes")
    idx = {str(h): i for i, h in enumerate(head)}
    need = ["Alapengedely Szama", "Hatarozat Kiadas Datuma", "Engedely Hatalya", "Telephey",
            "Névleges teljesítőképesség (MW)", "Névleges Tárolókapacitás (MWh)", "Visszavonas Datuma"]
    if any(n not in idx for n in need):
        raise ValueError("Storage sheet layout changed")
    sto, seen = [], set()
    for r in data:
        no = str(r[idx["Alapengedely Szama"]]).strip()
        key = (str(r[0]).strip(), no, str(r[idx["Telephey"]]).strip(), r[idx["Névleges Tárolókapacitás (MWh)"]])
        if key in seen:
            raise ValueError("Duplicate storage licence")
        seen.add(key)
        sto.append(dict(zip(STO_COLUMNS, [
            str(r[0]).strip(), str(r[1]).strip(), no, _date(r[idx["Hatarozat Kiadas Datuma"]]),
            _date(r[idx["Engedely Hatalya"]]), str(r[idx["Telephey"]] or "").strip(),
            _num(r[idx["Névleges teljesítőképesség (MW)"]]), _num(r[idx["Névleges Tárolókapacitás (MWh)"]]),
            _date(r[idx["Visszavonas Datuma"]]), URL, stamp])))
    if len(agg) < 20 or len(sto) < 20:
        raise ValueError("Register implausibly short")
    return agg, sto


def _csv(columns, rows):
    out = io.StringIO(newline="")
    w = csv.DictWriter(out, fieldnames=columns, lineterminator="\n")
    w.writeheader()
    w.writerows(rows)
    return out.getvalue()


def update_mekh_register(output_dir=".", session=None):
    stamp = utc_now()
    own = session is None
    session = session or requests.Session()
    try:
        resp = session.get(URL, headers=HEADERS, timeout=(10, 90))
        if resp.status_code != 200 or resp.content[:2] != b"PK":
            raise ValueError("Download failed or not XLSX (HTTP " + str(resp.status_code) + ")")
        agg, sto = parse(resp.content, stamp)
        root = Path(output_dir)
        atomic_text(str(root / "mekh_aggregators.csv"), _csv(AGG_COLUMNS, agg))
        atomic_text(str(root / "mekh_storage_licences.csv"), _csv(STO_COLUMNS, sto))
        power = sum(r["Nominal_Power_MW"] for r in sto if r["Nominal_Power_MW"] != "" and not r["Revoked_Date"])
        record("mekh_aggregators", "OK", source=URL, publication_page=PAGE, aggregator_count=len(agg),
               detail="MEKH register of registered aggregators as published; not the MAVIR BRP list")
        sites = [(r["Licensee"], r["Licence_No"], r["Site"]) for r in sto]
        record("mekh_storage_licences", "OK", source=URL, publication_page=PAGE, licence_count=len(sto),
               nominal_power_mw_unrevoked=round(power, 3),
               rows_sharing_licence_and_site=len(sites) - len(set(sites)),
               detail="MEKH electricity-storage licences as published (licensed, not necessarily operating; rows copied as published, so a licence/site listed twice by MEKH appears twice)")
        return agg, sto
    except Exception as exc:
        detail = str(exc) if isinstance(exc, ValueError) else type(exc).__name__
        for n in ("mekh_aggregators", "mekh_storage_licences"):
            record(n, "FAILED", detail="Saved CSV retained; " + detail, source=PAGE)
        return None
    finally:
        if own:
            session.close()


if __name__ == "__main__":
    update_mekh_register()
