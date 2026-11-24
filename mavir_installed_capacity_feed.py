"""MAVIR gross installed capacity by primary source (MW), Hungarian power system.

Source: MAVIR "Energia MIX" publication (BT = brutto beepitett teljesitokepesseg), a PDF updated by MAVIR
with the latest day in its file name. The page link is discovered at run time. The first table gives
year-end values 2016-2025 plus the two most recent dates, excluding household-size small plants (HMKE) and
own-use plants (SCTE). This is the system-level answer to "plant MW by technology"; it is NOT a per-plant
list (the MEKH plant sheets carry licensees only). Unit: MW gross. Missing values stay blank.
"""
import csv
import datetime as dt
import io
import re
from pathlib import Path

import requests
from pipeline_status import atomic_text, record, utc_now
from substation_capacity_feed import pdf_text

PAGE = "https://www.mavir.hu/web/mavir/energia-mix-eromuvi-beepitett-teljesitokepesseg-adatok"
BASE = "https://www.mavir.hu"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"}
LINK = re.compile(r'href="(/documents/[^"]*BT_2015-(\d{8})_ig_BR[^"]*NT_HU[^"]*\.pdf[^"]*)"')
FUELS = [("Nukleáris", "Nuclear"), ("Szén, lignit", "Coal_Lignite"), ("Földgáz", "Gas"), ("Tüzelőolaj", "Oil"),
         ("akkumlátoros", "Battery_Storage_Market_plus_MAVIR"), ("Nap ipari", "Solar_PV_Industrial"),
         ("Szél", "Wind"), ("Biomassza", "Biomass_Biogas"), ("Víz", "Hydro"), ("Hulladék", "Waste"),
         ("Geotermikus", "Geothermal"), ("Egyéb megújuló", "Other_Renewable"),
         ("bruttó BT összesen", "TOTAL_excl_HMKE_SCTE")]
YEARS = [str(y) for y in range(2016, 2026)]
COLUMNS = ["Fuel_Group", "Period", "Gross_Installed_MW", "Source", "Retrieved_UTC"]
NUM = re.compile(r"-?\d+(?:,\d+)?")


def _f(s):
    return float(s.replace(",", "."))


def parse_text(text, last_date, source, stamp):
    """last_date: date of the newest column (from the file name); the column before it is the day earlier."""
    first = text.split("Primer forrás")[0]
    lines = first.splitlines()
    found = []
    for i, line in enumerate(lines):
        nums = NUM.findall(line.split("%")[0] if "%" in line else line)
        body = re.split(r"\s{2,}\d", line, maxsplit=1)[0]
        values = re.findall(r"(?<![\d,])-?\d+,\d(?![\d,])", line)
        if len(values) >= 12:
            ctx = " ".join(lines[max(0, i - 3):i + 1])
            found.append((ctx, [_f(v) for v in values[:12]]))
    if len(found) != len(FUELS):
        raise ValueError("Installed-capacity table layout changed (%d rows)" % len(found))
    periods = YEARS + [(last_date - dt.timedelta(days=1)).isoformat(), last_date.isoformat()]
    rows, by = [], {}
    for (key, name), (ctx, vals) in zip(FUELS, found):
        if key not in ctx:
            raise ValueError("Unexpected row order near " + name)
        if any(v < 0 for v in vals):
            raise ValueError("Negative capacity")
        by[name] = vals
        rows += [dict(zip(COLUMNS, [name, p, v, source, stamp])) for p, v in zip(periods, vals)]
    for i in range(12):
        parts = sum(by[n][i] for _, n in FUELS[:-1])
        if abs(parts - by["TOTAL_excl_HMKE_SCTE"][i]) > 1.5:
            raise ValueError("Fuel rows do not sum to published total")
    return rows


def _csv(rows):
    out = io.StringIO(newline="")
    w = csv.DictWriter(out, fieldnames=COLUMNS, lineterminator="\n")
    w.writeheader()
    w.writerows(rows)
    return out.getvalue()


def update_installed_capacity(output_dir=".", session=None):
    stamp = utc_now()
    own = session is None
    session = session or requests.Session()
    try:
        page = session.get(PAGE, headers=HEADERS, timeout=(10, 60))
        m = LINK.search(page.text) if page.status_code == 200 else None
        if not m:
            raise ValueError("Installed-capacity PDF link not found on MAVIR page")
        href, ymd = m.group(1).replace("&amp;", "&"), m.group(2)
        last = dt.date(int(ymd[:4]), int(ymd[4:6]), int(ymd[6:]))
        resp = session.get(BASE + href, headers=HEADERS, timeout=(10, 90))
        if resp.status_code != 200 or resp.content[:4] != b"%PDF":
            raise ValueError("Download failed or not PDF (HTTP " + str(resp.status_code) + ")")
        rows = parse_text(pdf_text(resp.content), last, BASE + href, stamp)
        atomic_text(str(Path(output_dir) / "mavir_installed_capacity.csv"), _csv(rows))
        tot = [r for r in rows if r["Fuel_Group"] == "TOTAL_excl_HMKE_SCTE" and r["Period"] == last.isoformat()][0]
        record("mavir_installed_capacity", "OK", source=BASE + href, publication_page=PAGE, row_count=len(rows),
               latest_period=last.isoformat(), total_gross_mw=tot["Gross_Installed_MW"],
               detail="MAVIR gross installed capacity by primary source (MW), excl. household-size (HMKE) and "
                      "own-use (SCTE) plants; system level, not per plant; updated by MAVIR on its own schedule")
        return rows
    except Exception as exc:
        detail = str(exc) if isinstance(exc, ValueError) else type(exc).__name__
        record("mavir_installed_capacity", "FAILED", detail="Saved CSV retained; " + detail, source=PAGE)
        return None
    finally:
        if own:
            session.close()


if __name__ == "__main__":
    update_installed_capacity()
