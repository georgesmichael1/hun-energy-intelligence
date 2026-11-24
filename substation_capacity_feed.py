"""Distribution-substation free consumption (offtake) capacity published by E.ON Hungary.

Source: the E.ON "Grid hosting capacity" PDF for the EEL (ELMU area) and EDE labels. Each row is a
HV/MV substation with the voltage level [kV], the currently releasable offtake capacity [MVA] on the
MV network and the capacity probable in the next five years [MVA]. The PDF states quarterly updates.

Scope limits, kept in the data and the status detail:
- consumption (offtake) side only, and only these distribution licensees. Other Hungarian
  DSOs (MVM Halozat, MVM Emasz) and MAVIR transmission-node publications are NOT parsed here.
- the same substation can appear twice with different voltage levels; rows are copied as published.
- MVA, not MW. Nothing is converted or inferred. A missing value stays blank, never zero.
"""
import csv
import io
import re
import subprocess
import tempfile
from pathlib import Path

import requests
from pipeline_status import atomic_text, record, utc_now

URL = "https://www.eon.hu/content/dam/eon/eon-hungary/documents/Lakossagi/aram/szabad_kapacitas_ELMU.pdf"
HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                         "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"}
COLUMNS = ["Table_Label", "Substation", "Voltage_kV", "Releasable_Offtake_MVA", "Probable_5y_Offtake_MVA",
           "Source", "Retrieved_UTC"]
SECTION = re.compile(r"Grid hosting capacity\s*-\s*([A-Z]{2,4})")
ROW = re.compile(r"^(?P<name>\S.*?)\s{2,}(?P<kv>\d{1,3})\s+(?P<now>\d{1,4})\s+(?P<fut>\d{1,4})\s*$")
VOLTAGES = {11, 22, 35}


def parse_text(text, stamp, source=URL):
    rows, label, seen = [], None, set()
    for line in text.splitlines():
        m = SECTION.search(line)
        if m:
            label = m.group(1)
            continue
        if label is None or "Módszertan" in line:
            if "Módszertan" in line:
                label = None
            continue
        m = ROW.match(line.strip("\f"))
        if not m:
            continue
        kv = int(m.group("kv"))
        if kv not in VOLTAGES:
            raise ValueError("Unexpected voltage level " + str(kv) + " kV")
        key = (label, m.group("name").strip(), kv)
        if key in seen:
            raise ValueError("Duplicate substation row")
        seen.add(key)
        rows.append({"Table_Label": label, "Substation": m.group("name").strip(), "Voltage_kV": kv,
                     "Releasable_Offtake_MVA": int(m.group("now")),
                     "Probable_5y_Offtake_MVA": int(m.group("fut")), "Source": source, "Retrieved_UTC": stamp})
    if len(rows) < 40 or len({r["Table_Label"] for r in rows}) < 1:
        raise ValueError("Substation table implausibly short or layout changed")
    return rows


def pdf_text(content):
    with tempfile.TemporaryDirectory() as tmp:
        path = Path(tmp) / "cap.pdf"
        path.write_bytes(content)
        out = subprocess.run(["pdftotext", "-layout", str(path), "-"], capture_output=True, timeout=60, check=True)
    return out.stdout.decode("utf-8")


def _csv(rows):
    out = io.StringIO(newline="")
    w = csv.DictWriter(out, fieldnames=COLUMNS, lineterminator="\n")
    w.writeheader()
    w.writerows(rows)
    return out.getvalue()


def update_substation_capacity(output_dir=".", session=None):
    stamp = utc_now()
    own = session is None
    session = session or requests.Session()
    try:
        resp = session.get(URL, headers=HEADERS, timeout=(10, 60))
        if resp.status_code != 200 or resp.content[:4] != b"%PDF":
            raise ValueError("Download failed or not PDF (HTTP " + str(resp.status_code) + ")")
        rows = parse_text(pdf_text(resp.content), stamp)
        atomic_text(str(Path(output_dir) / "hun_substation_capacity.csv"), _csv(rows))
        record("substation_capacity", "OK", source=URL, row_count=len(rows),
               labels=sorted({r["Table_Label"] for r in rows}),
               releasable_mva_total=sum(r["Releasable_Offtake_MVA"] for r in rows),
               detail="E.ON Hungary published offtake hosting capacity per HV/MV substation (MVA, not MW); "
                      "E.ON areas only, not MVM or MAVIR; quarterly publication, not real time")
        return rows
    except Exception as exc:
        detail = str(exc) if isinstance(exc, ValueError) else type(exc).__name__
        record("substation_capacity", "FAILED", detail="Saved CSV retained; " + detail, source=URL)
        return None
    finally:
        if own:
            session.close()


if __name__ == "__main__":
    update_substation_capacity()
