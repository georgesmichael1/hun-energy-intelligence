"""Refresh only the current-year annual nominal USD GDP projection/estimate for Hungary.
Historical GDP rows are deliberately not revised by this consumer.
"""
import csv
import datetime as dt
import io
import math
from pathlib import Path
from pipeline_status import atomic_text, record

SOURCE = 'https://www.imf.org/external/datamapper/api/v2/NGDPD'
COUNTRY = 'HUN'
COLUMNS = ['Year', 'Value_USD_Billions', 'Status', 'Vintage', 'Source_Modified', 'Source', 'Retrieved_UTC']


def parse_current(payload, year, stamp):
    meta = payload['indicators']['NGDPD']
    if meta['unit'] != 'Billions of U.S. dollars' or meta['dataset'] != 'WEO' or meta['label'] != 'GDP, current prices':
        raise ValueError('IMF GDP indicator/unit changed')
    vintage = meta['source']
    if not vintage.startswith('World Economic Outlook ('):
        raise ValueError('IMF vintage missing')
    modified = meta['last-modified']
    if dt.datetime.fromisoformat(modified).year < year:
        raise ValueError('No current-year IMF release')
    projection = meta['projection-year']
    if not isinstance(projection, int):
        raise ValueError('Missing projection boundary')
    value = payload['values']['NGDPD'][COUNTRY][str(year)]
    if isinstance(value, bool) or not isinstance(value, (float, int)) or not math.isfinite(value) or value <= 0:
        raise ValueError('Invalid current-year GDP')
    # Do not upgrade WEO staff estimates into official finalized national accounts.
    status = 'IMF forecast' if year >= projection else 'IMF estimate'
    return dict(zip(COLUMNS, [year, value, status, vintage, modified, SOURCE, stamp]))


def update_current_gdp(path='current_gdp.csv', session=None, now=None):
    now = now or dt.datetime.now(dt.timezone.utc)
    try:
        if session is None:
            # The IMF datamapper rejects plain requests (HTTP 403); impersonate Chrome.
            from curl_cffi import requests as cffi_requests
            session = cffi_requests.Session(impersonate='chrome124')
        r = session.get(SOURCE, timeout=(10, 35))
        r.raise_for_status()
        row = parse_current(r.json(), now.year, now.isoformat(timespec='seconds'))
        if Path(path).exists():
            prior = next(csv.DictReader(Path(path).open(encoding='utf-8')))
            if row['Source_Modified'] < prior['Source_Modified']:
                raise ValueError('IMF release would roll backwards')
        out = io.StringIO(newline='')
        w = csv.DictWriter(out, fieldnames=COLUMNS, lineterminator='\n')
        w.writeheader()
        w.writerow(row)
        atomic_text(path, out.getvalue())
        record('current_gdp', 'OK', detail='Annual nominal GDP, current-year IMF forecast/estimate, not a real-time observation',
               year=row['Year'], value_usd_billions=row['Value_USD_Billions'], classification=row['Status'],
               vintage=row['Vintage'], source_modified=row['Source_Modified'], source=SOURCE)
        return row
    except Exception as exc:
        record('current_gdp', 'FAILED', detail='Saved dated GDP retained; ' + type(exc).__name__ + ': ' + str(exc), source=SOURCE)
        return None


if __name__ == '__main__':
    import json
    from pipeline_status import DATASETS
    update_current_gdp()
    path = Path('status.json')
    status = json.loads(path.read_text(encoding='utf-8')) if path.exists() else {'datasets': {}, 'dataset_details': {}}
    entry = DATASETS['current_gdp']
    old = status.get('dataset_details', {}).get('current_gdp', {})
    entry['last_success_at'] = entry['checked_at'] if entry['state'] == 'OK' else old.get('last_success_at')
    status.setdefault('dataset_details', {})['current_gdp'] = entry
    status.setdefault('datasets', {})['current_gdp'] = 'OK' if entry['state'] == 'OK' else 'FAILED: ' + entry['detail']
    atomic_text(path, json.dumps(status, indent=2, allow_nan=False) + '\n')
