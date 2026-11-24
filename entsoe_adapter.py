"""Bounded Hungary ENTSO-E feeds: load forecast, physical flows, day-ahead prices.

Adapted from the Bosnia pipeline's adapter with Hungary's own area codes,
timezone and borders. All timestamps in the CSV primary key are UTC; local ISO
timestamps include the UTC offset, including repeated fall-back hours. Positive
directional quantities mean From_Area -> To_Area, not a signed net flow. MAW is
the ENTSO-E code for MW. Day-ahead prices are the official Transparency
Platform A44 auction for the HU bidding zone (primary owner: HUPX); reuse
rights follow ENTSO-E Terms of Use, not a blanket CC-BY grant.

The token travels only as the securityToken query parameter and is never
written to CSVs, status files, logs or error messages. No token is stored in
this repository.
"""
import csv
import datetime as dt
import io
import math
import os
from pathlib import Path
import re
import time
import xml.etree.ElementTree as ET
from zoneinfo import ZoneInfo

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from pipeline_status import atomic_text, record, utc_now

ENDPOINT = 'https://web-api.tp.entsoe.eu/api'
TZ = ZoneInfo('Europe/Budapest')
UTC = dt.timezone.utc
AREAS = {'HU': '10YHU-MAVIR----U', 'AT': '10YAT-APG------L',
         'SK': '10YSK-SEPS-----K', 'SI': '10YSI-ELES-----O',
         'RO': '10YRO-TEL------P', 'HR': '10YHR-HEP------M',
         'RS': '10YCS-SERBIATSOV', 'UA': '10YUA-WEPS-----0'}
NEIGHBORS = ('AT', 'SK', 'SI', 'RO', 'HR', 'RS', 'UA')
COLUMNS = ['Dataset', 'From_Area', 'To_Area', 'Interval_Start_UTC',
           'Interval_End_UTC', 'Interval_Start_Local', 'Resolution', 'Value_MW',
           'Source_Document', 'Revision', 'Fetched_At_UTC']
PRICE_COLUMNS = ["Market", "Country", "Product", "Delivery_Date", "Interval_Start_UTC",
                 "Interval_End_UTC", "Price_EUR_MWh", "Source", "Attribution", "License", "Fetched_At"]
PRICE_CSV = 'entsoe_day_ahead_prices.csv'


class DataError(ValueError):
    pass


class APIError(RuntimeError):
    pass


class StopAPI(APIError):
    """Stop this run on authentication or rate limiting, without leaking headers."""


def instant(text):
    value = dt.datetime.fromisoformat(text.replace('Z', '+00:00'))
    if value.tzinfo is None:
        raise DataError('Timezone missing')
    return value.astimezone(UTC)


def iso(value):
    return value.astimezone(UTC).isoformat().replace('+00:00', 'Z')


def day_bounds(day):
    start = dt.datetime.combine(day, dt.time(), TZ).astimezone(UTC)
    end = dt.datetime.combine(day + dt.timedelta(days=1), dt.time(), TZ).astimezone(UTC)
    return start, end


def text(node, path):
    value = node.findtext(path)
    if value is None:
        raise DataError('Missing required XML field: ' + path)
    return value.strip()


def _safe_text(value):
    # Keep only plain letters, digits and simple punctuation, and cut it short.
    # This stops URLs, XML and anything else from reaching the log.
    return re.sub(r"[^A-Za-z0-9 .,:()\-]", '', value).strip()[:150]


def parse_document(body, dataset, source, target, start, end, fetched_at):
    if b'<!DOCTYPE' in body.upper() or b'<!ENTITY' in body.upper():
        raise DataError('XML declarations not supported')
    try:
        root = ET.fromstring(body)
    except ET.ParseError as exc:
        raise DataError('Invalid XML') from exc
    if root.tag.split('}')[-1] == 'Acknowledgement_MarketDocument':
        codes = [n.text or '' for n in root.findall('.//{*}Reason/{*}code')]
        reasons = [n.text or '' for n in root.findall('.//{*}Reason/{*}text')]
        raise DataError('API acknowledgement, code ' + ','.join(codes) +
                        ': ' + _safe_text(' '.join(reasons)))
    expected_type = {'load_forecast': 'A65', 'physical_flows': 'A11'}[dataset]
    if text(root, '{*}type') != expected_type:
        raise DataError('Unexpected document type')
    if dataset == 'load_forecast' and text(root, '{*}process.processType') != 'A01':
        raise DataError('Unexpected forecast process')
    document = text(root, '{*}mRID')
    revision = int(text(root, '{*}revisionNumber'))
    series = root.findall('{*}TimeSeries')
    # Do not silently add duplicate/alternative series or choose a revision.
    if len(series) != 1:
        raise DataError('Expected exactly one aggregate directional series')
    ts = series[0]
    if text(ts, '{*}quantity_Measure_Unit.name') != 'MAW':
        raise DataError('Unexpected unit, MW required')
    if dataset == 'load_forecast':
        if text(ts, '{*}outBiddingZone_Domain.mRID') != AREAS['HU']:
            raise DataError('Unexpected forecast area')
    else:
        if (text(ts, '{*}out_Domain.mRID') != AREAS[source] or
                text(ts, '{*}in_Domain.mRID') != AREAS[target]):
            raise DataError('Direction mismatch')
    curve = text(ts, '{*}curveType')
    if curve not in ('A01', 'A03'):
        raise DataError('Unsupported curve type')
    result = {}
    for period in ts.findall('{*}Period'):
        pstart = instant(text(period, '{*}timeInterval/{*}start'))
        pend = instant(text(period, '{*}timeInterval/{*}end'))
        resolution = text(period, '{*}resolution')
        match = re.fullmatch(r'PT(\d+)M', resolution)
        if not match or int(match[1]) not in (15, 30, 60):
            raise DataError('Unsupported interval resolution')
        step = dt.timedelta(minutes=int(match[1]))
        count, remainder = divmod(pend - pstart, step)
        if remainder or count <= 0 or count > 10000:
            raise DataError('Invalid Period bounds')
        points = {}
        for point in period.findall('{*}Point'):
            position = int(text(point, '{*}position'))
            value = float(text(point, '{*}quantity'))
            if position in points or not 1 <= position <= count or not math.isfinite(value) or value < 0:
                raise DataError('Duplicate/invalid position or quantity')
            points[position] = value
        if not points or (curve == 'A03' and 1 not in points):
            raise DataError('Missing first block')
        current = None
        for position in range(1, count + 1):
            if position in points:
                current = points[position]
            elif curve == 'A01':
                current = None
            left = pstart + step * (position - 1)
            right = left + step
            if left < start or right > end:
                raise DataError('Period outside requested window')
            if left in result:
                raise DataError('Overlapping Periods')
            if current is None:
                raise DataError('Missing sequential interval')
            result[left] = {
                'Dataset': dataset, 'From_Area': source, 'To_Area': target,
                'Interval_Start_UTC': iso(left), 'Interval_End_UTC': iso(right),
                'Interval_Start_Local': left.astimezone(TZ).isoformat(),
                'Resolution': resolution, 'Value_MW': current,
                'Source_Document': document, 'Revision': revision,
                'Fetched_At_UTC': fetched_at,
            }
    rows = [result[key] for key in sorted(result)]
    if (not rows or instant(rows[0]['Interval_Start_UTC']) != start or
            instant(rows[-1]['Interval_End_UTC']) != end):
        raise DataError('Incomplete requested window')
    for previous, following in zip(rows, rows[1:]):
        if previous['Interval_End_UTC'] != following['Interval_Start_UTC']:
            raise DataError('Gap between Periods')
    return rows


def make_session():
    session = requests.Session()
    session.mount('https://', HTTPAdapter(max_retries=Retry(
        total=2, backoff_factor=1, status_forcelist=[500, 502, 503, 504],
        allowed_methods=['GET'], raise_on_status=False)))
    return session


def fetch(session, token, parameters):
    # ENTSO-E's Transparency API is a GET request. The token travels only as the
    # securityToken query parameter and is never copied into an error message.
    query = dict(parameters)
    query['securityToken'] = token
    try:
        response = session.get(ENDPOINT, params=query, timeout=(10, 45), allow_redirects=False)
    except requests.RequestException:
        # Discard the exception text: it can contain the full URL, including the token.
        raise APIError('Network request failed') from None
    if response.status_code in (401, 403, 429):
        raise StopAPI('Authentication/rate-limit HTTP ' + str(response.status_code))
    if response.status_code != 200:
        raise APIError('API HTTP ' + str(response.status_code))
    return response.content


def persist(path, rows, cutoff):
    # Replace only after EVERY direction in the dataset validates. Never erase
    # previously valid rows because a new request failed or had no data.
    path = Path(path)
    old = []
    if path.exists():
        with path.open(newline='', encoding='utf-8') as handle:
            reader = csv.DictReader(handle)
            if reader.fieldnames != COLUMNS:
                raise DataError('Existing CSV schema mismatch')
            old = list(reader)

    def key(row):
        return tuple(row[k] for k in ('Dataset', 'From_Area', 'To_Area', 'Interval_Start_UTC'))
    merged = {key(row): row for row in old if instant(row['Interval_Start_UTC']) >= cutoff}
    for row in rows:
        prior = merged.get(key(row))
        if prior and prior['Source_Document'] == row['Source_Document']:
            if int(prior['Revision']) > int(row['Revision']):
                raise DataError('Revision regression, stored CSV retained')
        merged[key(row)] = row
    out = io.StringIO(newline='')
    writer = csv.DictWriter(out, fieldnames=COLUMNS, lineterminator='\n')
    writer.writeheader()
    writer.writerows(merged[k] for k in sorted(merged))
    atomic_text(path, out.getvalue())


def persist_prices(path, rows, cutoff_day):
    path = Path(path)
    old = []
    if path.exists():
        with path.open(newline='', encoding='utf-8') as handle:
            reader = csv.DictReader(handle)
            if reader.fieldnames != PRICE_COLUMNS:
                raise DataError('Existing price CSV schema mismatch')
            old = list(reader)
    retained = [r for r in old if r.get('Delivery_Date', '') >= cutoff_day
                and r.get('Delivery_Date') != rows[0]['Delivery_Date']]
    out = io.StringIO(newline='')
    writer = csv.DictWriter(out, fieldnames=PRICE_COLUMNS, lineterminator='\n')
    writer.writeheader()
    writer.writerows(sorted(retained + rows, key=lambda r: r['Interval_Start_UTC']))
    atomic_text(path, out.getvalue())


def update_day_ahead_prices(output_dir='.', day=None, session=None, token=None):
    """Official A44 auction curve for one Budapest delivery day (primary owner: HUPX)."""
    from price_xml_parser import parse_prices
    day = day or (dt.datetime.now(TZ).date() - dt.timedelta(days=1))
    start, end = day_bounds(day)
    stamp = utc_now()
    token = token if token is not None else os.environ.get('ENTSOE_API_TOKEN', '').strip()
    owns_session = session is None
    session = session or make_session()
    meta = {'target_local_day': day.isoformat(), 'source': ENDPOINT,
            'coverage_start_utc': iso(start), 'coverage_end_utc': iso(end)}
    try:
        if not token:
            raise DataError('Token missing')
        body = fetch(session, token, {
            'documentType': 'A44', 'contract_MarketAgreement.type': 'A01',
            'in_Domain': AREAS['HU'], 'out_Domain': AREAS['HU'],
            'periodStart': start.strftime('%Y%m%d%H%M'), 'periodEnd': end.strftime('%Y%m%d%H%M')})
        parsed = parse_prices(body, AREAS['HU'])
        if parsed[0]['start'].astimezone(UTC) != start or parsed[-1]['end'].astimezone(UTC) != end:
            raise DataError('Wrong delivery-day boundaries')
        if any(a['end'] != b['start'] for a, b in zip(parsed, parsed[1:])):
            raise DataError('Delivery-day interval gap or overlap')
        rows = [{'Market': 'HU', 'Country': 'Hungary', 'Product': 'DAY_AHEAD',
                 'Delivery_Date': day.isoformat(),
                 'Interval_Start_UTC': iso(row['start']), 'Interval_End_UTC': iso(row['end']),
                 'Price_EUR_MWh': row['price'], 'Source': ENDPOINT,
                 'Attribution': 'ENTSO-E Transparency Platform; primary owner: HUPX',
                 'License': 'ENTSO-E Terms of Use; primary-owner rights apply, not a blanket CC-BY grant',
                 'Fetched_At': stamp} for row in parsed]
        cutoff = (day - dt.timedelta(days=29)).isoformat()
        persist_prices(Path(output_dir) / PRICE_CSV, rows, cutoff)
        record('entsoe_prices', 'OK', detail='Complete official auction day; primary owner HUPX',
               row_count=len(rows), unit='EUR/MWh', **meta)
        return rows
    except StopAPI as exc:
        record('entsoe_prices', 'FAILED', detail=str(exc), **meta)
        return None
    except (APIError, DataError) as exc:
        record('entsoe_prices', 'FAILED',
               detail='Request/validation failed: ' + str(exc).replace(token or '\x00', '[redacted]'), **meta)
        return None
    except (ValueError, OSError, KeyError, TypeError, ET.ParseError) as exc:
        record('entsoe_prices', 'FAILED', detail='Fetch/validation/write failed: ' + type(exc).__name__, **meta)
        return None
    finally:
        if owns_session:
            session.close()


def update_entsoe(output_dir='.', day=None, session=None, token=None, include_physical=True):
    day = day or (dt.datetime.now(TZ).date() - dt.timedelta(days=1))
    start, end = day_bounds(day)
    groups = {
        'load_forecast': [('HU', 'HU')],
        'physical_flows': [(a, b) for other in NEIGHBORS
                           for a, b in [('HU', other), (other, 'HU')]],
    }
    if not include_physical:
        groups.pop('physical_flows')
    token = token if token is not None else os.environ.get('ENTSOE_API_TOKEN', '').strip()
    owns_session = session is None
    session = session or make_session()
    stopped = False
    try:
        for dataset, directions in groups.items():
            name = 'entsoe_' + dataset
            meta = {'target_local_day': day.isoformat(), 'source': ENDPOINT,
                    'coverage_start_utc': iso(start), 'coverage_end_utc': iso(end)}
            if not token or stopped:
                record(name, 'FAILED', detail='Token missing' if not token else 'Stopped after authentication/rate limit', **meta)
                continue
            try:
                rows = []
                for source, target in directions:
                    parameters = {
                        'documentType': {'load_forecast': 'A65', 'physical_flows': 'A11'}[dataset],
                        'periodStart': start.strftime('%Y%m%d%H%M'),
                        'periodEnd': end.strftime('%Y%m%d%H%M'),
                    }
                    if dataset == 'load_forecast':
                        parameters.update({'processType': 'A01', 'outBiddingZone_Domain': AREAS['HU']})
                    else:
                        parameters.update({'in_Domain': AREAS[target], 'out_Domain': AREAS[source]})
                    body = fetch(session, token, parameters)
                    rows.extend(parse_document(body, dataset, source, target, start, end, utc_now()))
                    time.sleep(1)
                persist(Path(output_dir) / (name + '.csv'), rows, start - dt.timedelta(days=29))
                record(name, 'OK', detail='Complete requested historical day; not current-day data',
                       row_count=len(rows), directions=len(directions), unit='MW', **meta)
            except StopAPI as exc:
                stopped = True
                record(name, 'FAILED', detail=str(exc), **meta)
            except (APIError, DataError) as exc:
                # These messages are fixed strings from this code, or ENTSO-E reason
                # codes. They never contain the token or raw response text.
                record(name, 'FAILED', detail='Request/validation failed: ' + str(exc).replace(token, '[redacted]'), **meta)
            except (ValueError, OSError, KeyError, TypeError) as exc:
                # Generic errors may echo input text, so only the type name is recorded.
                record(name, 'FAILED', detail='Fetch/validation/write failed: ' + type(exc).__name__, **meta)
    finally:
        if owns_session:
            session.close()
