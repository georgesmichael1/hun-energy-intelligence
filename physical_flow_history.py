"""31 Budapest delivery dates of raw directional MW, including published today.
Covers all seven Hungarian borders in both directions (14 directional series).
Does not change load forecasts or day-ahead prices.
"""
import copy
import collections
import csv
import datetime as dt
import io
import json
from pathlib import Path
import time
import xml.etree.ElementTree as ET
import entsoe_adapter as e
from pipeline_status import atomic_text, utc_now

DIRECTIONS = [(a, b) for n in e.NEIGHBORS for a, b in [('HU', n), (n, 'HU')]]


def parse_range(body, source, target, first_day, last_day, now):
    if b'<!DOCTYPE' in body.upper() or b'<!ENTITY' in body.upper():
        raise e.DataError('XML declarations not supported')
    root = ET.fromstring(body)
    if root.tag.split('}')[-1] == 'Acknowledgement_MarketDocument':
        raise e.DataError('No data document returned')
    series = root.findall('{*}TimeSeries')
    if not series:
        raise e.DataError('No directional series')
    start = e.day_bounds(first_day)[0]
    end = e.day_bounds(last_day)[1]
    rows = []
    for original_ts in series:
        periods = original_ts.findall('{*}Period')
        if not periods:
            raise e.DataError('No published periods')
        for period in periods:
            left = e.instant(e.text(period, '{*}timeInterval/{*}start'))
            right = e.instant(e.text(period, '{*}timeInterval/{*}end'))
            if left < start or right > end:
                raise e.DataError('Period outside requested window')
            part = copy.deepcopy(root)
            for ts in list(part.findall('{*}TimeSeries')):
                part.remove(ts)
            ts = copy.deepcopy(original_ts)
            for p in list(ts.findall('{*}Period')):
                ts.remove(p)
            published = copy.deepcopy(period)
            if right > now:
                resolution = e.text(period, '{*}resolution')
                seconds = {'PT15M': 900, 'PT30M': 1800, 'PT60M': 3600}.get(resolution)
                if seconds is None:
                    raise e.DataError('Unsupported resolution')
                completed = [p for p in published.findall('{*}Point')
                             if left + dt.timedelta(seconds=seconds * int(e.text(p, '{*}position'))) <= now]
                if not completed:
                    continue
                last = max(int(e.text(p, '{*}position')) for p in completed)
                right = min(right, left + dt.timedelta(seconds=seconds * last))
                published.find('{*}timeInterval/{*}end').text = e.iso(right)
                for p in list(published.findall('{*}Point')):
                    if int(e.text(p, '{*}position')) > last:
                        published.remove(p)
            ts.append(published)
            part.append(ts)
            rows.extend(e.parse_document(ET.tostring(part), 'physical_flows', source, target, left, right, utc_now()))
    rows.sort(key=lambda r: e.instant(r['Interval_Start_UTC']))
    unique = {}
    for r in rows:
        left = e.instant(r['Interval_Start_UTC'])
        right = e.instant(r['Interval_End_UTC'])
        if right > now:
            raise e.DataError('Future measured interval')
        if left in unique:
            prior = unique[left]
            # Exact repeated measurements are harmless; conflicting alternatives are not guessed.
            if any(prior[k] != r[k] for k in ('Interval_End_UTC', 'Resolution', 'Value_MW', 'Source_Document', 'Revision')):
                raise e.DataError('Conflicting duplicate interval')
        unique[left] = r
    rows = [unique[k] for k in sorted(unique)]
    for a, b in zip(rows, rows[1:]):
        if e.instant(a['Interval_End_UTC']) > e.instant(b['Interval_Start_UTC']):
            raise e.DataError('Overlapping intervals')
    return rows


def parse_available(body, source, target, day, now):
    return parse_range(body, source, target, day, day, now)


def complete(rows, day):
    start, end = e.day_bounds(day)
    return bool(rows and e.instant(rows[0]['Interval_Start_UTC']) == start and
                e.instant(rows[-1]['Interval_End_UTC']) == end and
                all(a['Interval_End_UTC'] == b['Interval_Start_UTC'] for a, b in zip(rows, rows[1:])))


def classify_failure(exc):
    """Separate a source that publishes nothing from transport, validation and auth problems."""
    msg = str(exc)[:200]
    if isinstance(exc, e.DataError) and msg.startswith('API acknowledgement, code 999'):
        return dict(category='SOURCE_NO_DATA', reason=msg + ' (ENTSO-E publishes no data for this direction/window; saved intervals retained)')
    if isinstance(exc, e.APIError):
        return dict(category='TRANSPORT', reason=msg + '; saved intervals retained')
    return dict(category='VALIDATION', reason=msg + '; saved intervals retained')


def update(output_dir='.', now=None, session=None, token=None):
    now = now or dt.datetime.now(dt.timezone.utc)
    today = now.astimezone(e.TZ).date()
    oldest = today - dt.timedelta(days=30)
    cutoff = e.day_bounds(oldest)[0]
    root = Path(output_dir)
    path = root / 'entsoe_physical_flows.csv'
    token = token if token is not None else e.os.environ.get('ENTSOE_API_TOKEN', '').strip()
    old = []
    if path.exists():
        with path.open(newline='', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            if reader.fieldnames != e.COLUMNS:
                raise e.DataError('Existing CSV schema mismatch')
            old = list(reader)
    merged = {(r['From_Area'], r['To_Area'], r['Interval_Start_UTC']): r
              for r in old if cutoff <= e.instant(r['Interval_Start_UTC']) < e.day_bounds(today)[1]}
    owns = session is None
    session = session or e.make_session()
    failures = []
    stopped = False
    try:
        for source, target in DIRECTIONS:
            if not token or stopped:
                failures.append(dict(direction=source + '>' + target,
                                     reason='Token missing' if not token else 'Stopped after authentication/rate limit'))
                continue
            # User-confirmed incremental design: refetch today/yesterday plus any
            # missing or incomplete older day; retain complete saved history.
            fetch_first = today - dt.timedelta(days=1)
            for offset in range(29):
                day = oldest + dt.timedelta(days=offset)
                saved = sorted([r for r in merged.values()
                                if r['From_Area'] == source and r['To_Area'] == target
                                and r['Interval_Start_Local'][:10] == str(day)],
                               key=lambda r: r['Interval_Start_UTC'])
                if not complete(saved, day):
                    fetch_first = day
                    break
            begin = time.monotonic()
            try:
                body = e.fetch(session, token, dict(
                    documentType='A11', in_Domain=e.AREAS[target], out_Domain=e.AREAS[source],
                    periodStart=e.day_bounds(fetch_first)[0].strftime('%Y%m%d%H%M'),
                    periodEnd=e.day_bounds(today)[1].strftime('%Y%m%d%H%M')))
                incoming = parse_range(body, source, target, fetch_first, today, now)
                for r in incoming:
                    k = (source, target, r['Interval_Start_UTC'])
                    prior = merged.get(k)
                    if prior and prior['Source_Document'] == r['Source_Document'] and int(prior['Revision']) > int(r['Revision']):
                        raise e.DataError('Revision regression')
                for r in incoming:
                    merged[source, target, r['Interval_Start_UTC']] = r
                print(json.dumps(dict(direction=source + '>' + target, rows=len(incoming),
                                      elapsed_seconds=round(time.monotonic() - begin, 2))), flush=True)
            except e.StopAPI:
                stopped = True
                failures.append(dict(direction=source + '>' + target, category='AUTH_OR_RATE_LIMIT',
                                     reason='Authentication or rate limit'))
            except (e.APIError, e.DataError, ValueError, KeyError, TypeError, ET.ParseError) as exc:
                failures.append(dict(direction=source + '>' + target, **classify_failure(exc)))
                print(json.dumps(dict(direction=source + '>' + target, result='unavailable_or_invalid',
                                      elapsed_seconds=round(time.monotonic() - begin, 2))), flush=True)
            time.sleep(1)
    finally:
        if owns:
            session.close()
    rows = [merged[k] for k in sorted(merged)]
    out = io.StringIO(newline='')
    w = csv.DictWriter(out, fieldnames=e.COLUMNS, lineterminator='\n')
    w.writeheader()
    w.writerows(rows)
    atomic_text(path, out.getvalue())
    coverage = []
    for offset in range(31):
        day = oldest + dt.timedelta(days=offset)
        for a, b in DIRECTIONS:
            rr = [r for r in rows if r['From_Area'] == a and r['To_Area'] == b
                  and r['Interval_Start_Local'][:10] == str(day)]
            coverage.append(dict(day=str(day), direction=a + '>' + b, rows=len(rr),
                                 complete=complete(rr, day),
                                 latest_end_utc=max((r['Interval_End_UTC'] for r in rr), default=None)))
    historic_gaps = [x for x in coverage if x['day'] < str(today) and not x['complete']]
    no_data_dirs = {f['direction'] for f in failures if f.get('category') == 'SOURCE_NO_DATA'}
    hard_failures = [f for f in failures if f.get('category') != 'SOURCE_NO_DATA']
    unexplained = [g for g in historic_gaps if g['direction'] not in no_data_dirs]
    if not historic_gaps and not failures:
        state = 'OK'
    elif not hard_failures and not unexplained:
        state = 'SOURCE_PARTIAL'   # only directions the source publishes nothing for
    else:
        state = 'PARTIAL'          # transport, validation, auth or unexplained gaps
    status = dict(state=state, checked_at=utc_now(), timezone='Europe/Budapest',
                  requested_first_day=str(oldest), requested_last_day=str(today),
                  requested_delivery_days=31, unit='MW', row_count=len(rows),
                  historical_incomplete_pairs=historic_gaps, request_failures=failures,
                  failure_categories=dict(collections.Counter(f.get('category', 'UNKNOWN') for f in failures)),
                  coverage=coverage,
                  note='Raw directional measurements. Today contains published completed intervals only. '
                       'Net export = HU outbound minus inbound; match exact UTC interval and resolution. '
                       'Missing pairs are gaps, never zero.')
    atomic_text(root / 'physical_flows_status.json', json.dumps(status, indent=2) + '\n')
    print(json.dumps({k: status[k] for k in ('state', 'requested_first_day', 'requested_last_day', 'row_count')}))
    return status


if __name__ == '__main__':
    status = update()
    # Persist partial good data and coverage before workflow quality gate.
    raise SystemExit(0 if status['state'] == 'OK' else 1)
