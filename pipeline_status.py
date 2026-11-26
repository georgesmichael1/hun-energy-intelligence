"""Per-dataset run state. A fetch timestamp is not a source verification date."""

import datetime as dt
import json
import os
import threading
from pathlib import Path

DATASETS = {}
# The token-free update and the ENTSO-E daily batch may share a process in
# parallel threads; both record() and write_status() touch DATASETS, so they
# are serialized here (a dict changing size mid-iteration would otherwise be
# able to crash write_status).
_LOCK = threading.RLock()


def utc_now():
    return dt.datetime.now(dt.timezone.utc).isoformat(timespec='seconds')


def record(name, state, *, detail='', **metadata):
    with _LOCK:
        now = utc_now()
        entry = {
            'state': state,
            'checked_at': now,
            'last_attempt_at': metadata.pop('last_attempt_at', now),
            'detail': detail,
            **metadata,
        }
        DATASETS[name] = entry


def atomic_text(path, text):
    path = Path(path)
    tmp = path.with_name(path.name + '.tmp')
    # newline='' stops Windows from adding a second line break.
    with open(tmp, 'w', encoding='utf-8', newline='') as handle:
        handle.write(text)
    os.replace(tmp, path)


def write_status(path, *, rows=None, legacy=None):
    with _LOCK:
        return _write_status_locked(path, rows=rows, legacy=legacy)


def _write_status_locked(path, *, rows=None, legacy=None):
    try:
        old = json.loads(Path(path).read_text(encoding='utf-8'))
    except (OSError, ValueError):
        old = {}
    previous = old.get('dataset_details', {})
    for name, item in DATASETS.items():
        prior = previous.get(name, {})
        item['last_attempt_at'] = item.get('last_attempt_at') or item['checked_at']
        if item['state'] == 'OK':
            item['last_success_at'] = item.get('last_success_at') or item['checked_at']
        else:
            item['last_success_at'] = prior.get('last_success_at')

    payload = {
        'schema_version': 2,
        'run_finished_at': utc_now(),
        # Kept for legacy VBA, deliberately not restamped from today's run.
        'last_verified_date': old.get('last_verified_date', ''),
        'last_verified_date_is_legacy': True,
        'verification_note': 'No whole-master verification performed; use dataset_details.',
        'rows': rows if rows is not None else old.get('rows'),
        **(legacy or {}),
        'datasets': {k: (v['state'] if v['state'] == 'OK' else
                          ('FAILED: ' if v['state'] in ('FAILED', 'PARTIAL') else '')
                          + v['state'] + ': ' + v['detail'])
                     for k, v in DATASETS.items()},
        'dataset_details': DATASETS,
    }
    atomic_text(path, json.dumps(payload, indent=2, allow_nan=False) + '\n')
    return payload


def problems(status):
    issues = [f'{k}: {v}' for k, v in status.get('datasets', {}).items()
              if v != 'OK' and status.get('dataset_details', {}).get(k, {}).get('state') not in ('STATIC_REFERENCE', 'SOURCE_PARTIAL')]
    if not status.get('datasets'):
        issues.append('Dataset state missing')
    return issues
