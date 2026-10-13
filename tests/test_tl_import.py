# -*- coding: utf-8 -*-
"""Импорт TravelLine: разбор броней в строки Excel (фикстуры, без сети)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import mapping_valid, read_table
from schemas import DNA_DATA_SCHEMA, TCH_DATA_SCHEMA
from tl_client import (TL_DATA_MAP, TL_EXTRA, bookings_to_rows, channel_label,
                       is_tl_export, write_data_workbook, write_rooms_stub)


def _stay(arr='2026-07-10T14:00', dep='2026-07-12T12:00', room=10000.0,
          services=(), guests=(('Иван', 'Петров'),), rt='Стандарт',
          rp='Базовый'):
    return {
        'stayDates': {'arrivalDateTime': arr, 'departureDateTime': dep},
        'total': {'priceAfterTax': room},
        'services': [{'total': {'priceAfterTax': s}} for s in services],
        'extraStayCharges': {'earlyArrival': None, 'lateDeparture': None},
        'guests': [{'firstName': f, 'lastName': l} for f, l in guests],
        'roomType': {'name': rt}, 'ratePlans': [{'name': rp}],
    }


def _book(number='20260701-8186-1', status='Active', stays=None,
          book_services=(), total=10000.0, source=None):
    return {'number': number, 'status': status,
            'roomStays': stays if stays is not None else [_stay()],
            'services': [{'total': {'priceAfterTax': s}}
                         for s in book_services],
            'total': {'priceAfterTax': total},
            'source': source or {'type': 'Channel', 'code': 'YAX'}}

QS, QE = '2026-07-01', '2026-10-01'


def test_plain_booking_no_services():
    rows, rep = bookings_to_rows([_book()], QS, QE)
    assert rep['votes'] == {'A': 0, 'B': 0, '?': 0}  # неинформативна
    assert len(rows) == 1
    r = rows[0]
    assert r['Стоимость'] == 10000.0 and r['Допуслуги'] == 0.0
    assert r['Гость'] == 'Иван Петров' and r['Ночей'] == 2
    assert r['Источник'] == 'TravelLine YAX'
    assert r['№ брони'] == '20260701-8186-1'


def test_mode_a_with_paid_services():
    b = _book(stays=[_stay(services=(1500.0,))], total=11500.0)
    rows, rep = bookings_to_rows([b], QS, QE)
    assert rep['votes']['A'] == 1 and not rep['mismatch']
    assert rows[0]['Стоимость'] == 10000.0
    assert rows[0]['Допуслуги'] == 1500.0


def test_mode_b_extras_inside():
    b = _book(stays=[_stay(room=11500.0, services=(1500.0,))], total=11500.0)
    rows, rep = bookings_to_rows([b], QS, QE)
    assert rep['votes']['B'] == 1
    assert rows[0]['Стоимость'] == 10000.0  # 11500 − 1500
    assert rows[0]['Допуслуги'] == 1500.0


def test_mismatch_flagged_base_kept():
    b = _book(stays=[_stay(services=(1500.0,))], total=99999.0)
    rows, rep = bookings_to_rows([b], QS, QE)
    assert rep['votes']['?'] == 1
    assert len(rep['mismatch']) == 1
    assert rep['mismatch'][0]['number'] == '20260701-8186-1'
    assert rows[0]['Стоимость'] == 10000.0  # база как есть + флаг


def test_inclusive_zero_neutral():
    b = _book(stays=[_stay(services=(0.0,))], total=10000.0)
    rows, rep = bookings_to_rows([b], QS, QE)
    assert rep['votes'] == {'A': 0, 'B': 0, '?': 0}  # неинформативна
    assert rows[0]['Стоимость'] == 10000.0


def test_cancelled_skipped_and_quarter_filter():
    cancelled = _book(number='c', status='Cancelled')
    out_q = _book(number='o', stays=[_stay(arr='2026-10-05T14:00',
                                           dep='2026-10-07T12:00')])
    rows, rep = bookings_to_rows([cancelled, out_q, _book()], QS, QE)
    assert rep['cancelled'] == 1 and rep['in_quarter'] == 1
    assert [r['№ брони'] for r in rows] == ['20260701-8186-1']


def test_group_booking_two_rows():
    b = _book(stays=[_stay(guests=(('А', 'А'),)),
                     _stay(room=8000.0, guests=(('Б', 'Б'),))],
              total=18000.0)
    rows, rep = bookings_to_rows([b], QS, QE)
    assert len(rows) == 2 and rep['rows'] == 2
    assert [r['Гость'] for r in rows] == ['А А', 'Б Б']


def test_channels():
    assert channel_label({'type': 'BookingEngine', 'code': 'PropertySite'}) \
        == 'TravelLine сайт'
    assert channel_label({'type': 'PMS', 'code': None}) == 'TravelLine ресепшен'
    assert channel_label({'type': 'Channel', 'code': 'OTK'}) == 'TravelLine OTK'
    assert channel_label({}) == 'TravelLine ?'


def test_workbooks_roundtrip_and_mapping(tmp_path):
    rows, _ = bookings_to_rows([_book()], QS, QE)
    dp = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rp = str(tmp_path / 'TL_tch_2026III_rooms.xlsx')
    write_data_workbook(rows, dp)
    write_rooms_stub(rp)
    assert is_tl_export(dp) and is_tl_export(rp)
    assert not is_tl_export(str(tmp_path / 'other.xlsx'))
    df = read_table(dp)
    assert list(df.columns) == ['Гость', 'Стоимость', 'Заезд', 'Выезд',
                                'Источник', 'Категория номера', '№ брони',
                                'Ночей', 'Тариф', 'Допуслуги', 'Итог брони']
    assert df.iloc[0]['Стоимость'] == 10000.0
    assert df.iloc[0]['Гость'] == 'Иван Петров'
    headers = [str(c) for c in df.columns]
    assert mapping_valid(TL_DATA_MAP, headers, TCH_DATA_SCHEMA)
    assert mapping_valid(TL_DATA_MAP, headers, DNA_DATA_SCHEMA)
    assert all(c in headers for c in TL_EXTRA)
    dr = read_table(rp)
    assert list(dr.columns) == ['ФИО', '№ комнаты'] and len(dr) == 0


def _fake_http_factory(log):
    def fake(method, path, token=None, params=None, data=None, ctype=None):
        log.append(path)
        if path == '/auth/token':
            return 200, {'access_token': 'tok'}
        if path.endswith('/bookings'):
            return 200, {'bookingSummaries': [{'number': 'n1'},
                                              {'number': 'n2'}],
                         'hasMoreData': False}
        return 200, {'booking': {'number': path.rsplit('/', 1)[-1],
                                 'status': 'Active'}}
    return fake


def test_fetch_details_progress(monkeypatch):
    import threading
    import tl_client
    log = []
    calls = []
    monkeypatch.setattr(tl_client, '_http', _fake_http_factory(log))
    monkeypatch.setattr(tl_client.time, 'sleep', lambda s: None)
    out = tl_client.fetch_details('id', 'sec', '8186', '2024-01-01',
                                  progress=lambda st, d, t: calls.append((st, d, t)),
                                  cancel_event=threading.Event())
    assert [b['number'] for b in out] == ['n1', 'n2']
    assert ('list', 2, 0) in calls
    assert ('details', 2, 2) in calls


def test_cancel_before_start():
    import threading
    import tl_client
    ev = threading.Event()
    ev.set()
    try:
        tl_client.fetch_details('id', 'sec', '8186',
                                cancel_event=ev)
    except tl_client.TLCancelled:
        pass
    else:
        raise AssertionError('установленный флаг обязан дать TLCancelled')


def test_cancel_midway(monkeypatch):
    import threading
    import tl_client
    ev = threading.Event()
    monkeypatch.setattr(tl_client, '_http', _fake_http_factory([]))
    monkeypatch.setattr(tl_client.time, 'sleep', lambda s: None)

    def progress(stage, done, total):
        if stage == 'details':
            ev.set()  # отмена после первой детали

    try:
        tl_client.fetch_details('id', 'sec', '8186', '2024-01-01',
                                progress=progress, cancel_event=ev)
    except tl_client.TLCancelled:
        pass
    else:
        raise AssertionError('отмена на середине обязана дать TLCancelled')


def test_list_401_refreshes_token_and_retries(monkeypatch):
    import tl_client
    tokens = iter(['expired', 'fresh'])
    monkeypatch.setattr(tl_client, 'get_token', lambda *a: next(tokens))
    calls = []

    def fake(method, path, token=None, params=None, data=None, ctype=None):
        calls.append(token)
        if token == 'expired':
            raise tl_client.TLApiError('HTTP 401: {"message":"Unauthorized"}')
        return 200, {'bookingSummaries': [], 'hasMoreData': False}

    monkeypatch.setattr(tl_client, '_http', fake)
    out = tl_client.fetch_details('id', 'secret', '8186')
    assert out == []
    assert calls == ['expired', 'fresh']


def test_persistent_list_401_reports_stage_and_setup_hint(monkeypatch):
    import tl_client
    tokens = iter(['expired', 'fresh'])
    monkeypatch.setattr(tl_client, 'get_token', lambda *a: next(tokens))

    def fake(*args, **kwargs):
        raise tl_client.TLApiError('HTTP 401: {"message":"Unauthorized"}')

    monkeypatch.setattr(tl_client, '_http', fake)
    try:
        tl_client.fetch_details('id', 'do-not-log-secret', '8186')
    except tl_client.TLApiError as exc:
        assert 'Список броней (страница 1)' in str(exc)
        assert 'даже после выдачи нового токена' in str(exc)
        assert 'Read Reservation API' in str(exc)
        assert 'do-not-log-secret' not in str(exc)
    else:
        raise AssertionError('persisting HTTP 401 should raise a diagnostic error')


def test_booking_401_refreshes_token_and_retries(monkeypatch):
    import tl_client
    tokens = iter(['expired', 'fresh'])
    monkeypatch.setattr(tl_client, 'get_token', lambda *a: next(tokens))
    calls = []

    def fake(method, path, token=None, params=None, data=None, ctype=None):
        calls.append((path, token))
        if path.endswith('/bookings'):
            return 200, {'bookingSummaries': [{'number': 'n1'}],
                         'hasMoreData': False}
        if token == 'expired':
            raise tl_client.TLApiError('HTTP 401: {"message":"Unauthorized"}')
        return 200, {'booking': {'number': 'n1', 'status': 'Active'}}

    monkeypatch.setattr(tl_client, '_http', fake)
    monkeypatch.setattr(tl_client.time, 'sleep', lambda s: None)
    out = tl_client.fetch_details('id', 'secret', '8186')
    assert out == [{'number': 'n1', 'status': 'Active'}]
    assert [token for path, token in calls if path.endswith('/n1')] == [
        'expired', 'fresh']


def test_fetch_details_uses_requested_last_modification_date(monkeypatch):
    import tl_client
    monkeypatch.setattr(tl_client, 'get_token', lambda *a: 'tok')
    seen = []

    def fake(method, path, token=None, params=None, data=None, ctype=None):
        seen.append((path, params))
        return 200, {'bookingSummaries': [], 'hasMoreData': False}

    monkeypatch.setattr(tl_client, '_http', fake)
    out = tl_client.fetch_details('id', 'secret', '8186', '2026-07-01')
    assert out == []
    assert seen[0][1]['lastModification'] == '2026-07-01T00:00:00Z'


def test_import_cache_roundtrip_preserves_rows_and_report():
    from datetime import datetime
    from tl_client import decode_tl_import_cache, encode_tl_import_cache

    rows = [{
        'Гость': 'Иван Петров', 'Стоимость': 10000.0,
        'Заезд': datetime(2026, 7, 10, 14, 0),
        'Выезд': datetime(2026, 7, 12, 12, 0),
        'Источник': 'TravelLine YAX', '№ брони': 'TL-42',
    }]
    report = {'in_quarter': 1, 'rows': 1, 'cancelled': 0,
              'votes': {'A': 0, 'B': 0, '?': 0}, 'mismatch': []}

    payload = encode_tl_import_cache(rows, report)
    restored_rows, restored_report = decode_tl_import_cache(payload)
    assert restored_rows == rows
    assert isinstance(restored_rows[0]['Заезд'], datetime)
    assert isinstance(restored_rows[0]['Выезд'], datetime)
    assert restored_report == report


def test_import_cache_key_is_scoped_to_property_period_and_query():
    from tl_client import tl_import_cache_key

    common = ('tch', '8186', '2026-07-01', '2026-10-01', '2026-07-01')
    key = tl_import_cache_key(*common)
    assert key.startswith('tlcache-v1|')
    assert tl_import_cache_key(*common) == key
    assert tl_import_cache_key('tch', '479', *common[2:]) != key
    assert tl_import_cache_key('tch', '8186', '2026-10-01', '2027-01-01',
                               '2026-10-01') != key


def test_http_converts_read_timeout_to_diagnostic(monkeypatch):
    import socket
    import tl_client

    def timeout(*args, **kwargs):
        raise socket.timeout('The read operation timed out')

    monkeypatch.setattr(tl_client.urllib.request, 'urlopen', timeout)
    try:
        tl_client._http('GET', '/test')
    except tl_client.TLNetworkError as exc:
        assert 'Тайм-аут ответа TravelLine' in str(exc)
        assert '90 с' in str(exc)
    else:
        raise AssertionError('socket read timeout should be diagnosed')


def test_fetch_details_retries_a_transient_timeout(monkeypatch):
    import tl_client
    monkeypatch.setattr(tl_client, 'get_token', lambda *args: 'tok')
    monkeypatch.setattr(tl_client.time, 'sleep', lambda seconds: None)
    detail_calls = []

    def fake(method, path, token=None, params=None, data=None, ctype=None):
        if path.endswith('/bookings'):
            return 200, {'bookingSummaries': [{'number': 'n1'}],
                         'hasMoreData': False}
        detail_calls.append(path)
        if len(detail_calls) == 1:
            raise tl_client.TLNetworkError('Тайм-аут ответа TravelLine (90 с)')
        return 200, {'booking': {'number': 'n1', 'status': 'Active'}}

    monkeypatch.setattr(tl_client, '_http', fake)
    out = tl_client.fetch_details('id', 'secret', '8186')
    assert out == [{'number': 'n1', 'status': 'Active'}]
    assert len(detail_calls) == 2


def test_repeated_timeout_reports_api_stage_without_secret(monkeypatch):
    import tl_client
    monkeypatch.setattr(tl_client, 'get_token', lambda *args: 'tok')
    monkeypatch.setattr(tl_client.time, 'sleep', lambda seconds: None)

    def fake(method, path, token=None, params=None, data=None, ctype=None):
        if path.endswith('/bookings'):
            return 200, {'bookingSummaries': [{'number': 'n1'}],
                         'hasMoreData': False}
        raise tl_client.TLNetworkError('Тайм-аут ответа TravelLine (90 с)')

    monkeypatch.setattr(tl_client, '_http', fake)
    try:
        tl_client.fetch_details('id', 'do-not-log-secret', '8186')
    except tl_client.TLApiError as exc:
        assert 'Бронь n1' in str(exc)
        assert '2 попыток' in str(exc)
        assert 'do-not-log-secret' not in str(exc)
    else:
        raise AssertionError('repeated timeout should fail with a stage diagnostic')
