# -*- coding: utf-8 -*-
"""Импорт броней из TravelLine Read Reservation API.

Покрывает Чайковский (8186) и Дону (479): забирает брони за квартал,
раскладывает проживание/допуслуги/каналы и пишет Excel той же формы,
что ожидает обработка (канонический маппинг — диалог колонок не нужен).
Номера API не отдаёт (Read Reservation) — пишется пустая заглушка «Номера»,
все строки попадают на лист «нет корпуса»; на налог это не влияет.
"""
import json
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

import openpyxl

HOST = 'https://partner.tlintegration.com'
# Increment if the prepared-row/report cache format changes incompatibly.
TL_CACHE_VERSION = 1

# Канонические колонки выгрузки (имена пересекаются с алиасами схем,
# но маппинг задаётся явно — автоподбор не нужен).
COL_GUEST = 'Гость'
COL_PRICE = 'Стоимость'
COL_IN = 'Заезд'
COL_OUT = 'Выезд'
COL_SOURCE = 'Источник'
COL_CAT = 'Категория номера'
COL_BRON = '№ брони'
COL_NIGHTS = 'Ночей'
COL_TARIFF = 'Тариф'
COL_SVC = 'Допуслуги'
COL_TOTAL = 'Итог брони'

TL_DATA_MAP = {'guest': COL_GUEST, 'price': COL_PRICE, 'checkin': COL_IN,
               'checkout': COL_OUT, 'source': COL_SOURCE,
               'category': COL_CAT, 'bron': COL_BRON, 'nights': COL_NIGHTS}
TL_EXTRA = [COL_TARIFF, COL_SVC, COL_TOTAL]
TL_ROOMS_MAP = {'fio': 'ФИО', 'room': '№ комнаты'}


class TLApiError(Exception):
    pass


class TLCancelled(Exception):
    """Загрузка отменена пользователем."""


def is_tl_export(path):
    """Файл, сгенерированный импортом TravelLine (канонический маппинг)."""
    try:
        base = str(path).replace('\\', '/').rsplit('/', 1)[-1]
    except Exception:
        return False
    return base.startswith('TL_') and (
        base.endswith('_data.xlsx') or base.endswith('_rooms.xlsx'))


def tl_export_paths(tmpdir, code, year, quarter):
    stem = f'TL_{code}_{year}{quarter}_'
    return (tmpdir + '/' + stem + 'data.xlsx',
            tmpdir + '/' + stem + 'rooms.xlsx')


def tl_import_cache_key(hotel_code, property_id, period_start, period_end,
                        modified_since):
    return (f'tlcache-v{TL_CACHE_VERSION}|{hotel_code}|{property_id}|'
            f'{period_start}|{period_end}|{modified_since}')


def encode_tl_import_cache(rows, report):
    """Сериализует только строки, нужные импорту (без сырого API JSON/контактов)."""
    def default(value):
        if hasattr(value, 'isoformat'):
            return value.isoformat()
        return str(value)
    return json.dumps({'rows': rows, 'report': report}, ensure_ascii=False,
                      default=default)


def decode_tl_import_cache(payload):
    """Восстанавливает строки кэша, включая Excel-совместимые даты."""
    data = json.loads(payload)
    rows, report = data['rows'], data['report']
    for row in rows:
        for col in (COL_IN, COL_OUT):
            value = row.get(col)
            if isinstance(value, str) and value:
                try:
                    row[col] = datetime.fromisoformat(value)
                except ValueError:
                    pass
    return rows, report


def _http(method, path, token=None, params=None, data=None, ctype=None):
    url = HOST + path
    if params:
        url += '?' + urllib.parse.urlencode(params)
    body = data.encode('utf-8') if isinstance(data, str) else data
    req = urllib.request.Request(url, data=data if body is None else body,
                                 method=method)
    if token:
        req.add_header('Authorization', 'Bearer ' + token)
    if ctype:
        req.add_header('Content-Type', ctype)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.status, json.loads(r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        try:
            detail = e.read().decode('utf-8', 'replace')[:300]
        except Exception:
            detail = ''
        raise TLApiError(f'HTTP {e.code}: {detail}')


def get_token(client_id, secret):
    try:
        _, data = _http('POST', '/auth/token',
                        data=urllib.parse.urlencode(
                            {'grant_type': 'client_credentials',
                             'client_id': client_id,
                             'client_secret': secret}),
                        ctype='application/x-www-form-urlencoded')
        return data['access_token']
    except (TLApiError, KeyError) as e:
        raise TLApiError(f'Не удалось получить токен ({e}).')


def fetch_details(client_id, secret, property_id, depth='2024-01-01',
                  progress=None, cancel_event=None):
    """Все детали броней, изменённых с depth. progress(stage, done, total).

    При HTTP 401 один раз перевыпускает OAuth-токен и повторяет запрос.
    Ошибки API содержат этап операции, но никогда не содержат client_secret.
    """
    def cancelled():
        return cancel_event is not None and cancel_event.is_set()

    if cancelled():
        raise TLCancelled()
    token = get_token(client_id, secret)

    def api_get(path, params=None, stage='TravelLine API'):
        """Выполнить GET; на 401 получить свежий token и повторить один раз."""
        nonlocal token

        def contextual(error):
            text = str(error)
            if 'HTTP 403' in text:
                return TLApiError(
                    f'{stage}: доступ запрещён (403). В подключении TravelLine '
                    'проверьте разрешение Read Reservation API и доступ этого '
                    f'объекта. Ответ API: {text}')
            if 'HTTP 401' in text:
                return TLApiError(
                    f'{stage}: TravelLine не принял токен (401 Unauthorized). '
                    'Проверьте client_id/client_secret именно из активного '
                    'подключения этого объекта и повторите проверку. '
                    f'Ответ API: {text}')
            return TLApiError(f'{stage}: {text}')

        try:
            return _http('GET', path, token, params)
        except TLApiError as first:
            if 'HTTP 429' in str(first):
                # Учитываем лимит TravelLine, но оставляем пользователю отмену.
                for _ in range(60):
                    if cancelled():
                        raise TLCancelled()
                    time.sleep(1)
                try:
                    return _http('GET', path, token, params)
                except TLApiError as retry_error:
                    raise contextual(retry_error)
            if 'HTTP 401' not in str(first):
                raise contextual(first)
            # Токен действует 15 минут. Обновляем его по 401 и повторяем
            # конкретный запрос один раз; не оставляем сырой ответ без этапа.
            try:
                token = get_token(client_id, secret)
            except TLApiError as auth_error:
                raise TLApiError(
                    f'{stage}: API вернул 401, а получить новый токен не удалось. '
                    f'{auth_error}')
            try:
                return _http('GET', path, token, params)
            except TLApiError as retry_error:
                if 'HTTP 401' in str(retry_error):
                    raise TLApiError(
                        f'{stage}: HTTP 401 даже после выдачи нового токена. '
                        'TravelLine не принимает авторизацию этого API-запроса. '
                        'Проверьте пару client_id/client_secret из одного '
                        'подключения и разрешение Read Reservation API. '
                        f'Ответ API: {retry_error}')
                raise contextual(retry_error)

    summaries, cont, page = [], None, 0
    while True:
        if cancelled():
            raise TLCancelled()
        page += 1
        if page > 200:
            raise TLApiError('Слишком много страниц списка броней.')
        params = {'count': 1000}
        if cont:
            params['continueToken'] = cont
        else:
            params['lastModification'] = depth + 'T00:00:00Z'
        _, data = api_get(
            f'/api/read-reservation/v1/properties/{property_id}/bookings',
            params, f'Список броней (страница {page})')
        if cancelled():
            raise TLCancelled()
        summaries.extend(data.get('bookingSummaries', []))
        if progress:
            progress('list', len(summaries), 0)
        if not data.get('hasMoreData'):
            break
        cont = data.get('continueToken')
        if not cont:
            break
        time.sleep(0.4)
    details = []
    total = len(summaries)
    for i, summary in enumerate(summaries, 1):
        if cancelled():
            raise TLCancelled()
        _, data = api_get(
            f'/api/read-reservation/v1/properties/{property_id}'
            f'/bookings/{summary["number"]}',
            stage=f'Бронь {summary["number"]}')
        details.append(data.get('booking', {}))
        if progress:
            progress('details', i, total)
        # Read Reservation limit for details: 200 requests/minute.
        time.sleep(0.31)
    return details


def _money(total):
    try:
        return float((total or {}).get('priceAfterTax') or 0)
    except (TypeError, ValueError):
        return 0.0


def _stay_overlaps(stay_dates, qs, qe):
    try:
        a = (stay_dates.get('arrivalDateTime') or '')[:10]
        d = (stay_dates.get('departureDateTime') or '')[:10]
        return bool(a and d and a < qe and d > qs)
    except Exception:
        return False


def _parse_dt(s):
    try:
        return datetime.strptime(s[:16], '%Y-%m-%dT%H:%M')
    except (TypeError, ValueError):
        return None


def _guest_name(g):
    return ' '.join(x for x in (g.get('firstName'), g.get('lastName'))
                    if x).strip()


def channel_label(source):
    """Источник брони человекочитаемо (идёт в колонку «Источник»)."""
    source = source or {}
    typ, code = source.get('type'), source.get('code')
    if typ == 'BookingEngine':
        return 'TravelLine сайт'
    if typ == 'PMS':
        return 'TravelLine ресепшн'
    if typ == 'Channel' and code:
        return f'TravelLine {code}'
    if typ or code:
        return f'TravelLine {typ or "?"}:{code or "?"}'
    return 'TravelLine ?'


def bookings_to_rows(details, qs, qe):
    """Детали броней -> (строки Excel, отчёт).

    qs/qe — 'YYYY-MM-DD' границ квартала. Отменённые пропускаются.
    Арифметика каждой брони с платными допами самопроверяется:
    A: итог == проживание + допы; B: итог == проживание (допы внутри).
    Иначе бронь помечается расхождением (база = проживание как есть).
    """
    rows, votes = [], {'A': 0, 'B': 0, '?': 0}
    rep = {'total': len(details), 'in_quarter': 0, 'cancelled': 0,
           'rows': 0, 'mismatch': [], 'channels': {}}
    for b in details or []:
        if b.get('status') == 'Cancelled':
            rep['cancelled'] += 1
            continue
        stays = b.get('roomStays') or []
        room = svc_stay = svc_book = extra = 0.0
        for rs in stays:
            room += _money(rs.get('total'))
            for s in rs.get('services') or []:
                svc_stay += _money(s.get('total'))
            ex = rs.get('extraStayCharges') or {}
            for k in ('earlyArrival', 'lateDeparture'):
                extra += _money((ex.get(k) or {}).get('total'))
        for s in b.get('services') or []:
            svc_book += _money(s.get('total'))
        tot = _money(b.get('total'))
        paid_extras = svc_stay + svc_book + extra
        if paid_extras < 0.01:
            mode = 'A'
        elif abs(tot - (room + paid_extras)) < 0.02:
            mode = 'A'
            votes['A'] += 1
        elif abs(tot - room) < 0.02:
            mode = 'B'
            votes['B'] += 1
        else:
            mode = '?'
            votes['?'] += 1
            rep['mismatch'].append(
                {'number': b.get('number'), 'room': round(room, 2),
                 'svc': round(svc_stay + svc_book, 2),
                 'extra': round(extra, 2), 'total': round(tot, 2)})
        hit = [rs for rs in stays
               if _stay_overlaps(rs.get('stayDates') or {}, qs, qe)]
        if not hit:
            continue
        rep['in_quarter'] += 1
        src = channel_label(b.get('source'))
        ch = rep['channels'].setdefault(src, [0, 0.0, 0.0])
        ch[0] += 1
        first = True
        for rs in hit:
            sd = rs.get('stayDates') or {}
            dt_in = _parse_dt(sd.get('arrivalDateTime') or '')
            dt_out = _parse_dt(sd.get('departureDateTime') or '')
            if dt_in and dt_out:
                nights = max(1, (dt_out.date() - dt_in.date()).days)
            else:
                nights = 1
            stay_room = _money(rs.get('total'))
            stay_svc = sum(_money(s.get('total'))
                           for s in rs.get('services') or [])
            row_svc = stay_svc + (svc_book + extra if first else 0.0)
            if mode == 'B':
                base = stay_room - row_svc
                base = base if base > 0 else 0.0
            else:
                base = stay_room
            guests = ', '.join(_guest_name(g) for g in rs.get('guests') or []
                               if _guest_name(g))
            rt = rs.get('roomType') or {}
            rps = rs.get('ratePlans') or []
            rows.append({
                COL_GUEST: guests,
                COL_PRICE: round(base, 2),
                COL_IN: dt_in, COL_OUT: dt_out,
                COL_SOURCE: src,
                COL_CAT: rt.get('name') or '',
                COL_BRON: b.get('number') or '',
                COL_NIGHTS: nights,
                COL_TARIFF: ', '.join(r.get('name') or '' for r in rps),
                COL_SVC: round(row_svc, 2),
                COL_TOTAL: round(tot, 2),
            })
            first = False
        ch[1] += room
        ch[2] += paid_extras
    rep['rows'] = len(rows)
    rep['votes'] = votes
    return rows, rep


def write_data_workbook(rows, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Брони'
    cols = [COL_GUEST, COL_PRICE, COL_IN, COL_OUT, COL_SOURCE, COL_CAT,
            COL_BRON, COL_NIGHTS, COL_TARIFF, COL_SVC, COL_TOTAL]
    ws.append(cols)
    for r in rows:
        ws.append([r.get(c) for c in cols])
    for col in ('C', 'D'):
        for cell in ws[col][1:]:
            if isinstance(cell.value, datetime):
                cell.number_format = 'DD.MM.YYYY HH:MM'
    wb.save(path)


def write_rooms_stub(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Номера'
    ws.append(['ФИО', '№ комнаты'])
    wb.save(path)
