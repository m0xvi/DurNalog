# -*- coding: utf-8 -*-
"""Импорт броней и номерного фонда из TravelLine API.

Покрывает Чайковский (8186) и Дону (479): забирает брони за квартал,
раскладывает проживание/допуслуги/каналы и пишет Excel той же формы,
что ожидает обработка (канонический маппинг — диалог колонок не нужен).
Read Reservation возвращает roomStays[].roomId; для Чайковского он
сопоставляется с комнатами/корпусами из PMS API PropertyRoom. Если
проживание не назначено или его roomId отсутствует в номерном фонде, номер
и корпус остаются пустыми — идентификаторы API не подменяются номером комнаты.
"""
import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime

import openpyxl

HOST = 'https://partner.tlintegration.com'
HTTP_TIMEOUT_SECONDS = 90
NETWORK_RETRIES = 1
NETWORK_RETRY_DELAY = 2
# Increment if the prepared-row/report cache format changes incompatibly.
TL_CACHE_VERSION = 4

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
COL_SVC_NAME = 'Наименование допуслуг'
COL_SVC_EXTENSION = 'Сумма продления номера'
COL_TOTAL = 'Итог брони'
COL_ROOM = '№ комнаты'
COL_CORPUS = 'Корпус'
COL_BUILDING_NAME = 'Здание TravelLine'
ROOM_LABEL_TEXT_PREFIX = '__TL_ROOM_LABEL__:'

TL_DATA_MAP = {'guest': COL_GUEST, 'price': COL_PRICE, 'checkin': COL_IN,
               'checkout': COL_OUT, 'source': COL_SOURCE,
               'category': COL_CAT, 'bron': COL_BRON, 'nights': COL_NIGHTS,
               'room': COL_ROOM, 'corpus': COL_CORPUS}
TL_EXTRA = [COL_TARIFF, COL_SVC, COL_SVC_NAME, COL_SVC_EXTENSION, COL_TOTAL]
TL_ROOMS_MAP = {'fio': 'ФИО', 'room': COL_ROOM, 'corpus': COL_CORPUS,
                'building_name': COL_BUILDING_NAME}


class TLApiError(Exception):
    pass


class TLNetworkError(TLApiError):
    """Временная ошибка соединения/тайм-аут, которую можно повторить."""


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
        with urllib.request.urlopen(req, timeout=HTTP_TIMEOUT_SECONDS) as r:
            return r.status, json.loads(r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        try:
            detail = e.read().decode('utf-8', 'replace')[:300]
        except Exception:
            detail = ''
        raise TLApiError(f'HTTP {e.code}: {detail}')
    except (urllib.error.URLError, OSError) as e:
        reason = getattr(e, 'reason', e)
        reason_text = str(reason)
        if (isinstance(reason, TimeoutError)
                or 'timed out' in reason_text.lower()
                or 'time out' in reason_text.lower()):
            detail = f'Тайм-аут ответа TravelLine ({HTTP_TIMEOUT_SECONDS} с)'
        else:
            detail = f'Ошибка соединения с TravelLine: {reason_text}'
        raise TLNetworkError(detail) from e


def _http_with_retry(method, path, token=None, params=None, data=None,
                     ctype=None, cancel_event=None):
    """Повторяет один раз только временные сетевые ошибки, не HTTP-ошибки."""
    for attempt in range(NETWORK_RETRIES + 1):
        if cancel_event is not None and cancel_event.is_set():
            raise TLCancelled()
        try:
            return _http(method, path, token, params, data, ctype)
        except TLNetworkError as exc:
            if attempt >= NETWORK_RETRIES:
                tries = attempt + 1
                raise TLNetworkError(
                    f'{exc}; не удалось после {tries} попыток') from exc
            if cancel_event is not None:
                if cancel_event.wait(NETWORK_RETRY_DELAY):
                    raise TLCancelled()
            else:
                time.sleep(NETWORK_RETRY_DELAY)


def get_token(client_id, secret, cancel_event=None):
    try:
        _, data = _http_with_retry(
            'POST', '/auth/token',
            data=urllib.parse.urlencode(
                {'grant_type': 'client_credentials',
                 'client_id': client_id,
                 'client_secret': secret}),
            ctype='application/x-www-form-urlencoded',
            cancel_event=cancel_event)
        return data['access_token']
    except (TLApiError, KeyError) as e:
        raise TLApiError(f'Не удалось получить токен ({e}).')


def fetch_details(client_id, secret, property_id, depth='2024-01-01',
                  progress=None, cancel_event=None,
                  include_room_inventory=False):
    """Детали броней и, по запросу, инвентарь комнат/корпусов.

    Брони берутся с lastModification >= depth. При include_room_inventory
    дополнительно запрашивает PMS API PropertyRoom и возвращает пару
    ``(bookings, accommodationInventory response)``. progress(stage, done,
    total); HTTP 401 обновляет OAuth-токен один раз.
    Ошибки API содержат этап операции, но никогда не содержат client_secret.
    """
    def cancelled():
        return cancel_event is not None and cancel_event.is_set()

    if cancelled():
        raise TLCancelled()
    token = get_token(client_id, secret, cancel_event)

    def api_get(path, params=None, stage='TravelLine API'):
        """Выполнить GET; на 401 получить свежий token и повторить один раз."""
        nonlocal token

        def contextual(error):
            text = str(error)
            if isinstance(error, TLNetworkError):
                return TLApiError(
                    f'{stage}: {text}. Проверьте соединение с интернетом; '
                    'TravelLine мог временно не ответить, повторите запрос позже.')
            if 'HTTP 403' in text:
                if 'номерной фонд' in stage.casefold():
                    permission = 'чтение номерного фонда (PMS API/PropertyRoom)'
                else:
                    permission = 'Read Reservation API'
                return TLApiError(
                    f'{stage}: доступ запрещён (403). В подключении TravelLine '
                    f'проверьте разрешение {permission} и доступ этого объекта. '
                    f'Ответ API: {text}')
            if 'HTTP 401' in text:
                return TLApiError(
                    f'{stage}: TravelLine не принял токен (401 Unauthorized). '
                    'Проверьте client_id/client_secret именно из активного '
                    'подключения этого объекта и повторите проверку. '
                    f'Ответ API: {text}')
            return TLApiError(f'{stage}: {text}')

        try:
            return _http_with_retry(
                'GET', path, token, params, cancel_event=cancel_event)
        except TLApiError as first:
            if 'HTTP 429' in str(first):
                # Учитываем лимит TravelLine, но оставляем пользователю отмену.
                for _ in range(60):
                    if cancelled():
                        raise TLCancelled()
                    time.sleep(1)
                try:
                    return _http_with_retry(
                        'GET', path, token, params,
                        cancel_event=cancel_event)
                except TLApiError as retry_error:
                    raise contextual(retry_error)
            if 'HTTP 401' not in str(first):
                raise contextual(first)
            # Токен действует 15 минут. Обновляем его по 401 и повторяем
            # конкретный запрос один раз; не оставляем сырой ответ без этапа.
            try:
                token = get_token(client_id, secret, cancel_event)
            except TLApiError as auth_error:
                raise TLApiError(
                    f'{stage}: API вернул 401, а получить новый токен не удалось. '
                    f'{auth_error}')
            try:
                return _http_with_retry(
                    'GET', path, token, params, cancel_event=cancel_event)
            except TLApiError as retry_error:
                if 'HTTP 401' in str(retry_error):
                    permission = ('PMS API/PropertyRoom' if 'номерной фонд'
                                  in stage.casefold() else 'Read Reservation API')
                    raise TLApiError(
                        f'{stage}: HTTP 401 даже после выдачи нового токена. '
                        'TravelLine не принимает авторизацию этого API-запроса. '
                        'Проверьте пару client_id/client_secret из одного '
                        f'подключения и разрешение {permission}. '
                        f'Ответ API: {retry_error}')
                raise contextual(retry_error)

    inventory = None
    if include_room_inventory:
        if cancelled():
            raise TLCancelled()
        if progress:
            progress('inventory', 0, 1)
        _, inventory = api_get(
            f'/api/pms/v2/properties/{property_id}/obtain-accommodation-inventory',
            stage='Номерной фонд (комнаты и корпуса)')
        if not isinstance(inventory, dict):
            inventory = {}
        if cancelled():
            raise TLCancelled()
        if progress:
            progress('inventory', 1, 1)

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
    return (details, inventory) if include_room_inventory else details


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


def _room_number(value):
    """Номер из displayName; сохраняет буквенные номера, не roomId."""
    if value is None:
        return None
    text = str(value).strip()
    if not text or text.casefold() in ('none', 'null', 'nan'):
        return None
    try:
        number = float(text)
    except ValueError:
        return text
    if number.is_integer():
        # Room labels are display values; don't strip meaningful leading zeroes.
        return text if re.fullmatch(r'0\d+', text) else int(number)
    return text


def _building_number(building, ordinal):
    """Нормализует номер корпуса, используя имя API или порядок зданий.

    PMS хранит корпуса как произвольные названия. Если в названии есть
    явный номер («Корпус 2»), используем его; для имён вроде «Main building»
    остаётся порядковый номер в accommodationInventory.buildings.
    """
    name = str((building or {}).get('name') or '').strip()
    patterns = (
        r'(?:корпус|корп\.?|здание|строение|building|block)\D{0,8}(\d+)',
        r'(\d+)\s*(?:корпус|корп\.?|здание|строение|building|block)',
    )
    for pattern in patterns:
        match = re.search(pattern, name, flags=re.IGNORECASE)
        if match:
            try:
                return int(match.group(1))
            except ValueError:
                pass
    if re.fullmatch(r'\d+', name):
        return int(name)
    return ordinal


def accommodation_room_map(response):
    """TravelLine ``roomId`` -> (номер, корпус, исходное имя здания).

    PropertyRoom возвращает вложенную структуру buildings -> floors -> rooms;
    идентификатор комнаты сопоставляется с её displayName, а не используется
    как номер. Неизвестные/пустые поля остаются пустыми.
    """
    if not isinstance(response, dict):
        return {}
    inventory = response.get('accommodationInventory')
    if not isinstance(inventory, dict):
        inventory = response
    buildings = inventory.get('buildings')
    if not isinstance(buildings, list):
        return {}

    room_map = {}
    for building_index, building in enumerate(buildings, 1):
        if not isinstance(building, dict):
            continue
        corpus = _building_number(building, building_index)
        building_name = str(building.get('name') or '').strip()
        rooms = list(building.get('rooms') or [])
        for floor in building.get('floors') or []:
            if isinstance(floor, dict):
                rooms.extend(floor.get('rooms') or [])
        for room in rooms:
            if not isinstance(room, dict):
                continue
            room_id = room.get('id')
            if room_id is None or not str(room_id).strip():
                continue
            room_map.setdefault(str(room_id), {
                'room': _room_number(room.get('displayName')),
                'corpus': corpus,
                'building_name': building_name,
            })
    return room_map


def _service_amount_text(amount):
    return f'{amount:,.2f}'.replace(',', ' ').replace('.', ',')


def _service_name(service, fallback=''):
    if not isinstance(service, dict):
        return str(service or fallback or '').strip()
    name = next((str(service.get(key)).strip() for key in
                 ('name', 'displayName', 'serviceName', 'description',
                  'pmsSideCode', 'code', 'serviceCode')
                 if service.get(key)), '')
    return name or fallback or 'Дополнительная услуга'


def _service_description(service, fallback=''):
    """Читабельное название допуслуги с ценой, если API её передал."""
    name = _service_name(service, fallback)
    amount = _money(service.get('total')) if isinstance(service, dict) else 0.0
    if amount > 0:
        return f'{name} — {_service_amount_text(amount)} ₽'
    return name


def _is_extension_service(service):
    name = _service_name(service)
    normalized = ' '.join(name.casefold().replace('ё', 'е').split())
    return normalized == 'продление номера'


def _extension_amount(services):
    return sum(_money(service.get('total')) for service in services or []
               if isinstance(service, dict) and _is_extension_service(service))


def _service_descriptions(services):
    out = []
    for service in services or []:
        if not isinstance(service, dict):
            continue
        name = _service_description(service)
        if name:
            out.append(name)
    return out


def channel_label(source):
    """Источник брони человекочитаемо (идёт в колонку «Источник»)."""
    source = source or {}
    typ, code = source.get('type'), source.get('code')
    if typ == 'BookingEngine':
        return 'TravelLine сайт'
    if typ == 'PMS':
        return 'TravelLine ресепшен'
    if typ == 'Channel' and code:
        return f'TravelLine {code}'
    if typ or code:
        return f'TravelLine {typ or "?"}:{code or "?"}'
    return 'TravelLine ?'


def bookings_to_rows(details, qs, qe, accommodation_inventory=None):
    """Детали броней -> (строки Excel, отчёт).

    qs/qe — 'YYYY-MM-DD' границ квартала. Отменённые пропускаются.
    Арифметика каждой брони с платными допами самопроверяется:
    A: итог == проживание + допы; B: итог == проживание (допы внутри).
    Иначе бронь помечается расхождением (база = проживание как есть).
    При наличии инвентаря roomId сопоставляется с номером/корпусом; ID
    незанятого или неизвестного номера не переносится в отчёт как номер.
    """
    room_map = accommodation_room_map(accommodation_inventory)
    has_room_inventory = accommodation_inventory is not None
    rows, votes = [], {'A': 0, 'B': 0, '?': 0}
    rep = {'total': len(details), 'in_quarter': 0, 'cancelled': 0,
           'rows': 0, 'mismatch': [], 'channels': {}}
    if has_room_inventory:
        rep['room_assignments'] = {
            'inventory_rooms': len(room_map), 'assigned': 0, 'unassigned': 0}
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
            row_svc_extension = _extension_amount(rs.get('services') or [])
            row_svc_names = _service_descriptions(rs.get('services') or [])
            if first:
                booking_services = b.get('services') or []
                row_svc_extension += _extension_amount(booking_services)
                row_svc_names.extend(_service_descriptions(booking_services))
                for extra_rs in stays:
                    ex = extra_rs.get('extraStayCharges') or {}
                    for key, label in (('earlyArrival', 'Ранний заезд'),
                                       ('lateDeparture', 'Поздний выезд')):
                        charge = ex.get(key) or {}
                        if _is_extension_service(charge):
                            row_svc_extension += _money(charge.get('total'))
                        if _money(charge.get('total')) > 0:
                            row_svc_names.append(
                                _service_description(charge, label))
            if mode == 'B':
                base = stay_room - row_svc
                base = base if base > 0 else 0.0
            else:
                base = stay_room
            guests = ', '.join(_guest_name(g) for g in rs.get('guests') or []
                               if _guest_name(g))
            rt = rs.get('roomType') or {}
            rps = rs.get('ratePlans') or []
            room_id = rs.get('roomId')
            assignment = (room_map.get(str(room_id))
                          if room_id is not None and str(room_id).strip()
                          else None)
            room_number = assignment.get('room') if assignment else None
            corpus = assignment.get('corpus') if assignment else None
            building_name = assignment.get('building_name', '') if assignment else ''
            if has_room_inventory:
                bucket = 'assigned' if room_number is not None else 'unassigned'
                rep['room_assignments'][bucket] += 1
            rows.append({
                COL_GUEST: guests,
                COL_PRICE: round(base, 2),
                COL_IN: dt_in, COL_OUT: dt_out,
                COL_SOURCE: src,
                COL_CAT: rt.get('name') or '',
                COL_BRON: b.get('number') or '',
                COL_NIGHTS: nights,
                COL_ROOM: room_number,
                COL_CORPUS: corpus if room_number is not None else None,
                COL_BUILDING_NAME: building_name if room_number is not None else '',
                COL_TARIFF: ', '.join(r.get('name') or '' for r in rps),
                COL_SVC: round(row_svc, 2),
                COL_SVC_NAME: '; '.join(row_svc_names),
                COL_SVC_EXTENSION: round(row_svc_extension, 2),
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
            COL_BRON, COL_NIGHTS, COL_TARIFF, COL_SVC, COL_SVC_NAME,
            COL_SVC_EXTENSION, COL_TOTAL, COL_ROOM, COL_CORPUS]
    ws.append(cols)
    for r in rows:
        values = []
        for column in cols:
            value = r.get(column)
            if column == COL_ROOM and isinstance(value, str):
                # pandas may coerce text labels like "001" to the integer 1
                # when it later reads the workbook; preserve the exact API label.
                value = ROOM_LABEL_TEXT_PREFIX + value
            values.append(value)
        ws.append(values)
    for col in ('C', 'D'):
        for cell in ws[col][1:]:
            if isinstance(cell.value, datetime):
                cell.number_format = 'DD.MM.YYYY HH:MM'
    # Room assignment is consumed by the processor but kept out of the way
    # in the source workbook.
    for column in ('N', 'O'):
        ws.column_dimensions[column].hidden = True
    wb.save(path)


def write_rooms_workbook(rows, path):
    """Пишет доступные из API имя гостя, номер и корпус для проверки.

    Неприсвоенные и не найденные в номерном фонде roomId остаются пустыми.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Номера'
    ws.append(['ФИО', COL_ROOM, COL_CORPUS, COL_BUILDING_NAME])
    for row in rows or []:
        guests = [name.strip() for name in
                  re.split(r'[,;\n]+', str(row.get(COL_GUEST) or ''))
                  if name.strip()]
        if not guests:
            guests = ['']
        for guest in guests:
            ws.append([guest, row.get(COL_ROOM), row.get(COL_CORPUS),
                       row.get(COL_BUILDING_NAME) or ''])
    wb.save(path)


def write_rooms_stub(path):
    """Совместимая пустая заглушка для импорта отелей без номерного фонда."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Номера'
    ws.append(['ФИО', COL_ROOM])
    wb.save(path)
