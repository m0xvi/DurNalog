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
from collections import defaultdict
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
TL_CACHE_VERSION = 5

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
COL_TL_DISCOUNT = 'Скидка TravelLine, ₽'
COL_TL_DISCOUNT_PERCENT = 'Процент скидки TravelLine'
COL_TL_DISCOUNT_DATES = 'Предполагаемые ночи со скидкой TravelLine'
COL_TL_DAILY_SUM = 'Сумма dailyRates'
COL_TL_DAILY_CHECK = 'Сверка dailyRates'
ROOM_LABEL_TEXT_PREFIX = '__TL_ROOM_LABEL__:'

TL_DATA_MAP = {'guest': COL_GUEST, 'price': COL_PRICE, 'checkin': COL_IN,
               'checkout': COL_OUT, 'source': COL_SOURCE,
               'category': COL_CAT, 'bron': COL_BRON, 'nights': COL_NIGHTS,
               'room': COL_ROOM, 'corpus': COL_CORPUS}
TL_EXTRA = [COL_TARIFF, COL_SVC, COL_SVC_NAME, COL_SVC_EXTENSION, COL_TOTAL,
            COL_TL_DISCOUNT, COL_TL_DISCOUNT_PERCENT,
            COL_TL_DISCOUNT_DATES, COL_TL_DAILY_SUM, COL_TL_DAILY_CHECK]
TL_ROOMS_MAP = {'fio': 'ФИО', 'room': COL_ROOM, 'corpus': COL_CORPUS,
                'building_name': COL_BUILDING_NAME, 'bron': COL_BRON}


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


def _number(value):
    if value is None or isinstance(value, bool):
        return None
    try:
        return float(value)
    except (TypeError, ValueError, OverflowError):
        return None


def _discount_details(stay):
    """Скидка по проживанию и даты ночей, если их можно определить однозначно.

    API отдаёт скидки агрегированно в ``total.discounts``, а dailyRates — уже
    итоговую цену каждой ночи. Для одного процентного правила распределяем
    сумму на даты только когда одна группа одинаковых ночных цен точно
    объясняет общую сумму скидки. В остальных случаях показываем сумму и не
    угадываем даты.
    """
    total = stay.get('total') or {}
    discounts = total.get('discounts') or []
    if isinstance(discounts, dict):
        discounts = [discounts]
    discounts = [item for item in discounts if isinstance(item, dict)]
    if not discounts:
        return 0.0, '', ''

    amount = round(sum(_number(item.get('amount')) or 0.0
                       for item in discounts), 2)
    percent_labels = []
    for item in discounts:
        percent = _number(item.get('percent'))
        if percent is not None:
            label = f'{percent:g}%'
        elif str(item.get('discountType') or '').casefold() == 'fixed':
            label = 'фиксированная'
        else:
            label = str(item.get('discountType') or 'скидка')
        if label not in percent_labels:
            percent_labels.append(label)

    dates = []
    daily_rates = stay.get('dailyRates') or []
    if len(discounts) == 1:
        item = discounts[0]
        percent = _number(item.get('percent'))
        discount_amount = _number(item.get('amount'))
        kind = str(item.get('discountType') or '').casefold()
        if (kind == 'percentage' and percent is not None
                and 0 < percent < 100 and discount_amount is not None
                and discount_amount > 0 and isinstance(daily_rates, list)):
            by_price = defaultdict(list)
            complete = True
            for rate in daily_rates:
                if not isinstance(rate, dict):
                    complete = False
                    break
                price = _number(rate.get('priceBeforeTax'))
                date = str(rate.get('date') or '')[:10]
                if price is None or len(date) != 10:
                    complete = False
                    break
                by_price[round(price, 2)].append(date)
            if complete and by_price:
                candidates = []
                for net_price, rate_dates in by_price.items():
                    undiscounted = net_price / (1 - percent / 100)
                    per_night = round(undiscounted - net_price, 2)
                    explained = round(per_night * len(rate_dates), 2)
                    tolerance = max(0.02, 0.01 * len(rate_dates))
                    if abs(explained - discount_amount) <= tolerance:
                        candidates.append(rate_dates)
                if len(candidates) == 1:
                    dates = sorted(set(candidates[0]))

    if dates:
        date_text = ', '.join(dates)
    elif amount > 0:
        date_text = 'Не распределена по датам'
    else:
        date_text = ''
    return amount, '; '.join(percent_labels), date_text


def _daily_rate_audit(stay):
    """Сверяет сумму ночных тарифов с итогом проживания, не меняя его."""
    daily_rates = stay.get('dailyRates') or []
    if not isinstance(daily_rates, list) or not daily_rates:
        return None, 'Нет dailyRates'

    prices = []
    for rate in daily_rates:
        if not isinstance(rate, dict):
            return None, 'Неполные dailyRates'
        price = _number(rate.get('priceBeforeTax'))
        if price is None:
            return None, 'Неполные dailyRates'
        prices.append(price)
    daily_sum = round(sum(prices), 2)

    total = stay.get('total') or {}
    target_before_tax = _number(total.get('priceBeforeTax'))
    target_after_tax = _number(total.get('priceAfterTax'))
    # dailyRates[].priceBeforeTax — сравниваем с тем же налоговым базисом;
    # priceAfterTax используем только если API не прислал priceBeforeTax.
    target = (target_before_tax if target_before_tax is not None
              else target_after_tax)
    if target is None:
        return daily_sum, 'Нет итога для сверки'
    if abs(daily_sum - target) <= 0.02:
        return daily_sum, 'Сошлось по ночным тарифам'

    services = sum(_money(service.get('total'))
                   for service in stay.get('services') or []
                   if isinstance(service, dict))
    extras = stay.get('extraStayCharges') or {}
    if not isinstance(extras, dict):
        extras = {}
    extras = sum(_money(extras.get(key))
                 for key in ('earlyArrival', 'lateDeparture'))
    with_services = round(daily_sum + services + extras, 2)
    if ((services or extras)
            and abs(with_services - target) <= 0.02):
        return daily_sum, 'Сошлось с услугами'

    compared_sum = with_services if services or extras else daily_sum
    difference = round(compared_sum - target, 2)
    return daily_sum, f'Расхождение {difference:+.2f} ₽'


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


def _room_id_key(value):
    if isinstance(value, dict):
        nested_id = value.get('id')
        if nested_id is None or (isinstance(nested_id, str)
                                 and not nested_id.strip()):
            nested_id = value.get('roomId')
        value = nested_id
    return '' if value is None else str(value).strip()


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
            room_id = _room_id_key(room.get('id'))
            if not room_id:
                room_id = _room_id_key(room.get('roomId'))
            if not room_id:
                continue
            display_name = next((room.get(field) for field in
                                 ('displayName', 'roomNumber', 'number', 'name')
                                 if room.get(field) is not None
                                 and str(room.get(field)).strip()), None)
            room_map.setdefault(room_id, {
                'room': _room_number(display_name),
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
           'rows': 0, 'mismatch': [], 'channels': {},
           'discount_audit': {
               'stays_with_discount': 0, 'discount_amount': 0.0,
               'nights_with_discount_identified': 0,
               'discounts_without_night_attribution': 0,
               'daily_rates_checked': 0, 'daily_rates_matched': 0,
               'daily_rates_mismatched': 0, 'daily_rates_missing': 0,
           }}
    if has_room_inventory:
        rep['room_assignments'] = {
            'inventory_rooms': len(room_map), 'assigned': 0, 'unassigned': 0,
            'room_id_present': 0, 'room_id_missing': 0,
            'room_id_not_in_inventory': 0, 'inventory_no_display': 0,
            'unknown_room_id_samples': [],
        }
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
            discount_amount, discount_percent, discount_dates = (
                _discount_details(rs))
            daily_rate_sum, daily_rate_check = _daily_rate_audit(rs)
            discount_stats = rep['discount_audit']
            if discount_amount > 0:
                discount_stats['stays_with_discount'] += 1
                discount_stats['discount_amount'] += discount_amount
                if discount_dates == 'Не распределена по датам':
                    discount_stats['discounts_without_night_attribution'] += 1
                else:
                    discount_stats['nights_with_discount_identified'] += len(
                        discount_dates.split(', '))
            if daily_rate_sum is None or daily_rate_check.startswith(
                    ('Нет', 'Неполные')):
                discount_stats['daily_rates_missing'] += 1
            else:
                discount_stats['daily_rates_checked'] += 1
                if daily_rate_check.startswith('Сошлось'):
                    discount_stats['daily_rates_matched'] += 1
                else:
                    discount_stats['daily_rates_mismatched'] += 1
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
            room_key = _room_id_key(rs.get('roomId'))
            if not room_key:
                room_key = _room_id_key(rs.get('room'))
            assignment = room_map.get(room_key) if room_key else None
            room_number = assignment.get('room') if assignment else None
            corpus = assignment.get('corpus') if assignment else None
            building_name = assignment.get('building_name', '') if assignment else ''
            if has_room_inventory:
                room_stats = rep['room_assignments']
                room_stats['room_id_present' if room_key
                           else 'room_id_missing'] += 1
                if room_key and assignment is None:
                    room_stats['room_id_not_in_inventory'] += 1
                    samples = room_stats['unknown_room_id_samples']
                    if room_key not in samples and len(samples) < 20:
                        samples.append(room_key)
                elif assignment is not None and room_number is None:
                    room_stats['inventory_no_display'] += 1
                bucket = 'assigned' if room_number is not None else 'unassigned'
                room_stats[bucket] += 1
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
                COL_TL_DISCOUNT: discount_amount,
                COL_TL_DISCOUNT_PERCENT: discount_percent,
                COL_TL_DISCOUNT_DATES: discount_dates,
                COL_TL_DAILY_SUM: daily_rate_sum,
                COL_TL_DAILY_CHECK: daily_rate_check,
            })
            first = False
        ch[1] += room
        ch[2] += paid_extras
    rep['rows'] = len(rows)
    rep['votes'] = votes
    rep['discount_audit']['discount_amount'] = round(
        rep['discount_audit']['discount_amount'], 2)
    return rows, rep


def write_data_workbook(rows, path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Брони'
    cols = [COL_GUEST, COL_PRICE, COL_IN, COL_OUT, COL_SOURCE, COL_CAT,
            COL_BRON, COL_NIGHTS, COL_TARIFF, COL_SVC, COL_SVC_NAME,
            COL_SVC_EXTENSION, COL_TOTAL, COL_ROOM, COL_CORPUS,
            COL_TL_DISCOUNT, COL_TL_DISCOUNT_PERCENT,
            COL_TL_DISCOUNT_DATES, COL_TL_DAILY_SUM, COL_TL_DAILY_CHECK]
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
    for col in ('P', 'S'):
        for cell in ws[col][1:]:
            cell.number_format = '#,##0.00'
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
    ws.append(['ФИО', COL_ROOM, COL_CORPUS, COL_BUILDING_NAME, COL_BRON])
    for row in rows or []:
        guests = [name.strip() for name in
                  re.split(r'[,;\n]+', str(row.get(COL_GUEST) or ''))
                  if name.strip()]
        if not guests:
            guests = ['']
        for guest in guests:
            ws.append([guest, row.get(COL_ROOM), row.get(COL_CORPUS),
                       row.get(COL_BUILDING_NAME) or '', row.get(COL_BRON)])
    wb.save(path)


def tl_inspection_path(tmpdir, code, year, quarter):
    """Путь к локальной книге для просмотра выгрузки/диагностики TravelLine."""
    return f'{tmpdir}/TL_{code}_{year}{quarter}_inspection.xlsx'


def _inspection_json(value, limit=32000):
    try:
        text = json.dumps(value, ensure_ascii=False, default=str)
    except (TypeError, ValueError):
        text = str(value)
    if len(text) > limit:
        return text[:limit - 24] + ' … [обрезано]'
    return text


def _inspection_cell(value):
    if value is None:
        return ''
    if isinstance(value, (dict, list, tuple)):
        return _inspection_json(value)
    return value


def _inspection_redact(value):
    """Сохраняет техполя для отладки, убирая ФИО/контакты из сырого JSON."""
    private_parts = ('guest', 'name', 'email', 'phone', 'passport',
                     'document', 'contact')
    if isinstance(value, dict):
        return {
            key: ('<скрыто>' if any(part in str(key).casefold()
                                    for part in private_parts)
                  else _inspection_redact(item))
            for key, item in value.items()
        }
    if isinstance(value, (list, tuple)):
        return [_inspection_redact(item) for item in value]
    return value


def write_api_inspection_workbook(rows, report, path, details=None,
                                  accommodation_inventory=None):
    """Создаёт открытую пользователю книгу с экспортом и сырой диагностикой.

    В кэше TravelLine хранятся только подготовленные строки. Если ``details``
    отсутствует, книга честно помечает, что исходные API поля доступны только
    после выбора «Обновить из TravelLine».
    """
    wb = openpyxl.Workbook()
    summary = wb.active
    summary.title = 'Сводка'
    stats = (report or {}).get('room_assignments') or {}
    discount_stats = (report or {}).get('discount_audit') or {}

    def stat_value(key):
        return stats.get(key, '—' if details is None else 0)

    summary_rows = [
        ('Источник', 'Свежий ответ TravelLine' if details is not None
         else 'Локальный кэш (без исходных ответов API)'),
        ('Брони в ответе API', len(details) if details is not None else 'нет в кэше'),
        ('Подготовлено строк для приложения', len(rows or [])),
        ('Комнат в извлечённом API-инвентаре', stats.get('inventory_rooms', 0)),
        ('Проживаний сопоставлено с номером', stats.get('assigned', 0)),
        ('Проживаний без номера', stats.get('unassigned', 0)),
        ('Проживаний с roomId/room', stat_value('room_id_present')),
        ('Проживаний без roomId/room', stat_value('room_id_missing')),
        ('ID не найден в инвентаре', stat_value('room_id_not_in_inventory')),
        ('Комната найдена, но нет отображаемого номера',
         stat_value('inventory_no_display')),
        ('Проживаний со скидкой TravelLine',
         discount_stats.get('stays_with_discount', 0)),
        ('Общая скидка TravelLine, ₽',
         discount_stats.get('discount_amount', 0.0)),
        ('Предполагаемых ночей со скидкой (вывод по dailyRates)',
         discount_stats.get('nights_with_discount_identified', 0)),
        ('Скидок без однозначной даты',
         discount_stats.get('discounts_without_night_attribution', 0)),
        ('Сверок dailyRates совпало',
         discount_stats.get('daily_rates_matched', 0)),
        ('Сверок dailyRates с расхождением',
         discount_stats.get('daily_rates_mismatched', 0)),
        ('Сверок dailyRates без данных',
         discount_stats.get('daily_rates_missing', 0)),
    ]
    if details is None:
        summary_rows.append((
            'Важно', 'Кэш содержит только готовые строки; выберите '
            '«Обновить из TravelLine», чтобы получить исходные roomId и инвентарь.'))
    for row_index, item in enumerate(summary_rows, 1):
        summary.append(item)
        if item[0] == 'Общая скидка TravelLine, ₽':
            summary.cell(row=row_index, column=2).number_format = '#,##0.00'
    if stats.get('unknown_room_id_samples'):
        summary.append(('Примеры ID без совпадения',
                        ', '.join(map(str, stats['unknown_room_id_samples']))))
    summary.column_dimensions['A'].width = 48
    summary.column_dimensions['B'].width = 100
    summary.freeze_panes = 'A1'

    data_sheet = wb.create_sheet('Данные')
    data_columns = [COL_GUEST, COL_PRICE, COL_IN, COL_OUT, COL_SOURCE, COL_CAT,
                    COL_BRON, COL_NIGHTS, COL_TARIFF, COL_SVC, COL_SVC_NAME,
                    COL_SVC_EXTENSION, COL_TOTAL, COL_ROOM, COL_CORPUS,
                    COL_BUILDING_NAME, COL_TL_DISCOUNT,
                    COL_TL_DISCOUNT_PERCENT, COL_TL_DISCOUNT_DATES,
                    COL_TL_DAILY_SUM, COL_TL_DAILY_CHECK]
    data_sheet.append(data_columns)
    for row in rows or []:
        data_sheet.append([row.get(column) for column in data_columns])
    for col in ('Q', 'T'):
        for cell in data_sheet[col][1:]:
            cell.number_format = '#,##0.00'

    rooms_sheet = wb.create_sheet('Номера')
    rooms_sheet.append(['ФИО', COL_ROOM, COL_CORPUS,
                        COL_BUILDING_NAME, COL_BRON])
    for row in rows or []:
        guests = [name.strip() for name in
                  re.split(r'[,;\n]+', str(row.get(COL_GUEST) or ''))
                  if name.strip()]
        if not guests:
            guests = ['']
        for guest in guests:
            rooms_sheet.append([guest, row.get(COL_ROOM), row.get(COL_CORPUS),
                                row.get(COL_BUILDING_NAME) or '',
                                row.get(COL_BRON)])

    inventory_sheet = wb.create_sheet('Номерной фонд API')
    inventory_headers = [
        '№ корпуса', 'ID корпуса', 'Корпус API', 'ID этажа', 'Этаж API',
        'room.id', 'room.roomId', 'displayName', 'roomNumber', 'number', 'name',
        'Поля комнаты (JSON)',
    ]
    inventory_sheet.append(inventory_headers)
    inventory = accommodation_inventory
    if isinstance(inventory, dict) and isinstance(
            inventory.get('accommodationInventory'), dict):
        inventory = inventory['accommodationInventory']
    buildings = inventory.get('buildings') if isinstance(inventory, dict) else None
    if isinstance(buildings, list):
        for building_index, building in enumerate(buildings, 1):
            if not isinstance(building, dict):
                continue
            room_entries = [(None, room) for room in (building.get('rooms') or [])]
            for floor in building.get('floors') or []:
                if isinstance(floor, dict):
                    room_entries.extend((floor, room)
                                        for room in (floor.get('rooms') or []))
            for floor, room in room_entries:
                if not isinstance(room, dict):
                    continue
                inventory_sheet.append([
                    _building_number(building, building_index),
                    _inspection_cell(building.get('id')),
                    building.get('name') or '',
                    _inspection_cell(floor.get('id')) if floor else '',
                    floor.get('name') or '' if floor else '',
                    _inspection_cell(room.get('id')),
                    _inspection_cell(room.get('roomId')),
                    _inspection_cell(room.get('displayName')),
                    _inspection_cell(room.get('roomNumber')),
                    _inspection_cell(room.get('number')),
                    _inspection_cell(room.get('name')),
                    _inspection_json(_inspection_redact(room)),
                ])
    if inventory_sheet.max_row == 1:
        inventory_sheet.append(['Нет сырого номерного фонда в этой выгрузке.'])

    stays_sheet = wb.create_sheet('Проживания API')
    stays_sheet.append([
        '№ брони', '№ проживания', 'pmsRoomStayId', 'roomId', 'room',
        'roomTypeId', 'roomType', 'Заезд API', 'Выезд API', 'Все ключи roomStay',
        'Поля размещения (JSON)', 'roomId использован из',
        'Результат сопоставления', 'Номер из инвентаря', 'Корпус',
    ])
    room_map = accommodation_room_map(accommodation_inventory)
    if details is not None:
        for booking in details or []:
            if not isinstance(booking, dict):
                continue
            booking_number = booking.get('number') or ''
            for index, stay in enumerate(booking.get('roomStays') or [], 1):
                if not isinstance(stay, dict):
                    stays_sheet.append([booking_number, index, '', '', '', '', '',
                                        '', '', '', _inspection_json(stay), '',
                                        'roomStay не является объектом', '', ''])
                    continue
                raw_room_id = stay.get('roomId')
                raw_room = stay.get('room')
                room_key = _room_id_key(raw_room_id)
                room_id_source = 'roomId'
                if not room_key:
                    room_key = _room_id_key(raw_room)
                    room_id_source = 'room' if room_key else ''
                assignment = room_map.get(room_key) if room_key else None
                if not room_key:
                    match_status = 'roomId/room отсутствует'
                elif assignment is None:
                    match_status = 'ID не найден в инвентаре'
                elif assignment.get('room') is None:
                    match_status = 'ID найден, display name отсутствует'
                else:
                    match_status = 'Сопоставлено'
                room_type = stay.get('roomType') or {}
                if isinstance(room_type, dict):
                    room_type = room_type.get('name') or room_type.get('displayName') \
                        or _inspection_json(room_type)
                stay_dates = stay.get('stayDates') or {}
                safe_stay = _inspection_redact(stay)
                stays_sheet.append([
                    booking_number, index,
                    _inspection_cell(stay.get('pmsRoomStayId')),
                    _inspection_cell(raw_room_id), _inspection_cell(raw_room),
                    _inspection_cell(stay.get('roomTypeId')),
                    _inspection_cell(room_type),
                    _inspection_cell(stay_dates.get('arrivalDateTime'))
                    if isinstance(stay_dates, dict) else '',
                    _inspection_cell(stay_dates.get('departureDateTime'))
                    if isinstance(stay_dates, dict) else '',
                    ', '.join(map(str, stay.keys())),
                    _inspection_json(safe_stay), room_id_source, match_status,
                    assignment.get('room') if assignment else None,
                    assignment.get('corpus') if assignment else None,
                ])
    if details is None:
        stays_sheet.append([
            'Сырые проживания не сохранены в кэше. Нажмите «Обновить из TravelLine».'])
    elif stays_sheet.max_row == 1:
        stays_sheet.append(['В API-ответе нет roomStays.'])

    default_widths = {
        'Данные': [26, 16, 20, 20, 26, 28, 26, 10, 28, 14, 48, 20,
                   16, 16, 12, 28, 18, 24, 38, 16, 34],
        'Номера': [32, 16, 12, 28, 28],
        'Номерной фонд API': [12, 18, 28, 18, 24, 28, 28, 22, 22, 18, 26, 72],
        'Проживания API': [28, 12, 28, 32, 32, 24, 32, 24, 24, 54, 72, 18, 36, 20, 14],
    }
    for worksheet in wb.worksheets:
        worksheet.freeze_panes = 'A2' if worksheet.max_row > 1 else 'A1'
        if worksheet.max_row > 1 and worksheet.title != 'Сводка':
            worksheet.auto_filter.ref = worksheet.dimensions
        for index, width in enumerate(default_widths.get(worksheet.title, []), 1):
            worksheet.column_dimensions[openpyxl.utils.get_column_letter(index)].width = width
    wb.save(path)


def write_rooms_stub(path):
    """Совместимая пустая заглушка для импорта отелей без номерного фонда."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Номера'
    ws.append(['ФИО', COL_ROOM])
    wb.save(path)
