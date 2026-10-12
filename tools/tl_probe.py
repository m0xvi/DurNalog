# -*- coding: utf-8 -*-
"""Зонд TravelLine Read Reservation API.

Что делает:
  1. Забирает брони вашего объекта за указанный квартал.
  2. Проверяет арифметику: входит ли сумма допуслуг в roomStays.total
     или идёт сверху (гипотезы A/B — критично для формулы турналога).
  3. Показывает разбивку по каналам/статусам и сохраняет выгрузку в JSON.

Только стандартная библиотека Python. Секреты никуда не отправляются,
кроме api TravelLine. Запуск:  python tools/tl_probe.py
"""
import getpass
import json
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from datetime import date

try:
    sys.stdout.reconfigure(encoding='utf-8')
except Exception:
    pass

HOST = 'https://partner.tlintegration.com'
QUARTERS = {'1': 1, 'i': 1, '2': 2, 'ii': 2, '3': 3, 'iii': 3, '4': 4, 'iv': 4}
VERDICT_STATUSES = ('Confirmed', 'InHouse', 'CheckedOut')


class ApiError(Exception):
    pass


def _http(method, path, token=None, params=None, data=None, ctype=None):
    url = HOST + path
    if params:
        url += '?' + urllib.parse.urlencode(params)
    body = data.encode('utf-8') if isinstance(data, str) else data
    req = urllib.request.Request(url, data=body, method=method)
    if token:
        req.add_header('Authorization', 'Bearer ' + token)
    if ctype:
        req.add_header('Content-Type', ctype)
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            return r.status, json.loads(r.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        try:
            detail = e.read().decode('utf-8', 'replace')[:500]
        except Exception:
            detail = ''
        raise ApiError(f'HTTP {e.code}: {detail}')


def get_token(client_id, secret):
    try:
        _, data = _http('POST', '/auth/token',
                        data=urllib.parse.urlencode(
                            {'grant_type': 'client_credentials',
                             'client_id': client_id,
                             'client_secret': secret}),
                        ctype='application/x-www-form-urlencoded')
        return data['access_token']
    except (ApiError, KeyError) as e:
        raise ApiError(f'Не удалось получить токен ({e}). '
                       'Проверьте client_id/secret и принятую оферту в кабинете.')


def list_summaries(token, pid, last_mod):
    """Все сводки броней постранично. Возвращает список."""
    out, cont, page = [], None, 0
    while True:
        page += 1
        if page > 200:
            raise ApiError('Слишком много страниц (>200) — уточните lastModification.')
        params = {'count': 1000}
        if cont:
            params['continueToken'] = cont
        else:
            params['lastModification'] = last_mod
        _, data = _http('GET',
                        f'/api/read-reservation/v1/properties/{pid}/bookings',
                        token, params)
        out.extend(data.get('bookingSummaries', []))
        print(f'  страница {page}: +{len(data.get("bookingSummaries", []))} '
              f'(всего {len(out)})')
        if not data.get('hasMoreData'):
            return out
        cont = data.get('continueToken')
        if not cont:
            return out
        time.sleep(0.4)


def get_details(token, pid, number):
    _, data = _http(
        'GET', f'/api/read-reservation/v1/properties/{pid}/bookings/{number}',
        token)
    return data.get('booking', {})


def money(total):
    try:
        return float((total or {}).get('priceAfterTax') or 0)
    except (TypeError, ValueError):
        return 0.0


def split_booking(b):
    """(проживание, допы_проживаний, допы_брони, ранний/поздний, итог_брони)."""
    room = svc_stay = svc_book = extra = 0.0
    for rs in b.get('roomStays', []) or []:
        room += money(rs.get('total'))
        for s in rs.get('services', []) or []:
            svc_stay += money(s.get('total'))
        ex = rs.get('extraStayCharges') or {}
        for k in ('earlyArrival', 'lateDeparture'):
            extra += money((ex.get(k) or {}).get('total'))
    for s in b.get('services', []) or []:
        svc_book += money(s.get('total'))
    return room, svc_stay, svc_book, extra, money(b.get('total'))


def overlaps(stays, qs, qe):
    for rs in stays:
        sd = rs.get('stayDates') or {}
        try:
            a = sd.get('arrivalDateTime', '')[:10]
            d = sd.get('departureDateTime', '')[:10]
            if a and d and a < qe.isoformat() and d > qs.isoformat():
                return True
        except Exception:
            continue
    return False


PII_KEYS = {'firstname', 'lastname', 'middlename', 'phone', 'email',
            'cardnumber', 'comment'}


def anonymize(obj):
    if isinstance(obj, dict):
        return {k: ('***' if k.lower() in PII_KEYS else anonymize(v))
                for k, v in obj.items()}
    if isinstance(obj, list):
        return [anonymize(v) for v in obj]
    return obj


def main():
    print('=== Зонд TravelLine: проживание vs допуслуги ===\n')
    print('Где взять данные:')
    print('  client_id / client_secret — кабинет TravelLine → Настройки гостиницы')
    print('                              → Подключения API → ваше подключение;')
    print('  propertyId — средняя часть номера любой брони (ГГГГММДД-XXXX-...).\n')
    cid = input('client_id: ').strip()
    try:
        secret = getpass.getpass('client_secret (ввод не виден): ').strip()
    except Exception:
        secret = input('client_secret: ').strip()
    pid = input('propertyId: ').strip()
    year = int(input('Год (например 2026): ').strip())
    q = QUARTERS.get(input('Квартал (1-4 или I-IV): ').strip().lower())
    if not q:
        print('Не понял квартал.')
        return
    qs = date(year, (q - 1) * 3 + 1, 1)
    qe = date(year + (1 if q == 4 else 0), 1 if q == 4 else q * 3 + 1, 1)
    depth = input('Искать изменения с даты [2024-01-01]: ').strip() or '2024-01-01'
    last_mod = depth + 'T00:00:00Z'
    print(f'\nКвартал: {qs} .. {qe} (по датам проживания)')

    print('\n[1/4] Токен…')
    token = get_token(cid, secret)
    print('  OK')

    print('[2/4] Список броней…')
    try:
        summaries = list_summaries(token, pid, last_mod)
    except ApiError as e:
        print(f'  ОШИБКА: {e}')
        if '403' in str(e):
            print('  Подсказка: примите оферту в кабинете и проверьте, что в подключении')
            print('  отмечен Read Reservation API именно для этого объекта.')
        return
    print(f'  Всего сводок: {len(summaries)}')

    print('[3/4] Детали броней (лимит API ~10/сек, это займёт время)…')
    details, unauth = [], 0
    for i, s in enumerate(summaries, 1):
        try:
            b = get_details(token, pid, s['number'])
        except ApiError as e:
            if '401' in str(e) and unauth == 0:
                unauth += 1
                token = get_token(cid, secret)  # токен живёт 15 минут
                b = get_details(token, pid, s['number'])
            elif '429' in str(e):
                print('  Лимит исчерпан, пауза 60 сек…')
                time.sleep(60)
                b = get_details(token, pid, s['number'])
            else:
                print(f'  пропуск {s["number"]}: {e}')
                continue
        details.append(b)
        if i % 25 == 0 or i == len(summaries):
            print(f'  {i}/{len(summaries)}')
        time.sleep(0.15)

    in_q = [b for b in details
            if overlaps(b.get('roomStays', []) or [], qs, qe)]
    print(f'\n[4/4] В квартал попадает броней: {len(in_q)} из {len(details)}')

    # --- вердикт A/B ---
    votes_a = votes_b = votes_q = 0
    examples = []
    for b in in_q:
        if b.get('status') not in VERDICT_STATUSES:
            continue
        room, s_stay, s_book, extra, tot = split_booking(b)
        if s_stay + s_book + extra < 0.01:
            continue  # без допов гипотезы неразличимы
        a = room + s_stay + s_book + extra
        if abs(tot - a) < 0.02:
            votes_a += 1
        elif abs(tot - room) < 0.02:
            votes_b += 1
        else:
            votes_q += 1
            if len(examples) < 5:
                examples.append((b.get('number'), room, s_stay + s_book,
                                 extra, tot))
    print('\n=== ВЕРДИКТ ===')
    print(f'Гипотеза A (допы СВЕРХУ проживания): {votes_a}')
    print(f'Гипотеза B (допы ВНУТРИ roomStays.total): {votes_b}')
    print(f'Не сошлось ни то ни другое: {votes_q}')
    if votes_a and not votes_b and not votes_q:
        print('>>> Формула: проживание = Σ roomStays.total (допы отдельно).')
    elif votes_b and not votes_a and not votes_q:
        print('>>> Формула: проживание = Σ roomStays.total − Σ services.')
    else:
        print('>>> Смешанная картина — покажите примеры ниже разработчику.')
    for n, r, sv, ex, t in examples:
        print(f'  {n}: room={r:.2f} svc={sv:.2f} extra={ex:.2f} total={t:.2f}')

    # --- каналы и статусы ---
    print('\n=== Каналы (все статусы, в квартале) ===')
    ch = {}
    for b in in_q:
        src = b.get('source') or {}
        code = f"{src.get('type', '?')}/{src.get('code', '?')}"
        room, s_stay, s_book, extra, tot = split_booking(b)
        c = ch.setdefault(code, [0, 0.0, 0.0])
        c[0] += 1
        c[1] += room
        c[2] += s_stay + s_book + extra
    for code, (n, r, sv) in sorted(ch.items()):
        print(f'  {code}: броней={n} проживание={r:.2f} допы={sv:.2f}')
    print('\n=== Статусы ===')
    st = {}
    for b in in_q:
        st[b.get('status', '?')] = st.get(b.get('status', '?'), 0) + 1
    for k, v in sorted(st.items()):
        print(f'  {k}: {v}')

    base = f'tl_{pid}_{year}Q{q}'
    with open(base + '_full.json', 'w', encoding='utf-8') as f:
        json.dump(details, f, ensure_ascii=False, indent=1)
    anon = [anonymize(b) for b in in_q]
    with open(base + '_anon.json', 'w', encoding='utf-8') as f:
        json.dump(anon, f, ensure_ascii=False, indent=1)
    print(f'\nСохранено: {base}_full.json (ВСЕ брони, есть ПДн — никому не отправлять),')
    print(f'           {base}_anon.json (квартал, ФИО скрыты — можно показать разработчику).')
    if anon:
        print('\nОбразец первой брони (обезличен):')
        print(json.dumps(anon[0], ensure_ascii=False, indent=1)[:2000])


if __name__ == '__main__':
    try:
        main()
    except ApiError as e:
        print(f'\nОШИБКА API: {e}')
    except KeyboardInterrupt:
        print('\\nПрервано.')
    except Exception as e:
        print(f'\nНеожиданная ошибка: {e}')
