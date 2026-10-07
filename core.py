# -*- coding: utf-8 -*-
"""Боевая логика Турналога (без GUI): обработка, история, сопоставления."""
import io
import logging
import os
import re
import subprocess
import sys
import zipfile
from datetime import datetime

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill

from agents import default_agents, match_agent, norm_key
from schemas import DNA_DATA_SCHEMA, TCH_DATA_SCHEMA, TCH_ROOMS_SCHEMA
from tax import TaxParams, parse_price

log = logging.getLogger('turnalog')

FILL_GUEST = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
FILL_DATE = PatternFill(start_color="FFCC99", end_color="FFCC99", fill_type="solid")
FILL_PRICE = PatternFill(start_color="99CCFF", end_color="99CCFF", fill_type="solid")
FILL_ROOM = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
FILL_DUP = PatternFill(start_color="CC99FF", end_color="CC99FF", fill_type="solid")
FILL_TOTALS = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
DATE_FMT = 'DD.MM.YYYY HH:MM'

def norm_name(s):
    """Нормализация ФИО: регистр, ё->е, лишние пробелы."""
    return re.sub(r'\s+', ' ', str(s).lower().replace('ё', 'е')).strip()


def clean_guest_names(text):
    if pd.isna(text):
        return ""
    return re.sub(r'\d+\.\s*', '', str(text)).strip()


def read_table(path, nrows=None):
    """Чтение .xlsx/.csv с автоподбором кодировки для CSV."""
    if str(path).lower().endswith('.csv'):
        last_err = None
        for enc in ('utf-8-sig', 'utf-8', 'cp1251', 'cp866'):
            try:
                return pd.read_csv(path, encoding=enc, nrows=nrows)
            except UnicodeDecodeError as e:
                last_err = e
        raise ValueError(f"Не удалось определить кодировку CSV: {last_err}")
    return pd.read_excel(path, nrows=nrows)


def read_headers(path):
    return [str(c) for c in read_table(path, nrows=0).columns]


def guess_mapping(headers, schema):
    """Автоподбор колонок файла под поля приложения (точное, затем по подстроке)."""
    nh = {h: norm_key(h) for h in headers}
    taken, mapping = set(), {}
    for f in schema:  # pass 1: точное совпадение синонима
        for h, k in nh.items():
            if h in taken:
                continue
            if any(norm_key(a) == k for a in f['aliases']):
                mapping[f['key']] = h
                taken.add(h)
                break
    for f in schema:  # pass 2: подстрока (самый длинный синоним)
        if f['key'] in mapping:
            continue
        best, bestlen = None, 0
        for h, k in nh.items():
            if h in taken or not k:
                continue
            for a in f['aliases']:
                na = norm_key(a)
                if na and (na in k or k in na) and len(na) > bestlen:
                    best, bestlen = h, len(na)
        if best:
            mapping[f['key']] = best
            taken.add(best)
    return mapping


def mapping_valid(mapping, headers, schema):
    if not mapping:
        return False
    hs = set(headers)
    return all(mapping.get(f['key']) in hs for f in schema if f['required'])


def open_in_system(path):
    try:
        if os.name == 'nt':
            os.startfile(path)
        elif sys.platform == 'darwin':
            subprocess.Popen(['open', path])
        else:
            subprocess.Popen(['xdg-open', path])
    except Exception as e:
        log.exception("open_in_system")
        try:
            from tkinter import messagebox  # лениво: core импортируется и без GUI
            messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{e}")
        except Exception:
            pass


def pack_originals(paths):
    if len(paths) == 1:
        p = paths[0]
        return os.path.basename(p), open(p, 'rb').read()
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
        for p in paths:
            z.write(p, os.path.basename(p))
    return 'исходные_файлы.zip', buf.getvalue()



def process_tchaikovsky( data_path, rooms_path, output_path,
                        map_data=None, map_rooms=None,
                        extra_data=None, dup_checker=None, tax=None,
                        agents=None, source_map=None):
    """Возвращает (errors_summary, bookings)."""
    tp = TaxParams(tax)
    log.info("Чайковский: %s / %s", data_path, rooms_path)
    df_data = read_table(data_path)
    df_rooms = read_table(rooms_path)

    headers_d = [str(c) for c in df_data.columns]
    headers_r = [str(c) for c in df_rooms.columns]
    map_data = map_data or guess_mapping(headers_d, TCH_DATA_SCHEMA)
    map_rooms = map_rooms or guess_mapping(headers_r, TCH_ROOMS_SCHEMA)
    miss = [f['label'] for f in TCH_DATA_SCHEMA if f['required'] and not map_data.get(f['key'])]
    if miss:
        raise ValueError(f"Не сопоставлены колонки: {', '.join(miss)}")
    if not (map_rooms.get('fio') and map_rooms.get('room')):
        raise ValueError("Не сопоставлены колонки файла «Номера» (ФИО, № комнаты)")
    extra_data = [c for c in (extra_data or []) if c in df_data.columns]

    room_mapping = {}
    for _, rr in df_rooms.iterrows():
        fio = rr[map_rooms['fio']]
        if pd.isna(fio):
            continue
        key = norm_name(fio)
        if key and key not in room_mapping:
            room_mapping[key] = rr[map_rooms['room']]

    def find_room(guest_string):
        if not guest_string:
            return None
        for guest in re.split(r'[,\n;]+', guest_string):
            key = norm_name(guest)
            if key in room_mapping:
                room = room_mapping[key]
                try:
                    return int(float(room))
                except (ValueError, TypeError):
                    return room
        return None

    def get(row, key):
        col = map_data.get(key)
        return row[col] if col else None

    records = []
    for _, row in df_data.iterrows():
        guest = clean_guest_names(get(row, 'guest'))

        checkin = pd.to_datetime(get(row, 'checkin'), dayfirst=True, errors='coerce')
        checkout = pd.to_datetime(get(row, 'checkout'), dayfirst=True, errors='coerce')
        checkin = None if pd.isna(checkin) else checkin.to_pydatetime()
        checkout = None if pd.isna(checkout) else checkout.to_pydatetime()

        nights_v = get(row, 'nights')
        nights = None
        if nights_v is not None and not pd.isna(nights_v):
            try:
                nights = int(float(nights_v))
            except (ValueError, TypeError):
                nights = None
        if nights is None and checkin and checkout:
            nights = max(1, (checkout.date() - checkin.date()).days)

        category = get(row, 'category')
        category = '' if category is None or pd.isna(category) else str(category).strip()
        num_rooms = max(1, len([x for x in category.split('\n') if x.strip()])) if category else 1

        price = parse_price(get(row, 'price'))

        bron_v = get(row, 'bron')
        bron = '' if bron_v is None or pd.isna(bron_v) else str(bron_v).strip()

        source_v = get(row, 'source')
        source = '' if source_v is None or pd.isna(source_v) else str(source_v).strip()

        sm = (source_map or {}).get(source)
        if sm:
            commission, agent = float(sm[0]), sm[1]
        else:
            commission, agent = match_agent(
                source, agents or default_agents(),
                date=checkin.date() if checkin else None, hotel='tch')

        extra_vals = [row[c] if not pd.isna(row[c]) else '' for c in extra_data]

        records.append({
            'bron': bron, 'num_rooms': num_rooms, 'guest': guest, 'price': price,
            'commission': commission, 'agent': agent,
            'unknown_source': ((not agent) and bool(source)
                               and source not in (source_map or {})),
            'checkin': checkin, 'nights': nights, 'checkout': checkout,
            'source': source, 'category': category, 'room': find_room(guest),
            'extra_vals': extra_vals, 'dup_note': None,
        })

    if dup_checker:
        notes = dup_checker([{
            'bron': r['bron'], 'guest_norm': norm_name(r['guest']),
            'checkin': r['checkin'].date() if r['checkin'] else None,
            'checkout': r['checkout'].date() if r['checkout'] else None,
        } for r in records])
        for r, n in zip(records, notes):
            r['dup_note'] = n

    records.sort(key=lambda x: norm_name(x['guest']))

    errors_summary = {
        "Отсутствует № комнаты (КРАСНЫЙ)": 0,
        "Не указан Гость (ЖЕЛТЫЙ)": 0,
        "Отсутствуют Даты (ОРАНЖЕВЫЙ)": 0,
        "Отсутствует Стоимость (ГОЛУБОЙ)": 0,
        "Нераспознанный источник — комиссия 0% (ЖЕЛТЫЙ)": 0,
        "Переход кварталов (ФИОЛЕТОВЫЙ)": 0,
    }

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    headers = ['№ брони', 'Номеров', 'Гость', 'Стоимость', 'Комиссия %', 'Агент',
               'Стоимость без комиссии', 'Заезд', 'Ночей', 'Сут.ставка', 'Метод',
               'Налог', 'Выезд', 'Источник', 'Категория номера', '№ комнаты',
               'корпус', 'Примечание'] + extra_data
    n_cols = len(headers)
    font_bold = Font(bold=True)
    base = 'G' if tp.use_full else 'D'

    def write_sheet(name, recs, count_errors=False):
        ws = wb.create_sheet(name)
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.font = font_bold
        for i, rec in enumerate(recs):
            r = i + 2
            ws.cell(row=r, column=1, value=rec['bron'])
            ws.cell(row=r, column=2, value=rec['num_rooms'])

            c_guest = ws.cell(row=r, column=3, value=rec['guest'])
            c_guest.alignment = Alignment(wrap_text=True, vertical='top')
            if not rec['guest']:
                c_guest.fill = FILL_GUEST
                if count_errors: errors_summary["Не указан Гость (ЖЕЛТЫЙ)"] += 1

            c_price = ws.cell(row=r, column=4, value=rec['price'])
            c_price.number_format = '#,##0.00'
            if rec['price'] is None:
                c_price.fill = FILL_PRICE
                if count_errors: errors_summary["Отсутствует Стоимость (ГОЛУБОЙ)"] += 1

            c_comm = ws.cell(row=r, column=5, value=rec['commission'])
            c_comm.number_format = '0.##'
            c_agent = ws.cell(row=r, column=6, value=rec['agent'] or '')
            if rec['unknown_source']:
                c_comm.fill = FILL_GUEST
                c_agent.fill = FILL_GUEST
                if count_errors:
                    errors_summary["Нераспознанный источник — комиссия 0% (ЖЕЛТЫЙ)"] += 1

            c_full = ws.cell(row=r, column=7,
                             value=f'=ROUND(D{r}/(100-E{r})*100,2)')
            c_full.number_format = '#,##0.00'

            c_in = ws.cell(row=r, column=8, value=rec['checkin'])
            if rec['checkin']:
                c_in.number_format = DATE_FMT
            else:
                c_in.fill = FILL_DATE
                if count_errors: errors_summary["Отсутствуют Даты (ОРАНЖЕВЫЙ)"] += 1

            ws.cell(row=r, column=9, value=rec['nights'])
            c_rate = ws.cell(row=r, column=10, value=tp.f_daily(base, 'I', r))
            c_rate.number_format = '#,##0.00'
            ws.cell(row=r, column=11, value=tp.f_method('J', r))
            c_tax = ws.cell(row=r, column=12, value=tp.f_tax('K', 'I', base, r))
            c_tax.number_format = '#,##0.00'

            c_out = ws.cell(row=r, column=13, value=rec['checkout'])
            if rec['checkout']:
                c_out.number_format = DATE_FMT
            else:
                c_out.fill = FILL_DATE
                if count_errors: errors_summary["Отсутствуют Даты (ОРАНЖЕВЫЙ)"] += 1

            ws.cell(row=r, column=14, value=rec['source'])
            cat = ws.cell(row=r, column=15, value=rec['category'])
            cat.alignment = Alignment(wrap_text=True, vertical='top')

            c_room = ws.cell(row=r, column=16, value=rec['room'])
            if rec['room'] is None:
                c_room.fill = FILL_ROOM
                if count_errors: errors_summary["Отсутствует № комнаты (КРАСНЫЙ)"] += 1
            else:
                ws.cell(row=r, column=17, value=f'=IF(P{r}<50,1,2)')

            c_note = ws.cell(row=r, column=18, value=rec['dup_note'] or '')
            if rec['dup_note']:
                c_note.fill = FILL_DUP
                if count_errors: errors_summary["Переход кварталов (ФИОЛЕТОВЫЙ)"] += 1

            for k, v in enumerate(rec['extra_vals']):
                ws.cell(row=r, column=19 + k, value=v)

        if recs:
            last = len(recs) + 1
            tr = last + 1
            ws.cell(row=tr, column=3, value="ИТОГО:")
            ws.cell(row=tr, column=2, value=f'=SUM(B2:B{last})')
            ws.cell(row=tr, column=4, value=f'=SUM(D2:D{last})')
            ws.cell(row=tr, column=7, value=f'=SUM(G2:G{last})')
            ws.cell(row=tr, column=9, value=f'=SUM(I2:I{last})')
            ws.cell(row=tr, column=12, value=f'=SUM(L2:L{last})')
            for c in range(1, n_cols + 1):
                cell = ws.cell(row=tr, column=c)
                cell.font = font_bold
                cell.fill = FILL_TOTALS

        widths = {'A': 24, 'B': 9, 'C': 35, 'D': 12, 'E': 10, 'F': 32, 'G': 20,
                  'H': 17, 'I': 8, 'J': 11, 'K': 12, 'L': 11, 'M': 17, 'N': 25,
                  'O': 30, 'P': 10, 'Q': 8, 'R': 45}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
        return ws

    def corpus_of(rec):
        try:
            return 1 if float(rec['room']) < 50 else 2
        except (ValueError, TypeError):
            return None

    write_sheet('1&2', records, count_errors=True)
    write_sheet('нет корпуса', [x for x in records if x['room'] is None])
    write_sheet('1', [x for x in records if corpus_of(x) == 1])
    write_sheet('2', [x for x in records if corpus_of(x) == 2])

    wb.save(output_path)
    log.info("Чайковский сохранён: %s", output_path)

    bookings = [{
        'guest': r['guest'], 'guest_norm': norm_name(r['guest']), 'bron': r['bron'],
        'checkin': r['checkin'].date().isoformat() if r['checkin'] else None,
        'checkout': r['checkout'].date().isoformat() if r['checkout'] else None,
        'price': r['price'],
        'price_full': round(r['price'] / (100 - r['commission']) * 100, 2)
                      if r['price'] is not None else None,
        'nights': r['nights'], 'dup_note': r['dup_note'],
    } for r in records]
    return errors_summary, bookings


def process_dona(data_path, output_path, map_data=None,
                 extra_data=None, dup_checker=None, tax=None,
                 agents=None, source_map=None):
    """Возвращает (summary, bookings)."""
    tp = TaxParams(tax)
    log.info("Дона: %s", data_path)
    df_data = read_table(data_path)
    headers = [str(c) for c in df_data.columns]
    map_data = map_data or guess_mapping(headers, DNA_DATA_SCHEMA)
    miss = [f['label'] for f in DNA_DATA_SCHEMA if f['required'] and not map_data.get(f['key'])]
    if miss:
        raise ValueError(f"Не сопоставлены колонки: {', '.join(miss)}")
    extra_data = [c for c in (extra_data or []) if c in df_data.columns]

    def get(row, key):
        col = map_data.get(key)
        return row[col] if col else None

    records = []
    for _, row in df_data.iterrows():
        guest = clean_guest_names(get(row, 'guest'))

        checkin = pd.to_datetime(get(row, 'checkin'), dayfirst=True, errors='coerce')
        checkout = pd.to_datetime(get(row, 'checkout'), dayfirst=True, errors='coerce')
        checkin = None if pd.isna(checkin) else checkin.to_pydatetime()
        checkout = None if pd.isna(checkout) else checkout.to_pydatetime()

        nights_v = get(row, 'nights')
        nights = None
        if nights_v is not None and not pd.isna(nights_v):
            try:
                nights = int(float(nights_v))
            except (ValueError, TypeError):
                nights = None
        if nights is None and checkin and checkout:
            nights = max(1, (checkout.date() - checkin.date()).days)

        category = get(row, 'category')
        category = '' if category is None or pd.isna(category) else str(category).strip()
        num_rooms = max(1, len([x for x in category.split('\n') if x.strip()])) if category else 1

        price = parse_price(get(row, 'price'))

        bron_v = get(row, 'bron')
        bron = '' if bron_v is None or pd.isna(bron_v) else str(bron_v).strip()

        source_v = get(row, 'source')
        source = '' if source_v is None or pd.isna(source_v) else str(source_v).strip()

        sm = (source_map or {}).get(source)
        if sm:
            commission, agent = float(sm[0]), sm[1]
        else:
            # комиссия берётся на дату заезда (плавающие проценты)
            commission, agent = match_agent(
                source, agents or default_agents(),
                date=checkin.date() if checkin else None, hotel='dna')
        extra_vals = [row[c] if not pd.isna(row[c]) else '' for c in extra_data]

        records.append({
            'bron': bron, 'num_rooms': num_rooms, 'guest': guest, 'price': price,
            'commission': commission, 'agent': agent,
            'unknown_source': ((not agent) and bool(source)
                               and source not in (source_map or {})),
            'checkin': checkin, 'nights': nights, 'checkout': checkout,
            'source': source, 'category': category,
            'extra_vals': extra_vals, 'dup_note': None,
        })

    if dup_checker:
        notes = dup_checker([{
            'bron': r['bron'], 'guest_norm': norm_name(r['guest']),
            'checkin': r['checkin'].date() if r['checkin'] else None,
            'checkout': r['checkout'].date() if r['checkout'] else None,
        } for r in records])
        for r, n in zip(records, notes):
            r['dup_note'] = n

    records.sort(key=lambda x: norm_name(x['guest']))

    summary = {
        "Не указан Гость (ЖЕЛТЫЙ)": 0,
        "Отсутствуют Даты (ОРАНЖЕВЫЙ)": 0,
        "Отсутствует Стоимость (ГОЛУБОЙ)": 0,
        "Нераспознанный источник — комиссия 0% (ЖЕЛТЫЙ)": 0,
        "Переход кварталов (ФИОЛЕТОВЫЙ)": 0,
    }

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    headers_out = ['№ брони', 'Номеров', 'Гость', 'Стоимость', 'Комиссия %', 'Агент',
                   'Стоимость без комиссии', 'Заезд', 'Ночей', 'Сут.ставка', 'Метод',
                   'Налог', 'Выезд', 'Источник', 'Категория номера', 'Примечание'] + extra_data
    n_cols = len(headers_out)
    font_bold = Font(bold=True)

    ws = wb.create_sheet('Дона')
    for c, h in enumerate(headers_out, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = font_bold

    for i, rec in enumerate(records):
        r = i + 2
        ws.cell(row=r, column=1, value=rec['bron'])
        ws.cell(row=r, column=2, value=rec['num_rooms'])

        c_guest = ws.cell(row=r, column=3, value=rec['guest'])
        c_guest.alignment = Alignment(wrap_text=True, vertical='top')
        if not rec['guest']:
            c_guest.fill = FILL_GUEST
            summary["Не указан Гость (ЖЕЛТЫЙ)"] += 1

        c_price = ws.cell(row=r, column=4, value=rec['price'])
        c_price.number_format = '#,##0.00'
        if rec['price'] is None:
            c_price.fill = FILL_PRICE
            summary["Отсутствует Стоимость (ГОЛУБОЙ)"] += 1

        c_comm = ws.cell(row=r, column=5, value=rec['commission'])
        c_comm.number_format = '0'
        c_agent = ws.cell(row=r, column=6, value=rec['agent'] or '')
        if rec['unknown_source']:
            c_comm.fill = FILL_GUEST
            c_agent.fill = FILL_GUEST
            summary["Нераспознанный источник — комиссия 0% (ЖЕЛТЫЙ)"] += 1

        c_full = ws.cell(row=r, column=7, value=f'=ROUND(D{r}/(100-E{r})*100,2)')
        c_full.number_format = '#,##0.00'

        c_in = ws.cell(row=r, column=8, value=rec['checkin'])
        if rec['checkin']:
            c_in.number_format = DATE_FMT
        else:
            c_in.fill = FILL_DATE
            summary["Отсутствуют Даты (ОРАНЖЕВЫЙ)"] += 1

        ws.cell(row=r, column=9, value=rec['nights'])
        # налог от стоимости без комиссии (G) или как в документе (D)
        pcol = 'G' if tp.use_full else 'D'
        c_rate = ws.cell(row=r, column=10, value=tp.f_daily(pcol, 'I', r))
        c_rate.number_format = '#,##0.00'
        ws.cell(row=r, column=11, value=tp.f_method('J', r))
        c_tax = ws.cell(row=r, column=12, value=tp.f_tax('K', 'I', pcol, r))
        c_tax.number_format = '#,##0.00'

        c_out = ws.cell(row=r, column=13, value=rec['checkout'])
        if rec['checkout']:
            c_out.number_format = DATE_FMT
        else:
            c_out.fill = FILL_DATE
            summary["Отсутствуют Даты (ОРАНЖЕВЫЙ)"] += 1

        ws.cell(row=r, column=14, value=rec['source'])
        cat = ws.cell(row=r, column=15, value=rec['category'])
        cat.alignment = Alignment(wrap_text=True, vertical='top')

        c_note = ws.cell(row=r, column=16, value=rec['dup_note'] or '')
        if rec['dup_note']:
            c_note.fill = FILL_DUP
            summary["Переход кварталов (ФИОЛЕТОВЫЙ)"] += 1

        for k, v in enumerate(rec['extra_vals']):
            ws.cell(row=r, column=17 + k, value=v)

    if records:
        last = len(records) + 1
        tr = last + 1
        ws.cell(row=tr, column=3, value="ИТОГО:")
        ws.cell(row=tr, column=2, value=f'=SUM(B2:B{last})')
        ws.cell(row=tr, column=4, value=f'=SUM(D2:D{last})')
        ws.cell(row=tr, column=7, value=f'=SUM(G2:G{last})')
        ws.cell(row=tr, column=9, value=f'=SUM(I2:I{last})')
        ws.cell(row=tr, column=12, value=f'=SUM(L2:L{last})')
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=tr, column=c)
            cell.font = font_bold
            cell.fill = FILL_TOTALS

    widths = {'A': 24, 'B': 9, 'C': 35, 'D': 12, 'E': 11, 'F': 30, 'G': 20,
              'H': 17, 'I': 8, 'J': 11, 'K': 12, 'L': 11, 'M': 17, 'N': 35, 'O': 30, 'P': 45}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    wb.save(output_path)
    log.info("Дона сохранена: %s", output_path)

    bookings = [{
        'guest': r['guest'], 'guest_norm': norm_name(r['guest']), 'bron': r['bron'],
        'checkin': r['checkin'].date().isoformat() if r['checkin'] else None,
        'checkout': r['checkout'].date().isoformat() if r['checkout'] else None,
        'price': r['price'],
        'price_full': round(r['price'] / (100 - r['commission']) * 100, 2)
                      if r['price'] is not None else None,
        'nights': r['nights'], 'dup_note': r['dup_note'],
    } for r in records]
    return summary, bookings


def load_workbook_safe(path):
    try:
        return openpyxl.load_workbook(path, data_only=True)
    except Exception:
        with zipfile.ZipFile(path) as z:
            items = {n: z.read(n) for n in z.namelist()}
        styles_name = 'xl/styles.xml'
        if styles_name in items:
            styles = items[styles_name].decode('utf-8')

            def fix_rgb(m):
                v = m.group(1)
                return f'rgb="{v.rjust(8, "0")}"' if len(v) < 8 else m.group(0)

            styles = re.sub(r'rgb="([0-9A-Fa-f]{1,8})"', fix_rgb, styles)
            items[styles_name] = styles.encode('utf-8')

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as z:
            for name, data in items.items():
                z.writestr(name, data)
        buf.seek(0)
        return openpyxl.load_workbook(buf, data_only=True)


def process_usadba( report_path, output_path, dup_checker=None,
                   tax=None, agents=None, source_map=None):
    """Возвращает (stats, bookings)."""
    tp = TaxParams(tax)
    log.info("Усадьба: %s", report_path)
    wb_in = load_workbook_safe(report_path)
    ws_in = wb_in[wb_in.sheetnames[0]]

    valid_states = ('Проживал', 'Сейчас проживает')

    def parse_dt(v):
        if isinstance(v, datetime):
            return v
        if v is None:
            return None
        s = str(v).strip()
        for fmt in ('%d.%m.%Y %H:%M', '%d.%m.%Y %H:%M:%S', '%d.%m.%Y'):
            try:
                return datetime.strptime(s, fmt)
            except ValueError:
                continue
        return None

    records = []
    rooms = set()
    current_room = None
    stats = {'rows': 0, 'rooms': 0, 'zero_price': 0, 'bad_dates': 0,
             'still_living': 0, 'unknown_source': 0}

    for r in range(1, ws_in.max_row + 1):
        a_val = ws_in.cell(row=r, column=1).value
        if isinstance(a_val, str) and a_val.strip().startswith('Номер:'):
            current_room = a_val.split(':', 1)[1].strip()
            rooms.add(current_room)
            continue

        state = ws_in.cell(row=r, column=5).value
        fio = ws_in.cell(row=r, column=6).value
        if not fio or not state or str(state).strip() not in valid_states:
            continue

        checkin_raw = ws_in.cell(row=r, column=7).value
        checkout_raw = ws_in.cell(row=r, column=8).value
        price = parse_price(ws_in.cell(row=r, column=9).value)
        tariff = ws_in.cell(row=r, column=14).value
        org = ws_in.cell(row=r, column=16).value
        org_s = '' if org is None else str(org).strip()

        dt_in = parse_dt(checkin_raw)
        dt_out = parse_dt(checkout_raw)

        if dt_in and dt_out:
            nights = max(1, (dt_out.date() - dt_in.date()).days)
        else:
            nights = 1
            stats['bad_dates'] += 1

        if str(state).strip() == 'Сейчас проживает':
            stats['still_living'] += 1
        if not price:
            stats['zero_price'] += 1

        sm = (source_map or {}).get(org_s)
        if sm:
            commission, agent = float(sm[0]), sm[1]
        else:
            commission, agent = match_agent(
                org_s, agents or default_agents(),
                date=dt_in.date() if dt_in else None, hotel='usd')
        if not agent and org_s:
            stats['unknown_source'] += 1

        records.append({
            'fio': str(fio).strip(),
            'checkin': dt_in if dt_in else (str(checkin_raw or '').strip()),
            'checkout': dt_out if dt_out else (str(checkout_raw or '').strip()),
            'nights': nights,
            'price': price if price is not None else 0,
            'commission': commission, 'agent': agent,
            'unknown_source': ((not agent) and bool(org_s)
                               and org_s not in (source_map or {})),
            'tariff': str(tariff).strip() if tariff else '',
            'org': org_s,
            'room': current_room,
            'dup_note': None,
        })

    if not records:
        raise ValueError("В файле не найдено ни одной строки гостя.\n"
                         "Убедитесь, что выбран отчёт «Доход по номеру» из UCS Shelter.")

    if dup_checker:
        notes = dup_checker([{
            'bron': '', 'guest_norm': norm_name(rec['fio']),
            'checkin': rec['checkin'].date() if isinstance(rec['checkin'], datetime) else None,
            'checkout': rec['checkout'].date() if isinstance(rec['checkout'], datetime) else None,
        } for rec in records])
        for rec, n in zip(records, notes):
            rec['dup_note'] = n
    stats['dups'] = sum(1 for rec in records if rec['dup_note'])

    stats['rows'] = len(records)
    stats['rooms'] = len(rooms)

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = 'Page 1'

    headers = ['Фамилия Имя Отчество', 'Заселение', 'ночей', 'Выселение ',
               'Стоимость', 'Комиссия %', 'Агент', 'Стоимость без комиссии',
               'Сут.ставка', 'Метод', 'Налог', 'Тариф',
               'Организация\nзаказчик', 'Примечание']
    font_bold = Font(bold=True)
    for c, h in enumerate(headers, start=1):
        cell = ws_out.cell(row=1, column=c, value=h)
        cell.font = font_bold
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    base = 'H' if tp.use_full else 'E'

    for i, rec in enumerate(records):
        row = i + 2
        ws_out.cell(row=row, column=1, value=rec['fio'])

        c_in = ws_out.cell(row=row, column=2, value=rec['checkin'])
        c_out = ws_out.cell(row=row, column=4, value=rec['checkout'])
        if isinstance(rec['checkin'], datetime):
            c_in.number_format = DATE_FMT
        else:
            c_in.fill = FILL_DATE
        if isinstance(rec['checkout'], datetime):
            c_out.number_format = DATE_FMT
        else:
            c_out.fill = FILL_DATE

        ws_out.cell(row=row, column=3, value=rec['nights'])

        c_price = ws_out.cell(row=row, column=5, value=rec['price'])
        c_price.number_format = r'#,##0.00\ "₽"'
        if not rec['price']:
            c_price.fill = FILL_PRICE

        c_comm = ws_out.cell(row=row, column=6, value=rec['commission'])
        c_comm.number_format = '0.##'
        c_agent = ws_out.cell(row=row, column=7, value=rec['agent'] or '')
        if rec['unknown_source']:
            c_comm.fill = FILL_GUEST
            c_agent.fill = FILL_GUEST

        c_full = ws_out.cell(row=row, column=8,
                             value=f'=ROUND(E{row}/(100-F{row})*100,2)')
        c_full.number_format = '#,##0.00'

        c_rate = ws_out.cell(row=row, column=9, value=tp.f_daily(base, 'C', row))
        c_rate.number_format = '#,##0.00'
        ws_out.cell(row=row, column=10, value=tp.f_method('I', row))
        c_tax = ws_out.cell(row=row, column=11, value=tp.f_tax('J', 'C', base, row))
        c_tax.number_format = '#,##0.00'
        ws_out.cell(row=row, column=12, value=rec['tariff'])
        ws_out.cell(row=row, column=13, value=rec['org'])

        c_note = ws_out.cell(row=row, column=14, value=rec['dup_note'] or '')
        if rec['dup_note']:
            c_note.fill = FILL_DUP

    last_row = len(records) + 1
    total_row = last_row + 1

    ws_out.cell(row=total_row, column=1, value="ИТОГО:")
    ws_out.cell(row=total_row, column=5, value=f'=SUM(E2:E{last_row})').number_format = r'#,##0.00\ "₽"'
    ws_out.cell(row=total_row, column=8, value=f'=SUM(H2:H{last_row})')
    ws_out.cell(row=total_row, column=11, value=f'=SUM(K2:K{last_row})')
    for c in range(1, 15):
        cell = ws_out.cell(row=total_row, column=c)
        cell.font = font_bold
        cell.fill = FILL_TOTALS

    widths = {'A': 38, 'B': 18, 'C': 8, 'D': 18, 'E': 14, 'F': 10, 'G': 30,
              'H': 18, 'I': 12, 'J': 12, 'K': 12, 'L': 10, 'M': 40, 'N': 45}
    for col, w in widths.items():
        ws_out.column_dimensions[col].width = w

    ws2 = wb_out.create_sheet('Лист1')
    ws2.cell(row=1, column=1, value='Заселение').font = font_bold
    ws2.cell(row=1, column=2, value='Выселение ').font = font_bold
    for i, rec in enumerate(records):
        row = i + 2
        c_a = ws2.cell(row=row, column=1, value=rec['checkin'])
        c_b = ws2.cell(row=row, column=2, value=rec['checkout'])
        if isinstance(rec['checkin'], datetime):
            c_a.number_format = DATE_FMT
        if isinstance(rec['checkout'], datetime):
            c_b.number_format = DATE_FMT
        ws2.cell(row=row, column=4, value=f'=INT(A{row})')
        ws2.cell(row=row, column=5, value=f'=INT(B{row})')
        ws2.cell(row=row, column=6, value=f'=E{row}-D{row}')
    ws2.column_dimensions['A'].width = 18
    ws2.column_dimensions['B'].width = 18

    wb_out.save(output_path)
    log.info("Усадьба сохранена: %s", output_path)

    bookings = [{
        'guest': rec['fio'], 'guest_norm': norm_name(rec['fio']), 'bron': '',
        'checkin': rec['checkin'].date().isoformat() if isinstance(rec['checkin'], datetime) else None,
        'checkout': rec['checkout'].date().isoformat() if isinstance(rec['checkout'], datetime) else None,
        'price': rec['price'],
        'price_full': round(rec['price'] / (100 - rec['commission']) * 100, 2)
                      if rec['price'] else None,
        'nights': rec['nights'], 'dup_note': rec['dup_note'],
    } for rec in records]
    return stats, bookings


def usadba_sources(path):
    """Колонка «Организация-заказчик» из отчёта UCS Shelter (для диалога источников)."""
    wb = load_workbook_safe(path)
    ws = wb[wb.sheetnames[0]]
    out = []
    for r in range(1, ws.max_row + 1):
        state = ws.cell(row=r, column=5).value
        org = ws.cell(row=r, column=16).value
        if org and state and str(state).strip() in ('Проживал', 'Сейчас проживает'):
            out.append(str(org))
    return out
