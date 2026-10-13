# -*- coding: utf-8 -*-
"""Импорт TravelLine: разбор броней в строки Excel (фикстуры, без сети)."""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core import complete_mapping, guess_mapping, mapping_valid, read_table
from schemas import DNA_DATA_SCHEMA, TCH_DATA_SCHEMA, TCH_ROOMS_SCHEMA
from tl_client import (TL_DATA_MAP, TL_EXTRA, TL_ROOMS_MAP,
                       ROOM_LABEL_TEXT_PREFIX, accommodation_room_map,
                       bookings_to_rows, channel_label,
                       is_tl_export, write_api_inspection_workbook,
                       write_data_workbook, write_rooms_workbook)


def _service(value):
    if isinstance(value, dict):
        service = dict(value)
        amount = service.pop('amount', None)
        if amount is not None:
            service.setdefault('total', {'priceAfterTax': amount})
        return service
    return {'total': {'priceAfterTax': value}}


def _stay(arr='2026-07-10T14:00', dep='2026-07-12T12:00', room=10000.0,
          services=(), guests=(('Иван', 'Петров'),), rt='Стандарт',
          rp='Базовый'):
    return {
        'stayDates': {'arrivalDateTime': arr, 'departureDateTime': dep},
        'total': {'priceAfterTax': room},
        'services': [_service(s) for s in services],
        'extraStayCharges': {'earlyArrival': None, 'lateDeparture': None},
        'guests': [{'firstName': f, 'lastName': l} for f, l in guests],
        'roomType': {'name': rt}, 'ratePlans': [{'name': rp}],
    }


def _book(number='20260701-8186-1', status='Active', stays=None,
          book_services=(), total=10000.0, source=None):
    return {'number': number, 'status': status,
            'roomStays': stays if stays is not None else [_stay()],
            'services': [_service(s) for s in book_services],
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


def test_service_names_are_exported_with_amounts():
    stay = _stay(services=({'name': 'Завтрак', 'amount': 1500.0},))
    stay['extraStayCharges'] = {
        'earlyArrival': {'total': {'priceAfterTax': 100.0}},
        'lateDeparture': None,
    }
    b = _book(stays=[stay],
              book_services=({'name': 'Парковка', 'amount': 900.0},),
              total=12500.0)
    rows, _ = bookings_to_rows([b], QS, QE)
    row = rows[0]
    assert row['Допуслуги'] == 2500.0
    names = row['Наименование допуслуг']
    assert 'Завтрак — 1 500,00 ₽' in names
    assert 'Парковка — 900,00 ₽' in names
    assert 'Ранний заезд — 100,00 ₽' in names
    assert row['Сумма продления номера'] == 0


def test_room_extension_amount_is_separated_from_other_extras():
    stay = _stay(services=(
        {'name': 'Продление номера', 'amount': 1200.0},
        {'name': 'Завтрак', 'amount': 800.0},
    ))
    b = _book(stays=[stay], total=12000.0)
    rows, _ = bookings_to_rows([b], QS, QE)
    assert rows[0]['Допуслуги'] == 2000.0
    assert rows[0]['Наименование допуслуг'] == (
        'Продление номера — 1 200,00 ₽; Завтрак — 800,00 ₽')
    assert rows[0]['Сумма продления номера'] == 1200.0


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


def test_accommodation_inventory_maps_room_display_names_to_buildings():
    inventory = {'accommodationInventory': {'buildings': [
        {'id': 'building-1', 'name': 'Корпус № 1', 'floors': [
            {'id': 'floor-1', 'name': '1 этаж', 'rooms': [
                {'id': 'room-111', 'displayName': '111',
                 'roomTypeId': 'type-1', 'floorId': 'floor-1'},
            ]},
        ]},
        {'id': 'building-2', 'name': 'Main building', 'floors': [
            {'id': 'floor-2', 'name': '2 этаж', 'rooms': [
                {'id': 'room-205', 'displayName': '205',
                 'roomTypeId': 'type-2', 'floorId': 'floor-2'},
                {'id': 'room-001', 'displayName': '001',
                 'roomTypeId': 'type-2', 'floorId': 'floor-2'},
                {'roomId': {'id': ' room-A12 '}, 'name': 'A-12'},
            ]},
        ]},
    ]}}

    mapped = accommodation_room_map(inventory)
    assert mapped['room-111'] == {
        'room': 111, 'corpus': 1, 'building_name': 'Корпус № 1'}
    # Arbitrary building names have no number: use their stable API position.
    assert mapped['room-205'] == {
        'room': 205, 'corpus': 2, 'building_name': 'Main building'}
    assert mapped['room-001']['room'] == '001'
    assert mapped['room-A12'] == {
        'room': 'A-12', 'corpus': 2, 'building_name': 'Main building'}
    assert 'room-111' not in accommodation_room_map({'buildings': []})


def test_bookings_rows_keep_real_room_assignment_and_leave_missing_blank():
    assigned = _stay(guests=(('Иван', 'Петров'),))
    assigned.update({'roomId': 'room-72', 'pmsRoomStayId': 'stay-assigned'})
    unassigned = _stay(guests=(('Иван', 'Петров'),))
    unassigned.update({'roomId': None, 'pmsRoomStayId': 'stay-unassigned'})
    unknown = _stay(guests=(('Петр', 'Сидоров'),))
    unknown.update({'roomId': 'not-in-inventory', 'pmsRoomStayId': 'stay-unknown'})
    details = [_book(stays=[assigned, unassigned, unknown], total=30000.0)]
    inventory = {'accommodationInventory': {'buildings': [{
        'name': 'Корпус 1', 'floors': [{'rooms': [
            {'id': ' room-72 ', 'displayName': '', 'roomNumber': '72'},
        ]}],
    }]}}

    rows, report = bookings_to_rows(details, QS, QE, inventory)

    assert [row['№ комнаты'] for row in rows] == [72, None, None]
    assert [row['Корпус'] for row in rows] == [1, None, None]
    assert report['room_assignments']['inventory_rooms'] == 1
    assert report['room_assignments']['assigned'] == 1
    assert report['room_assignments']['unassigned'] == 2
    assert report['room_assignments']['room_id_present'] == 2
    assert report['room_assignments']['room_id_missing'] == 1
    assert report['room_assignments']['room_id_not_in_inventory'] == 1


def test_room_diagnostic_workbook_shows_api_ids_and_match_reasons(tmp_path):
    import openpyxl

    assigned = _stay()
    assigned.update({'roomId': 'room-72', 'pmsRoomStayId': 'stay-assigned'})
    missing = _stay()
    missing.update({'roomId': None, 'pmsRoomStayId': 'stay-missing'})
    unknown = _stay()
    unknown.update({'roomId': 'room-new', 'pmsRoomStayId': 'stay-unknown'})
    details = [_book(stays=[assigned, missing, unknown], total=30000.0)]
    inventory = {'accommodationInventory': {'buildings': [{
        'id': 'building-1', 'name': 'Корпус 1', 'floors': [{
            'id': 'floor-1', 'name': '1 этаж', 'rooms': [
                {'id': 'room-72', 'displayName': '72'},
            ],
        }],
    }]}}
    rows, report = bookings_to_rows(details, QS, QE, inventory)
    stats = report['room_assignments']
    assert stats['room_id_present'] == 2
    assert stats['room_id_missing'] == 1
    assert stats['room_id_not_in_inventory'] == 1
    assert stats['unknown_room_id_samples'] == ['room-new']

    path = str(tmp_path / 'inspection.xlsx')
    write_api_inspection_workbook(
        rows, report, path, details=details,
        accommodation_inventory=inventory)
    wb = openpyxl.load_workbook(path, data_only=True)
    stays = wb['Проживания API']
    headers = [cell.value for cell in stays[1]]
    status_col = headers.index('Результат сопоставления') + 1
    room_id_col = headers.index('roomId') + 1
    values = list(stays.iter_rows(min_row=2, values_only=True))
    assert [row[status_col - 1] for row in values] == [
        'Сопоставлено', 'roomId/room отсутствует', 'ID не найден в инвентаре']
    assert values[2][room_id_col - 1] == 'room-new'
    assert wb['Номерной фонд API']['F2'].value == 'room-72'
    assert wb['Данные']['N2'].value == 72


def test_tl_room_workbooks_drive_corpus_and_do_not_guess_unassigned_stay(tmp_path):
    import openpyxl
    import core

    first = _stay(guests=(('Иван', 'Петров'),))
    first.update({'roomId': 'room-72', 'pmsRoomStayId': 'stay-1'})
    second = _stay(guests=(('Ольга', 'Сидорова'),))
    second.update({'roomId': 'room-15', 'pmsRoomStayId': 'stay-2'})
    no_room = _stay(guests=(('Иван', 'Петров'),))
    no_room.update({'roomId': None, 'pmsRoomStayId': 'stay-3'})
    details = [
        _book(number='assigned-1', stays=[first]),
        _book(number='assigned-2', stays=[second]),
        _book(number='unassigned', stays=[no_room]),
    ]
    inventory = {'accommodationInventory': {'buildings': [
        {'name': 'Корпус 1', 'floors': [{'rooms': [
            {'id': 'room-72', 'displayName': '72'},
        ]}]},
        {'name': 'Корпус 2', 'floors': [{'rooms': [
            {'id': 'room-15', 'displayName': '15'},
        ]}]},
    ]}}
    rows, _ = bookings_to_rows(details, QS, QE, inventory)
    data_path = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rooms_path = str(tmp_path / 'TL_tch_2026III_rooms.xlsx')
    output_path = str(tmp_path / 'report.xlsx')
    write_data_workbook(rows, data_path)
    write_rooms_workbook(rows, rooms_path)

    core.process_tchaikovsky(
        data_path, rooms_path, output_path, TL_DATA_MAP, TL_ROOMS_MAP,
        extra_data=TL_EXTRA, source_map={'TravelLine YAX': (0, '')})
    wb = openpyxl.load_workbook(output_path, data_only=False)
    corpus_1 = wb['1']
    corpus_2 = wb['2']
    no_corpus = wb['нет корпуса']
    assert corpus_1['A2'].value == 'assigned-1'
    assert corpus_1['S2'].value == 72 and corpus_1['T2'].value == 1
    assert corpus_2['A2'].value == 'assigned-2'
    assert corpus_2['S2'].value == 15 and corpus_2['T2'].value == 2
    assert no_corpus['A2'].value == 'unassigned'
    assert no_corpus['S2'].value is None and no_corpus['T2'].value is None


def test_manual_rooms_fill_blank_direct_room_from_exported_data(tmp_path):
    import openpyxl
    import core

    stay = _stay(guests=(('Иван', 'Петров'),))
    stay.update({'roomId': None, 'pmsRoomStayId': 'stay-unassigned'})
    rows, _ = bookings_to_rows([_book(stays=[stay])], QS, QE)
    data_path = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rooms_path = str(tmp_path / 'manual_rooms.xlsx')
    output_path = str(tmp_path / 'report.xlsx')
    write_data_workbook(rows, data_path)

    manual_rooms = openpyxl.Workbook()
    sheet = manual_rooms.active
    sheet.title = 'Номера'
    sheet.append(['ФИО', '№ комнаты', 'Корпус'])
    sheet.append(['Иван Петров', 72, 1])
    manual_rooms.save(rooms_path)

    summary, _ = core.process_tchaikovsky(
        data_path, rooms_path, output_path, TL_DATA_MAP,
        {'fio': 'ФИО', 'room': '№ комнаты', 'corpus': 'Корпус'},
        extra_data=TL_EXTRA, source_map={'TravelLine YAX': (0, '')})

    report = openpyxl.load_workbook(output_path, data_only=False)['1']
    assert report['S2'].value == 72 and report['T2'].value == 1
    assert summary['Отсутствует № комнаты (КРАСНЫЙ)'] == 0


def test_manual_rooms_prefer_booking_number_when_guest_repeats(tmp_path):
    import openpyxl
    import core

    stay = _stay(guests=(('Иван', 'Петров'),))
    stay.update({'roomId': None, 'pmsRoomStayId': 'stay-unassigned'})
    data = _book(number='book-1', stays=[stay])
    rows, _ = bookings_to_rows([data], QS, QE)
    data_path = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rooms_path = str(tmp_path / 'manual_rooms.xlsx')
    output_path = str(tmp_path / 'report.xlsx')
    write_data_workbook(rows, data_path)

    manual_rooms = openpyxl.Workbook()
    sheet = manual_rooms.active
    sheet.append(['ФИО', '№ комнаты', 'Корпус', '№ брони'])
    sheet.append(['Иван Петров', 55, 2, 'other-book'])
    sheet.append(['Иван Петров', 72, 1, 'book-1'])
    manual_rooms.save(rooms_path)

    core.process_tchaikovsky(
        data_path, rooms_path, output_path, TL_DATA_MAP,
        {'fio': 'ФИО', 'room': '№ комнаты', 'corpus': 'Корпус',
         'bron': '№ брони'},
        extra_data=TL_EXTRA, source_map={'TravelLine YAX': (0, '')})
    report = openpyxl.load_workbook(output_path, data_only=False)['1']
    assert report['S2'].value == 72 and report['T2'].value == 1


def test_manual_rooms_can_join_by_booking_without_guest_column(tmp_path):
    import openpyxl
    import core

    stay = _stay(guests=(('Иван', 'Петров'),))
    stay.update({'roomId': None, 'pmsRoomStayId': 'stay-unassigned'})
    rows, _ = bookings_to_rows(
        [_book(number='book-91', stays=[stay])], QS, QE)
    data_path = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rooms_path = str(tmp_path / 'manual_rooms.xlsx')
    output_path = str(tmp_path / 'report.xlsx')
    write_data_workbook(rows, data_path)

    manual_rooms = openpyxl.Workbook()
    sheet = manual_rooms.active
    sheet.append(['№ брони', '№ комнаты', 'Корпус'])
    sheet.append(['book-91', 91, 2])
    manual_rooms.save(rooms_path)

    core.process_tchaikovsky(
        data_path, rooms_path, output_path, TL_DATA_MAP,
        {'fio': None, 'room': '№ комнаты', 'corpus': 'Корпус',
         'bron': '№ брони'},
        extra_data=TL_EXTRA, source_map={'TravelLine YAX': (0, '')})
    report = openpyxl.load_workbook(output_path, data_only=False)['2']
    assert report['S2'].value == 91 and report['T2'].value == 2


def test_manual_multi_room_booking_uses_unique_guest_after_booking_is_ambiguous(tmp_path):
    import openpyxl
    import core

    first = _stay(guests=(('Иван', 'Петров'),))
    first.update({'roomId': None, 'pmsRoomStayId': 'stay-1'})
    second = _stay(guests=(('Ольга', 'Сидорова'),))
    second.update({'roomId': None, 'pmsRoomStayId': 'stay-2'})
    rows, _ = bookings_to_rows(
        [_book(number='group-booking', stays=[first, second])], QS, QE)
    data_path = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rooms_path = str(tmp_path / 'manual_rooms.xlsx')
    output_path = str(tmp_path / 'report.xlsx')
    write_data_workbook(rows, data_path)

    manual_rooms = openpyxl.Workbook()
    sheet = manual_rooms.active
    sheet.append(['ФИО', '№ комнаты', 'Корпус', '№ брони'])
    sheet.append(['Иван Петров', 31, 1, 'group-booking'])
    sheet.append(['Ольга Сидорова', 82, 2, 'group-booking'])
    manual_rooms.save(rooms_path)

    core.process_tchaikovsky(
        data_path, rooms_path, output_path, TL_DATA_MAP,
        {'fio': 'ФИО', 'room': '№ комнаты', 'corpus': 'Корпус',
         'bron': '№ брони'},
        extra_data=TL_EXTRA, source_map={'TravelLine YAX': (0, '')})
    wb = openpyxl.load_workbook(output_path, data_only=False)
    assert wb['1']['S2'].value == 31
    assert wb['2']['S2'].value == 82


def test_arrivals_report_columns_auto_map_booking_and_stay_dates():
    mapping = guess_mapping(
        ['№ брони', 'ФИО гостей', 'Заезд', 'Выезд', '№ комнаты'],
        TCH_ROOMS_SCHEMA)
    assert mapping['bron'] == '№ брони'
    assert mapping['fio'] == 'ФИО гостей'
    assert mapping['room'] == '№ комнаты'
    assert mapping['checkin'] == 'Заезд'
    assert mapping['checkout'] == 'Выезд'


def test_legacy_room_mapping_gains_date_columns_without_overriding_choices():
    headers = ['№ брони', 'ФИО гостей', 'Заезд', 'Выезд', '№ комнаты']
    legacy = {'bron': '№ брони', 'fio': 'ФИО гостей',
              'room': '№ комнаты', 'corpus': None}
    updated = complete_mapping(legacy, headers, TCH_ROOMS_SCHEMA)
    assert updated['checkin'] == 'Заезд'
    assert updated['checkout'] == 'Выезд'

    legacy['checkin'] = None
    updated = complete_mapping(legacy, headers, TCH_ROOMS_SCHEMA)
    assert updated['checkin'] is None
    assert updated['checkout'] == 'Выезд'


def test_manual_multi_room_booking_uses_exact_stay_dates(tmp_path):
    import openpyxl
    import core

    first = _stay(arr='2026-07-10T14:00', dep='2026-07-12T12:00',
                  guests=(('Иван', 'Петров'),))
    first.update({'roomId': None, 'pmsRoomStayId': 'stay-1'})
    second = _stay(arr='2026-07-12T14:00', dep='2026-07-14T12:00',
                   guests=(('Ольга', 'Сидорова'),))
    second.update({'roomId': None, 'pmsRoomStayId': 'stay-2'})
    rows, _ = bookings_to_rows(
        [_book(number='group-dates', stays=[first, second])], QS, QE)
    data_path = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rooms_path = str(tmp_path / 'arrivals.xlsx')
    output_path = str(tmp_path / 'report.xlsx')
    write_data_workbook(rows, data_path)

    arrivals = openpyxl.Workbook()
    sheet = arrivals.active
    sheet.append(['№ брони', 'Заезд', 'Выезд', '№ комнаты'])
    sheet.append(['group-dates', '10.07.2026, 14:00',
                  '12.07.2026, 12:00', 31])
    sheet.append(['group-dates', '12.07.2026, 14:00',
                  '14.07.2026, 12:00', 82])
    arrivals.save(rooms_path)

    summary, _ = core.process_tchaikovsky(
        data_path, rooms_path, output_path, TL_DATA_MAP,
        {'fio': None, 'room': '№ комнаты', 'corpus': None,
         'bron': '№ брони', 'checkin': 'Заезд', 'checkout': 'Выезд'},
        extra_data=TL_EXTRA, source_map={'TravelLine YAX': (0, '')})
    ws = openpyxl.load_workbook(output_path, data_only=False)['1&2']
    rooms_by_guest = {
        ws.cell(row, 3).value: ws.cell(row, 19).value
        for row in range(2, ws.max_row)
    }
    corpora_by_guest = {
        ws.cell(row, 3).value: ws.cell(row, 20).value
        for row in range(2, ws.max_row)
    }
    assert rooms_by_guest == {'Иван Петров': 31, 'Ольга Сидорова': 82}
    # The data workbook has a Corpus column, but TravelLine left its cells
    # blank. Manual rooms still get literal corpora based on their room numbers.
    assert corpora_by_guest == {'Иван Петров': 1, 'Ольга Сидорова': 2}
    assert summary['Отсутствует корпус (КРАСНЫЙ)'] == 0


def test_manual_same_booking_and_stay_dates_remain_ambiguous(tmp_path):
    import openpyxl
    import core

    first = _stay(arr='2026-07-10T14:00', dep='2026-07-12T12:00',
                  guests=(('Иван', 'Петров'),))
    first.update({'roomId': None, 'pmsRoomStayId': 'stay-1'})
    second = _stay(arr='2026-07-10T14:00', dep='2026-07-12T12:00',
                   guests=(('Ольга', 'Сидорова'),))
    second.update({'roomId': None, 'pmsRoomStayId': 'stay-2'})
    rows, _ = bookings_to_rows(
        [_book(number='group-ambiguous', stays=[first, second])], QS, QE)
    data_path = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rooms_path = str(tmp_path / 'arrivals.xlsx')
    output_path = str(tmp_path / 'report.xlsx')
    write_data_workbook(rows, data_path)

    arrivals = openpyxl.Workbook()
    sheet = arrivals.active
    sheet.append(['№ брони', 'Заезд', 'Выезд', '№ комнаты'])
    sheet.append(['group-ambiguous', '10.07.2026, 14:00',
                  '12.07.2026, 12:00', 31])
    sheet.append(['group-ambiguous', '10.07.2026, 14:00',
                  '12.07.2026, 12:00', 82])
    arrivals.save(rooms_path)

    summary, _ = core.process_tchaikovsky(
        data_path, rooms_path, output_path, TL_DATA_MAP,
        {'fio': None, 'room': '№ комнаты', 'corpus': None,
         'bron': '№ брони', 'checkin': 'Заезд', 'checkout': 'Выезд'},
        extra_data=TL_EXTRA, source_map={'TravelLine YAX': (0, '')})
    ws = openpyxl.load_workbook(output_path, data_only=False)['1&2']
    assert [ws.cell(row, 19).value for row in range(2, ws.max_row)] == [None, None]
    assert summary['Отсутствует № комнаты (КРАСНЫЙ)'] == 2


def test_rooms_workbook_keeps_room_when_api_stay_has_no_guest_name(tmp_path):
    stay = _stay(guests=())
    stay.update({'roomId': 'room-44', 'pmsRoomStayId': 'stay-44'})
    inventory = {'accommodationInventory': {'buildings': [{
        'name': 'Корпус 1', 'floors': [{'rooms': [
            {'id': 'room-44', 'displayName': '44'},
        ]}],
    }]}}
    rows, _ = bookings_to_rows([_book(stays=[stay])], QS, QE, inventory)
    rooms_path = str(tmp_path / 'TL_tch_2026III_rooms.xlsx')
    write_rooms_workbook(rows, rooms_path)

    rooms = read_table(rooms_path)
    assert len(rooms) == 1
    assert rooms['ФИО'].isna().all()
    assert rooms.iloc[0]['№ комнаты'] == 44
    assert rooms.iloc[0]['Корпус'] == 1


def test_string_room_label_survives_excel_and_report_roundtrip(tmp_path):
    import openpyxl
    import core

    stay = _stay(guests=(('Иван', 'Петров'),))
    stay.update({'roomId': 'room-zero', 'pmsRoomStayId': 'stay-zero'})
    inventory = {'accommodationInventory': {'buildings': [{
        'name': 'Корпус 1', 'floors': [{'rooms': [
            {'id': 'room-zero', 'displayName': '001'},
        ]}],
    }]}}
    rows, _ = bookings_to_rows([_book(stays=[stay])], QS, QE, inventory)
    data_path = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rooms_path = str(tmp_path / 'TL_tch_2026III_rooms.xlsx')
    output_path = str(tmp_path / 'report.xlsx')
    write_data_workbook(rows, data_path)
    write_rooms_workbook(rows, rooms_path)

    data = read_table(data_path)
    assert data.iloc[0]['№ комнаты'] == ROOM_LABEL_TEXT_PREFIX + '001'
    core.process_tchaikovsky(
        data_path, rooms_path, output_path, TL_DATA_MAP, TL_ROOMS_MAP,
        extra_data=TL_EXTRA, source_map={'TravelLine YAX': (0, '')})
    report = openpyxl.load_workbook(output_path, data_only=False)['1']
    assert report['S2'].value == '001'
    assert report['T2'].value == 1


def test_workbooks_roundtrip_and_mapping(tmp_path):
    rows, _ = bookings_to_rows([_book()], QS, QE)
    dp = str(tmp_path / 'TL_tch_2026III_data.xlsx')
    rp = str(tmp_path / 'TL_tch_2026III_rooms.xlsx')
    write_data_workbook(rows, dp)
    write_rooms_workbook(rows, rp)
    assert is_tl_export(dp) and is_tl_export(rp)
    assert not is_tl_export(str(tmp_path / 'other.xlsx'))
    df = read_table(dp)
    assert list(df.columns) == ['Гость', 'Стоимость', 'Заезд', 'Выезд',
                                'Источник', 'Категория номера', '№ брони',
                                'Ночей', 'Тариф', 'Допуслуги', 'Наименование допуслуг',
                                'Сумма продления номера', 'Итог брони',
                                '№ комнаты', 'Корпус']
    assert df.iloc[0]['Стоимость'] == 10000.0
    assert df.iloc[0]['Гость'] == 'Иван Петров'
    headers = [str(c) for c in df.columns]
    assert mapping_valid(TL_DATA_MAP, headers, TCH_DATA_SCHEMA)
    assert mapping_valid(TL_DATA_MAP, headers, DNA_DATA_SCHEMA)
    assert all(c in headers for c in TL_EXTRA)
    dr = read_table(rp)
    assert list(dr.columns) == ['ФИО', '№ комнаты', 'Корпус',
                                'Здание TravelLine', '№ брони']
    assert len(dr) == 1
    assert dr.iloc[0]['ФИО'] == 'Иван Петров'
    assert dr['№ комнаты'].isna().all() and dr['Корпус'].isna().all()
    assert mapping_valid(TL_ROOMS_MAP, list(dr.columns), TCH_ROOMS_SCHEMA)


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


def test_fetch_details_can_also_request_pms_room_inventory(monkeypatch):
    import tl_client
    inventory = {'accommodationInventory': {'buildings': []}}
    seen, progress = [], []
    monkeypatch.setattr(tl_client, 'get_token', lambda *args: 'tok')

    def fake(method, path, token=None, params=None, data=None, ctype=None):
        seen.append(path)
        if path.endswith('/bookings'):
            return 200, {'bookingSummaries': [], 'hasMoreData': False}
        if path.endswith('/obtain-accommodation-inventory'):
            return 200, inventory
        raise AssertionError(f'unexpected API request: {path}')

    monkeypatch.setattr(tl_client, '_http', fake)
    result = tl_client.fetch_details(
        'id', 'sec', '8186', include_room_inventory=True,
        progress=lambda stage, done, total: progress.append((stage, done, total)))

    assert result == ([], inventory)
    assert seen[0] == (
        '/api/pms/v2/properties/8186/obtain-accommodation-inventory')
    assert seen[1].endswith('/bookings')
    assert ('inventory', 0, 1) in progress
    assert ('inventory', 1, 1) in progress


def test_pms_room_inventory_403_explains_permission_without_secret(monkeypatch):
    import tl_client
    monkeypatch.setattr(tl_client, 'get_token', lambda *args: 'tok')

    def fake(method, path, token=None, params=None, data=None, ctype=None):
        if path.endswith('/bookings'):
            return 200, {'bookingSummaries': [], 'hasMoreData': False}
        raise tl_client.TLApiError('HTTP 403: forbidden')

    monkeypatch.setattr(tl_client, '_http', fake)
    try:
        tl_client.fetch_details(
            'id', 'do-not-log-secret', '8186', include_room_inventory=True)
    except tl_client.TLApiError as exc:
        text = str(exc)
        assert 'PMS API/PropertyRoom' in text
        assert 'номерной фонд' in text.casefold()
        assert 'do-not-log-secret' not in text
    else:
        raise AssertionError('a denied PMS API request should fail explicitly')


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
        '№ комнаты': 117, 'Корпус': 2,
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
    assert key.startswith('tlcache-v4|')
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
