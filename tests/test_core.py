# -*- coding: utf-8 -*-
import os
import sys
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from agents import match_dona_agent
from tax import parse_price, dates_overlap, quarter_of, tax_formula_method
from db import HistoryDB

D = date


def test_match_agents():
    assert match_dona_agent("Ostrovok.ru (Emerging Travel Group)") == (15, "Островок")
    assert match_dona_agent("Bronevik.com") == (20, "Броневик")
    assert match_dona_agent("Cuva") == (10, "ИП лагун = КУВА")
    assert match_dona_agent("Т-Путешествия") == (15, "Тинькофф путешествия")
    assert match_dona_agent("yandex.travel") == (20, "Яндекс")
    assert match_dona_agent("прямая бронь (сайт)") == (0, None)


def test_parse_price():
    assert parse_price(5000) == 5000.0
    assert parse_price("5 000,00") == 5000.0
    assert parse_price("12 345.6 руб.") == 12345.6
    assert parse_price(None) is None
    assert parse_price("") is None
    assert parse_price("нет") is None


def test_dates_and_quarter():
    assert dates_overlap(D(2026, 3, 30), D(2026, 4, 2), D(2026, 3, 30), D(2026, 4, 2))
    assert dates_overlap(D(2026, 3, 28), D(2026, 4, 1), D(2026, 3, 30), D(2026, 4, 5))
    # смежные ночи — НЕ пересечение (две разные брони подряд)
    assert not dates_overlap(D(2026, 3, 30), D(2026, 4, 2), D(2026, 4, 2), D(2026, 4, 5))
    assert not dates_overlap(None, D(2026, 4, 2), D(2026, 3, 30), D(2026, 4, 2))
    assert quarter_of(D(2026, 1, 1)) == "2026-Q1"
    assert quarter_of(D(2026, 4, 1)) == "2026-Q2"
    assert quarter_of(D(2026, 12, 31)) == "2026-Q4"
    assert "5100" in tax_formula_method("J", 2)


def _mk_db(tmp_path):
    return HistoryDB(os.path.join(str(tmp_path), "test.db"))


def _booking(guest="Иванов Иван", bron="", ci="2026-03-30", co="2026-04-02"):
    return {"guest": guest, "guest_norm": guest.lower(), "bron": bron,
            "checkin": ci, "checkout": co, "price": 9000, "price_full": None,
            "nights": 3, "dup_note": None}


def test_db_duplicates(tmp_path):
    db = _mk_db(tmp_path)
    rid1 = db.add_report("Дона", "2026-Q1", "q1.csv", b"orig1", "q1.xlsx", b"out1", 1, 0)
    db.add_bookings(rid1, [_booking()])

    # та же бронь (ФИО+даты) во втором квартале -> дубль
    notes = db.check_duplicates("Дона", [{
        "bron": "", "guest_norm": "иванов иван",
        "checkin": D(2026, 3, 30), "checkout": D(2026, 4, 2)}])
    assert notes[0] and "2026-Q1" in notes[0]

    # другой гость -> не дубль
    notes = db.check_duplicates("Дона", [{
        "bron": "", "guest_norm": "петров",
        "checkin": D(2026, 3, 30), "checkout": D(2026, 4, 2)}])
    assert notes[0] is None

    # тот же гость, но другие (несмежные) даты -> не дубль
    notes = db.check_duplicates("Дона", [{
        "bron": "", "guest_norm": "иванов иван",
        "checkin": D(2026, 5, 1), "checkout": D(2026, 5, 3)}])
    assert notes[0] is None

    # по № брони — даже без пересечения дат
    rid2 = db.add_report("Дона", "2026-Q2", "q2.csv", b"o2", "q2.xlsx", b"o2", 1, 0)
    db.add_bookings(rid2, [_booking(guest="Сидоров", bron="D-777",
                                   ci="2026-04-10", co="2026-04-12")])
    notes = db.check_duplicates("Дона", [{
        "bron": "D-777", "guest_norm": "другой",
        "checkin": D(2026, 6, 1), "checkout": D(2026, 6, 2)}])
    assert notes[0] and "2026-Q2" in notes[0]

    # другой отель не пересекается
    notes = db.check_duplicates("Чайковский", [{
        "bron": "", "guest_norm": "иванов иван",
        "checkin": D(2026, 3, 30), "checkout": D(2026, 4, 2)}])
    assert notes[0] is None

    # файлы отчёта достаются из БД
    name, blob = db.get_file(rid1, "orig")
    assert name == "q1.csv" and blob == b"orig1"

    # отчёт того же квартала находится и удаляется
    assert db.find_same("Дона", "2026-Q1")["id"] == rid1
    db.delete_report(rid1)
    assert db.find_same("Дона", "2026-Q1") is None
    db.close()


def test_dona_end_to_end(tmp_path):
    import pandas as pd
    import openpyxl
    import core

    src = os.path.join(str(tmp_path), "dona.csv")
    out = os.path.join(str(tmp_path), "dona.xlsx")
    pd.DataFrame([
        ["Иванов Иван", 8500, "30.03.2026", "02.04.2026", "Ostrovok.ru", "D-1", 3, "Стандарт"],
        ["Иванов Иван", 8500, "30.03.2026", "02.04.2026", "Ostrovok.ru", "D-1", 3, "Стандарт"],
    ], columns=["Гость", "Стоимость", "Заезд", "Выезд", "Источник", "№ брони",
                "Ночей", "Категория номера"]).to_csv(src, index=False)

    # второй «отчёт» с той же бронью -> checker помечает дубль
    def checker(rows):
        return ["Переход кварталов: тест" if r["guest_norm"] == "иванов иван" else None
                for r in rows]

    summary, bookings = core.process_dona(src, out, dup_checker=checker)
    assert summary["Переход кварталов (ФИОЛЕТОВЫЙ)"] == 2
    assert all(b["dup_note"] for b in bookings)
    ws = openpyxl.load_workbook(out)["Дона"]
    assert ws["P2"].value == "Переход кварталов: тест"
    assert ws["P2"].fill.start_color.rgb.endswith("CC99FF")
    assert ws["E2"].value == 15
    assert ws["G2"].value == "=ROUND(D2/(100-E2)*100,2)"


def test_tchaikovsky_end_to_end(tmp_path):
    import pandas as pd
    import openpyxl
    import core

    src = os.path.join(str(tmp_path), "tch.csv")
    rooms = os.path.join(str(tmp_path), "rooms.csv")
    out = os.path.join(str(tmp_path), "tch.xlsx")
    pd.DataFrame([
        ["Иванов Иван", "5 000,00", "01.08.2026", "04.08.2026", "Сайт", "Стандарт", "B-1", 3],
    ], columns=["Гость", "Стоимость", "Заезд", "Выезд", "Источник",
                "Категория номера", "№ брони", "Ночей"]).to_csv(src, index=False)
    pd.DataFrame([["Иванов Иван", 12]], columns=["ФИО", "№ комнаты"]).to_csv(rooms, index=False)

    summary, bookings = core.process_tchaikovsky(src, rooms, out)
    assert summary["Отсутствует Стоимость (ГОЛУБОЙ)"] == 0  # "5 000,00" распознано
    ws = openpyxl.load_workbook(out)["1&2"]
    assert ws["D2"].value == 5000.0
    assert ws["G2"].value == '=ROUND(D2/(100-E2)*100,2)'
    assert ws["K2"].value == '=IF(J2>=5100,"ставка","мин.ставка")'
    assert ws["L2"].value == '=ROUND(IF(K2="мин.ставка",I2*100,G2/102*2),2)'
    assert ws["R1"].value == "Примечание"
    assert bookings[0]["price"] == 5000.0


def test_guess_mapping_renamed_headers():
    from schemas import TCH_DATA_SCHEMA
    from core import guess_mapping, mapping_valid
    headers = ["ФИО гостя", "Сумма", "Дата заезда", "Дата выезда",
               "Источник бронирования", "Категория", "ID брони", "Кол-во ночей",
               "Лишняя колонка"]
    m = guess_mapping(headers, TCH_DATA_SCHEMA)
    assert m['guest'] == "ФИО гостя"
    assert m['price'] == "Сумма"
    assert m['checkin'] == "Дата заезда"
    assert m['checkout'] == "Дата выезда"
    assert m['source'] == "Источник бронирования"
    assert m['bron'] == "ID брони"
    assert m['nights'] == "Кол-во ночей"
    assert mapping_valid(m, headers, TCH_DATA_SCHEMA)


def test_dona_extra_columns_and_custom_mapping(tmp_path):
    import pandas as pd
    import openpyxl
    import core

    src = os.path.join(str(tmp_path), "dona2.csv")
    out = os.path.join(str(tmp_path), "dona2.xlsx")
    pd.DataFrame([
        ["Иванов Иван", 8500, "01.08.2026", "04.08.2026", "островок", "Менеджер"],
    ], columns=["Гость", "Цена", "Прибытие", "Убытие", "Канал", "Менеджер"]
    ).to_csv(src, index=False)

    mapping = {"guest": "Гость", "price": "Цена", "checkin": "Прибытие",
               "checkout": "Убытие", "source": "Канал", "category": None,
               "bron": None, "nights": None}
    summary, bookings = core.process_dona(src, out, map_data=mapping,
                                              extra_data=["Менеджер"])
    ws = openpyxl.load_workbook(out)["Дона"]
    assert ws["Q1"].value == "Менеджер"          # доп. колонка в шапке
    assert ws["Q2"].value == "Менеджер"          # значение перенесено
    assert ws["D2"].value == 8500
    assert ws["E2"].value == 15                  # островок распознан
    assert summary["Отсутствует Стоимость (ГОЛУБОЙ)"] == 0


def test_tax_params_custom():
    from tax import TaxParams
    t = TaxParams()  # значения по умолчанию = прежние формулы
    assert t.f_daily('D', 'F', 2) == '=D2/F2'
    assert t.f_method('G', 2) == '=IF(G2>=5100,"ставка","мин.ставка")'
    assert t.f_tax('H', 'F', 'D', 2) == '=ROUND(IF(H2="мин.ставка",F2*100,D2/102*2),2)'
    t3 = TaxParams({'rate': 3, 'min_per_night': 200})
    assert t3.f_tax('H', 'F', 'D', 2) == '=ROUND(IF(H2="мин.ставка",F2*200,D2/103*3),2)'
    assert abs(t3.auto_threshold() - 200 * 103 / 3) < 1e-9
    t_not_incl = TaxParams({'incl': False, 'auto_threshold': True})
    assert t_not_incl.f_tax('H', 'F', 'D', 2) == '=ROUND(IF(H2="мин.ставка",F2*100,D2/100*2),2)'
    # свой шаблон формулы
    t_tpl = TaxParams({'tpl_tax': '=ROUND({N}*{rate},2)'})
    assert t_tpl.f_tax('H', 'F', 'D', 5) == '=ROUND(F5*2,2)'


def test_agent_overrides():
    from agents import match_dona_agent
    assert match_dona_agent("островок", {"Островок": 25}) == (25, "Островок")
    assert match_dona_agent("островок") == (15, "Островок")


def test_dona_custom_tax_in_output(tmp_path):
    import pandas as pd
    import openpyxl
    import core
    src = os.path.join(str(tmp_path), "d3.csv")
    out = os.path.join(str(tmp_path), "d3.xlsx")
    pd.DataFrame([["Иванов Иван", 8500, "01.08.2026", "04.08.2026", "островок"]],
                 columns=["Гость", "Стоимость", "Заезд", "Выезд", "Источник"]
                 ).to_csv(src, index=False)
    _, _ = core.process_dona(src, out, tax={'rate': 3})
    ws = openpyxl.load_workbook(out)["Дона"]
    assert ws["L2"].value == '=ROUND(IF(K2="мин.ставка",I2*100,G2/103*3),2)'


def test_match_cross_language_and_typos():
    from agents import match_dona_agent as m
    # другой язык (транслитерация)
    assert m("Ostrovok.ru")[1] == "Островок"
    assert m("ostrovok")[1] == "Островок"
    assert m("SUTOCHNO")[1] == "Суточно.ру"
    assert m("Yandex Travel")[1] == "Яндекс"
    assert m("Bronevik")[1] == "Броневик"
    # опечатка / сокращение
    assert m("Ostrovk.ru")[1] == "Островок"
    assert m("Янт.Мозаика")[1] == "Янт.Мозаика"
    # исправленная комиссия АВС Трэвел
    assert m("АВС Трэвел") == (10, "АВС Трэвел")
    # по-прежнему не распознаёт постороннее
    assert m("прямая бронь (сайт)") == (0, None)


def test_match_user_aliases():
    from agents import match_dona_agent as m
    assert m("ООО Ромашка", None, {"Островок": ["ромашка"]}) == (15, "Островок")
    assert m("ООО Ромашка") == (0, None)


def test_direct_channels_and_custom_agents():
    from agents import match_dona_agent as m
    direct = "Прямые бронирования (сайт/стойка/звонки/email)"
    assert m("Официальный сайт\nageevgroup.ru") == (0, direct)
    assert m("Бронирование из шахматки (От стойки)") == (0, direct)
    assert m("Бронирование из шахматки (По звонку)") == (0, direct)
    assert m("Бронирование из шахматки (Email)") == (0, direct)
    # действительно неизвестные — остаются нераспознанными (пойдут в диалог)
    assert m("Ozon (new)")[1] is None
    assert m("HotelBook")[1] is None
    assert m("Acase.ru (ООО \"АОН.РУ\")")[1] is None
    assert m("Roomlink (Zabroniryi.ru)")[1] is None
    assert m("OneTwoTrip")[1] is None
    assert m("Trip.com Group")[1] is None
    assert m("Бронирование из шахматки (Туроператор)")[1] is None
    # кастомный агент, добавленный пользователем
    cust = [{'name': 'Ozon Travel', 'commission': 15, 'aliases': ['ozon']}]
    assert m("Ozon (new)", None, None, cust) == (15, 'Ozon Travel')


def test_floating_rates_by_date():
    from agents import default_agents, match_agent, rate_for
    agents = default_agents()
    ost = next(a for a in agents if a['name'] == 'Островок')
    ost['rates'] = [{'from': '', 'to': '31.03.2026', 'pct': 15},
                    {'from': '01.04.2026', 'to': '', 'pct': 17}]
    from datetime import date
    assert rate_for(ost, date(2026, 3, 30)) == 15
    assert rate_for(ost, date(2026, 4, 1)) == 17
    assert rate_for(ost, None) == 15  # без даты — базовый (первый открытый? первый)
    c, n = match_agent("Ostrovok.ru", agents, date=date(2026, 5, 10))
    assert (c, n) == (17, 'Островок')
    c, n = match_agent("Ostrovok.ru", agents, date=date(2026, 1, 10))
    assert (c, n) == (15, 'Островок')


def test_agents_hotel_filter():
    from agents import default_agents, match_agent
    agents = default_agents()
    # агент Усадьбы не распознаётся в Доне и наоборот
    assert match_agent("Комфорт Букинг", agents, hotel='dna')[1] is None
    assert match_agent("Комфорт Букинг", agents, hotel='usd')[1] is not None
    # у Чайковского свой Островок
    assert match_agent("Ostrovok.ru", agents, hotel='tch')[1] is not None
    assert match_agent("Ostrovok.ru", agents, hotel='dna')[1] == 'Островок'


def test_import_export_roundtrip(tmp_path):
    from agents import default_agents, export_agents, import_agents
    agents = default_agents()
    agents.append({'name': 'Ozon Travel', 'aliases': ['ozon'], 'hotels': ['dna'],
                   'rates': [{'from': '', 'to': '31.03.2026', 'pct': 12},
                             {'from': '01.04.2026', 'to': '', 'pct': 15}]})
    p_x = str(tmp_path / "agents.xlsx")
    p_c = str(tmp_path / "agents.csv")
    export_agents(agents, p_x)
    export_agents(agents, p_c)
    for p in (p_x, p_c):
        back = import_agents(p)
        oz = next(a for a in back if a['name'] == 'Ozon Travel')
        assert len(oz['rates']) == 2 and oz['rates'][1]['pct'] == 15
        assert 'ozon' in oz['aliases']
        ost = next(a for a in back if a['name'] == 'Островок')
        assert ost['rates'][0]['pct'] == 15


def test_users_db(tmp_path):
    from db import HistoryDB
    db = HistoryDB(os.path.join(str(tmp_path), "u.db"))
    db.ensure_default_admin()
    assert db.check_user('admin', 'admin') is not None
    assert db.check_user('admin', 'nope') is None
    assert db.add_user('dona_user', 'pw', 'hotel', 'dna')
    assert not db.add_user('dona_user', 'x', 'hotel', 'dna')  # логин занят
    u = db.check_user('dona_user', 'pw')
    assert u['role'] == 'hotel' and u['hotel'] == 'dna'
    admin_id = [x for x in db.list_users() if x['login'] == 'admin'][0]['id']
    assert db.del_user(admin_id) is False  # последний админ неприкосновенен
    du = [x for x in db.list_users() if x['login'] == 'dona_user'][0]['id']
    assert db.del_user(du) is True
    db.set_password(admin_id, 'newpass')
    assert db.check_user('admin', 'newpass') is not None
    db.close()


def test_use_full_off(tmp_path):
    import pandas as pd
    import openpyxl
    import core
    src = os.path.join(str(tmp_path), "uf.csv")
    out = os.path.join(str(tmp_path), "uf.xlsx")
    pd.DataFrame([["Иванов Иван", 8500, "01.08.2026", "04.08.2026", "островок"]],
                 columns=["Гость", "Стоимость", "Заезд", "Выезд", "Источник"]
                 ).to_csv(src, index=False)
    _, _ = core.process_dona(src, out, tax={'use_full': False})
    ws = openpyxl.load_workbook(out)["Дона"]
    # налог от исходной стоимости D, а не от пересчитанной G
    assert ws["L2"].value == '=ROUND(IF(K2="мин.ставка",I2*100,D2/102*2),2)'
    assert ws["J2"].value == '=D2/I2'


def test_db_path_override(tmp_path):
    from db import HistoryDB, set_db_path, default_db_path
    p = os.path.join(str(tmp_path), "shared", "x.db")
    os.makedirs(os.path.dirname(p))
    set_db_path(p)
    try:
        assert default_db_path() == p
        db = HistoryDB()
        db.ensure_default_admin()
        db.close()
        assert os.path.exists(p)
        db2 = HistoryDB(p)
        assert db2.check_user('admin', 'admin') is not None
        db2.close()
    finally:
        set_db_path(None)
