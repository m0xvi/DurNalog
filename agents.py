# -*- coding: utf-8 -*-
"""Таблица агентов отеля «Дона» и нечёткое сопоставление источников.

Извлекается из «Коммисия агентов дона.pdf». Чтобы добавить агента или
поменять комиссию — правьте список DONA_AGENTS: (комиссия %, имя, [ключи]).
"""
import re

DONA_AGENTS = [
    (20, "Балтма тур",                ["балтма", "baltma"]),
    (20, "Янтарный край",             ["янтарныйкрай", "yantarnykray"]),
    (20, "УК ГЦБиТ",                  ["гцбит", "gtsbit", "городскойцентрбронирования",
                                       "gorodskoycentrbronirovaniya"]),
    (15, "Fun&Sun ТТ Трэвел",         ["funsun", "фансан", "фансун", "funandsun"]),
    (20, "Броневик",                  ["броневик", "bronevik"]),
    (20, "Янт.Мозаика",               ["янтмозаика", "yantmozaika", "мозаика", "mozaika"]),
    (20, "Анюта",                     ["анюта", "anyuta"]),
    (20, "Академсервис",              ["академсервис", "akademservice", "akademservis"]),
    (20, "ХБ Про (ХБ Трэвел)",        ["хбпро", "hbpro", "хбтрэвел", "hbtravel", "hbtrevel"]),
    (15, "Сибукинг",                  ["сибукинг", "sibooking", "sibooking", "sibuking"]),
    (17, "101 Отель ОНЛАЙН Инновации",["101отель", "101hotel", "онлайнинновации",
                                       "onlineinnovations"]),
    (15, "Алеан Отдых",               ["алеан", "alean", "aleane"]),
    (15, "Вайт Тревел",               ["вайттрэвел", "whitetravel", "waittravel", "vaittravel"]),
    (15, "Калининградские курорты",   ["калининградски", "kaliningradsk"]),
    (15, "Бюро путешествий",          ["бюропутешествий", "byuroputeshestviy", "travelbureau"]),
    (15, "Seven Travel (ОТИС)",       ["seventravel", "севентрэвел", "отис", "otis"]),
    (15, "Островок",                  ["островок", "ostrovok", "emergingtravel"]),
    (15, "Суточно.ру",                ["суточно", "sutochno"]),
    (18, "Эй энд Эй (ЦБП)",           ["эйэндэй", "центрбронированияпутевок"]),
    (15, "Юнона",                     ["юнона", "yunona", "iunona", "junona"]),
    (15, "Expedia",                   ["expedia", "экспедиа", "экспидиа"]),
    (15, "ВиндоТур",                  ["виндотур", "vindotur", "windotour", "windotur"]),
    (20, "Яндекс",                    ["яндекс", "yandex"]),
    (15, "Дельфин",                   ["дельфин", "delfin", "delphin", "dolphin"]),
    (15, "РТБ СТМ",                   ["ртбстм", "rtbstm"]),
    (15, "Тинькоф",                   ["тинькоф", "tinkof"]),
    (15, "Леонид (Хотел Стар)",       ["хотелстар", "hotelstar", "khotelstar", "хотелстар",
                                       "леонид", "leonid"]),
    (10, "АВС Трэвел",                ["австрэвел", "abctravel", "avstravel", "абстревел"]),
    (15, "Библио Глобус",             ["библиоглобус", "biblioglobus", "biblio globus"]),
    (10, "Балтик Вояж",               ["балтиквояж", "balticvoyage", "baltikvoyazh", "balticvoyazh"]),
    (10, "Бизнес Трэвел",             ["бизнестрэвел", "businesstravel", "biznestravel"]),
    (10, "Метро",                     ["метро", "metro"]),
    (10, "Золотая Русь",              ["золотаярусь", "zolotayarush", "zolotayarus", "zolotarus"]),
    (17, "Otello (ООО ДубльГИС)",     ["otello", "отелло", "дубльгис", "dublgis", "2gis"]),
    (10, "Абразив Эксперт",           ["абразив", "abraziv", "вэлдэксперт", "veldexpert", "weldexpert"]),
    (15, "Хомлин Лэнд (Pegas)",       ["хомлин", "khomlin", "homlin", "pegas", "пегас",
                                       "pegaso", "пегасо"]),
    (15, "Тривио",                    ["тривио", "trivio"]),
    (10, "Янтарный круиз",            ["янтарныйкруиз", "yantarnykruiz"]),
    (10, "мой калининград",           ["мойкалининград", "moykaliningrad", "mykaliningrad"]),
    (15, "Привет тур",                ["приветтур", "privettur", "privettour", "privet tour"]),
    (15, "Амитур",                    ["амитур", "amitur"]),
    (15, "русские сезоны",            ["русскиесезоны", "russkiesezony", "russianseasons"]),
    (17, "Интернет Тревел",           ["интернеттрэвел", "интеренттрэвел", "internettravel",
                                       "interenttravel"]),
    (15, "Цвет",                      ["цвет", "tsvet", "tzvet", "cvet"]),
    (15, "Тинькофф путешествия",      ["тпутешествия", "tputeshestviya", "тинькоффпутешествия",
                                       "tinkoffputeshestviya", "тинькофф", "tinkoff"]),
    (15, "Ростелеком",                ["ростелеком", "rostelecom"]),
    (15, "Универсал тур",             ["универсалтур", "universaltour", "universaltur"]),
    (None, "Avia center (отели онлайн)", ["aviacenter", "авиацентр", "отелионлайн"]),
    (10, "ИП лагун = КУВА",           ["кува", "cuva", "лагун", "lagun"]),
    (15, "Русский экспресс = веди групп", ["русскийэкспресс", "russianexpress", "ведигрупп",
                                           "vedigrupp", "vedigroup"]),
    # Прямые каналы отеля: комиссия 0%, считаются распознанными (без жёлтой подсветки)
    (0, "Прямые бронирования (сайт/стойка/звонки/email)",
     ["официальныйсайт", "ageevgroup", "отстойки", "позвонку", "email"]),
]


# ================================================================
# ЕДИНАЯ МОДЕЛЬ АГЕНТОВ (редактируется в «Настройки → Агенты»)
# агент = {
#   'name':    имя,
#   'aliases': [варианты написания в отчётах],
#   'hotels':  ['dna', 'tch', 'usd'] — к каким отелям применяется,
#   'rates':   [{'from': 'дд.мм.гггг'|'', 'to': '...'|'', 'pct': число}, ...]
#              период без дат = базовая ставка; берётся период, в который
#              попадает дата заезда.
# }
# ================================================================
def default_agents():
    """Встроенная таблица (50 агентов + прямые каналы) в виде модели."""
    out = []
    for comm, name, aliases in DONA_AGENTS:
        out.append({'name': name, 'aliases': list(aliases), 'hotels': ['dna'],
                    'rates': [{'from': '', 'to': '',
                               'pct': comm if comm is not None else 0}]})
    return out


def migrate_agents(cfg):
    """Собирает список агентов из конфига: новый ключ 'agents' или старые
    ключи (overrides / agent_aliases / custom_agents)."""
    if cfg.get('agents'):
        return cfg['agents']
    agents = default_agents()
    overrides = cfg.get('agent_overrides') or {}
    extra = cfg.get('agent_aliases') or {}
    for a in agents:
        if a['name'] in overrides:
            a['rates'][0]['pct'] = float(overrides[a['name']])
        if a['name'] in extra:
            a['aliases'] += [x for x in extra[a['name']] if x not in a['aliases']]
    for ca in (cfg.get('custom_agents') or []):
        agents.append({'name': ca.get('name', ''), 'aliases': ca.get('aliases', []),
                       'hotels': ca.get('hotels', ['dna']),
                       'rates': [{'from': '', 'to': '',
                                  'pct': float(ca.get('commission', 0))}]})
    return agents


def norm_key(s):
    """Нормализация строки для нечёткого сопоставления источников/агентов."""
    return re.sub(r'[^0-9a-zа-яё]+', '', str(s).lower().replace('ё', 'е'))


# Транслитерация кириллицы в латиницу — чтобы «Островок» и «Ostrovok»
# распознавались как один агент.
_TR = {'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g', 'д': 'd', 'е': 'e', 'ё': 'e',
       'ж': 'zh', 'з': 'z', 'и': 'i', 'й': 'y', 'к': 'k', 'л': 'l', 'м': 'm',
       'н': 'n', 'о': 'o', 'п': 'p', 'р': 'r', 'с': 's', 'т': 't', 'у': 'u',
       'ф': 'f', 'х': 'h', 'ц': 'c', 'ч': 'ch', 'ш': 'sh', 'щ': 'sch',
       'ъ': '', 'ы': 'y', 'ь': '', 'э': 'e', 'ю': 'yu', 'я': 'ya'}


def translit(s):
    return ''.join(_TR.get(ch, ch) for ch in str(s))


def _similarity(a, b):
    import difflib
    return difflib.SequenceMatcher(None, a, b).ratio()


def parse_dmy(s):
    """'дд.мм.гггг' -> date | None (пустая строка -> None)."""
    from datetime import datetime
    s = (s or '').strip()
    if not s:
        return None
    for fmt in ('%d.%m.%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def rate_for(agent, d=None):
    """Комиссия агента на дату d: период from..to (включительно);
    период без дат — базовый; если ничего не подошло — базовый/первый."""
    rates = agent.get('rates') or [{'from': '', 'to': '', 'pct': 0}]
    base = None
    for r in rates:
        f, t = parse_dmy(r.get('from')), parse_dmy(r.get('to'))
        if f is None and t is None:
            if base is None:
                base = r
            continue
        if d is not None and (f is None or f <= d) and (t is None or d <= t):
            return float(r.get('pct', 0))
    if base is not None:
        return float(base.get('pct', 0))
    return float(rates[0].get('pct', 0))


def match_agent(source, agents, date=None, hotel=None):
    """
    Возвращает (комиссия % на дату date, имя агента) или (0, None).
    hotel: если задан — рассматривать только агентов, применяемых к отелю.
    Ступени: точное совпадение варианта (в т.ч. транслитом) → подстрока →
    похожие строки/токены (>= 0.85) для сокращений и опечаток.
    """
    s = norm_key(source)
    if not s:
        return 0, None
    sl = translit(s)
    raw = translit(re.sub(r'[^0-9a-zа-яё]+', ' ', str(source).lower().replace('ё', 'е')))
    tokens = [t for t in re.split(r'[^a-z0-9]+', raw) if len(t) >= 5]
    cands = [sl] + tokens

    best_score, best_agent = -1, None
    for a in agents:
        if hotel and hotel not in (a.get('hotels') or []):
            continue
        for al in (a.get('aliases') or []) + [a.get('name', '')]:
            na = norm_key(al)
            if not na:
                continue
            nal = translit(na)
            if na == s or nal == sl:
                score = 1000 + len(na)
            elif na in s or s in na or nal in sl or sl in nal:
                score = 500 + len(na)
            elif min(len(nal), len(sl)) >= 5:
                r = max(_similarity(nal, c) for c in cands)
                if r < 0.85:
                    continue
                score = int(r * 100)
            else:
                continue
            if score > best_score:
                best_score, best_agent = score, a

    if best_agent is None:
        return 0, None
    return rate_for(best_agent, date), best_agent['name']


def match_dona_agent(source, overrides=None, extra_aliases=None,
                     custom_agents=None, date=None):
    """Совместимая обёртка над match_agent для старого API/тестов."""
    agents = default_agents()
    for a in agents:
        if overrides and a['name'] in overrides:
            a['rates'][0]['pct'] = float(overrides[a['name']])
        if extra_aliases and a['name'] in extra_aliases:
            a['aliases'] += list(extra_aliases[a['name']])
    for ca in (custom_agents or []):
        agents.append({'name': ca.get('name', ''), 'aliases': ca.get('aliases', []),
                       'hotels': ['dna'],
                       'rates': [{'from': '', 'to': '',
                                  'pct': float(ca.get('commission', 0))}]})
    return match_agent(source, agents, date=date)


# ================================================================
# ИМПОРТ / ЭКСПОРТ ТАБЛИЦЫ АГЕНТОВ (Excel / CSV)
# Колонки: Название | Отели | Синонимы | С даты | По дату | Комиссия %
# Несколько строк с одним названием = периоды комиссии.
# ================================================================
def agents_to_rows(agents):
    rows = []
    for a in agents:
        for r in (a.get('rates') or [{'from': '', 'to': '', 'pct': 0}]):
            rows.append({
                'Название': a['name'],
                'Отели': '|'.join(a.get('hotels') or []),
                'Синонимы': ', '.join(a.get('aliases') or []),
                'С даты': r.get('from', ''),
                'По дату': r.get('to', ''),
                'Комиссия %': r.get('pct', 0),
            })
    return rows


def rows_to_agents(rows):
    """rows: список dict с теми же ключами -> список агентов."""
    by_name = {}
    order = []
    for row in rows:
        name = str(row.get('Название') or '').strip()
        if not name:
            continue
        if name not in by_name:
            by_name[name] = {'name': name, 'aliases': [], 'hotels': ['dna'], 'rates': []}
            order.append(name)
        a = by_name[name]
        hotels = str(row.get('Отели') or '').replace(';', '|').split('|')
        hotels = [h.strip() for h in hotels if h.strip()]
        if hotels:
            a['hotels'] = hotels
        syn = re.split(r'[,;]+', str(row.get('Синонимы') or ''))
        for x in syn:
            x = x.strip()
            if x and x not in a['aliases']:
                a['aliases'].append(x)
        try:
            pct = float(str(row.get('Комиссия %') or '0').strip().replace(',', '.'))
        except ValueError:
            pct = 0.0
        a['rates'].append({'from': str(row.get('С даты') or '').strip(),
                           'to': str(row.get('По дату') or '').strip(),
                           'pct': pct})
    for a in by_name.values():
        if not a['rates']:
            a['rates'] = [{'from': '', 'to': '', 'pct': 0}]
    return [by_name[n] for n in order]


def export_agents(agents, path):
    import pandas as pd
    df = pd.DataFrame(agents_to_rows(agents),
                      columns=['Название', 'Отели', 'Синонимы', 'С даты', 'По дату', 'Комиссия %'])
    if str(path).lower().endswith('.csv'):
        df.to_csv(path, index=False, encoding='utf-8-sig', sep=';')
    else:
        df.to_excel(path, index=False)


def import_agents(path):
    import pandas as pd
    if str(path).lower().endswith('.csv'):
        df = None
        for enc in ('utf-8-sig', 'utf-8', 'cp1251'):
            try:
                df = pd.read_csv(path, encoding=enc, sep=';')
                break
            except (UnicodeDecodeError, Exception):
                continue
        if df is None:
            raise ValueError("Не удалось прочитать CSV")
    else:
        df = pd.read_excel(path)
    df = df.fillna('')
    return rows_to_agents(df.to_dict('records'))
