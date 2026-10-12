# -*- coding: utf-8 -*-
"""Таблица агентов отеля «Дона» и нечёткое сопоставление источников.

Извлекается из «Коммисия агентов дона.pdf». Чтобы добавить агента или
поменять комиссию — правьте список DONA_AGENTS: (комиссия %, имя, [ключи]).
Комиссия — число, либо [(с, по, %)] для плавающей ставки; СКИДКА по договору = 0%.
"""
import re

DONA_AGENTS = [
    (20, "Балтма тур",                ["балтма", "baltma"]),
    (20, "Янтарный край",             ["янтарныйкрай", "yantarnykray"]),
    (20, "УК ГЦБиТ",                  ["гцбит", "gtsbit", "городскойцентрбронирования",
                                       "gorodskoycentrbronirovaniya"]),
    (15, "Fun&Sun ТТ Трэвел",         ["funsun", "фансан", "фансун", "funandsun"]),
    (20, "Броневик",                  ["броневик", "bronevik", "компанияброневик",
                                       "компания броневик"]),
    (20, "Янт.Мозаика",               ["янтмозаика", "yantmozaika", "мозаика", "mozaika"]),
    (20, "Анюта",                     ["анюта", "anyuta"]),
    (20, "Академсервис",              ["академсервис", "akademservice", "akademservis",
                                       "академонлайн", "академ онлайн", "akademonline",
                                       "aonru", "аонру", "acase"]),
    (20, "ХБ Про (ХБ Трэвел)",        ["хбпро", "hbpro", "хбтрэвел", "hbtravel", "hbtrevel"]),
    (15, "Сибукинг",                  ["сибукинг", "sibooking", "sibooking", "sibuking"]),
    (17, "101 Отель ОНЛАЙН Инновации",["101отель", "101hotel", "онлайнинновации",
                                       "onlineinnovations"]),
    # СКИДКА по договору -> комиссия 0%
    (0, "Алеан Отдых",               ["алеан", "alean", "aleane", "системаалеан",
                                       "система алеан"]),
    (15, "Вайт Тревел",               ["вайттрэвел", "whitetravel", "waittravel", "vaittravel"]),
    (15, "Калининградские курорты",   ["калининградски", "kaliningradsk"]),
    (15, "Бюро путешествий",          ["бюропутешествий", "byuroputeshestviy", "travelbureau"]),
    (15, "Seven Travel (ОТИС)",       ["seventravel", "севентрэвел", "отис", "otis"]),
    # Островок +1 п.п. с 01.08.2026 по договору
    ([("", "31.07.2026", 15), ("01.08.2026", "", 16)], "Островок",
     ["островок", "ostrovok", "emergingtravel",
      "бронированиегостиниц", "бронирование гостиниц"]),
    (15, "Суточно.ру",                ["суточно", "sutochno"]),
    (18, "Эй энд Эй (ЦБП)",           ["эйэндэй", "центрбронированияпутевок",
                                       "zabroniryi", "забронируй",
                                       "roomlink", "румлинк"]),
    (15, "Юнона",                     ["юнона", "yunona", "iunona", "junona"]),
    (15, "Expedia",                   ["expedia", "экспедиа", "экспидиа"]),
    (15, "ВиндоТур",                  ["виндотур", "vindotur", "windotour", "windotur"]),
    (20, "Яндекс",                    ["яндекс", "yandex"]),
    (15, "Дельфин",                   ["дельфин", "delfin", "delphin", "dolphin"]),
    # Туроператор Дельфин отделён: по договору СКИДКА -> 0%
    (0, "Туроператор Дельфин",        ["туроператордельфин", "туроператор дельфин"]),
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
    # СКИДКА по договору -> комиссия 0%
    (0, "Хомлин Лэнд (Pegas)",       ["хомлин", "khomlin", "homlin", "pegas", "пегас",
                                       "pegaso", "пегасо", "пегасспб", "пегас спб"]),
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
    # --- из отчёта по агентам за Q3 2026 (ставки проверить в Настройках!) ---
    (15, "Ozon Travel",               ["ozon", "озон", "ozontravel"]),
    (15, "Туту.ру",                   ["tutu", "туту", "tutu.ru",
                                       "новые туристические технологии",
                                       "comfortbooking", "комфортбукинг",
                                       "агрегаторканалов"]),
    (15, "Trip.com",                  ["tripcom", "trip.com", "трипком"]),
    (15, "OneTwoTrip",                ["onetwotrip", "one two trip", "12т"]),
    (15, "HotelBook",                 ["hotelbook", "хотелбук"]),
    (0, "TravelLine",                 ["travelline", "трэвеллайн", "тревеллайн",
                                       "трэвел лайн", "tlintegration"]),
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
    """Встроенные таблицы агентов по отелям в виде модели."""
    out = []
    for comm, name, aliases in DONA_AGENTS:
        if isinstance(comm, list):  # плавающая ставка: [(с, по, %)]
            rates = [{'from': f, 'to': t, 'pct': p} for f, t, p in comm]
        else:
            rates = [{'from': '', 'to': '',
                      'pct': comm if comm is not None else 0}]
        out.append({'name': name, 'aliases': list(aliases), 'hotels': ['dna'],
                    'rates': rates})
    for name, spec, aliases in TCH_AGENTS:
        channel_aliases = TCH_TL_CHANNEL_ALIASES.get(name, [])
        out.append({'name': name + ' [Чайковский]',
                    'aliases': list(aliases) + list(channel_aliases),
                    'hotels': ['tch'], 'rates': _tch_rates(spec)})
    for name, pct, aliases in USADBA_AGENTS:
        out.append({'name': name + ' [Усадьба]', 'aliases': list(aliases),
                    'hotels': ['usd'],
                    'rates': [{'from': '', 'to': '', 'pct': pct}]})
    return out


def migrate_agents(cfg):
    """Собирает список агентов из конфига: новый ключ 'agents' или старые
    ключи (overrides / agent_aliases / custom_agents)."""
    if cfg.get('agents'):
        agents = cfg['agents']
        builtins = default_agents()
        direct_names = {
            'TravelLine Booking Engine [Чайковский]',
            'Прямые бронирования (стойка) [Чайковский]',
        }
        for builtin in builtins:
            tl_aliases = [x for x in builtin.get('aliases', [])
                          if str(x).startswith('TravelLine ')]
            if not tl_aliases:
                continue
            current = next((a for a in agents
                            if a.get('name') == builtin['name']), None)
            if current is None:
                if builtin['name'] in direct_names:
                    agents.append({
                        'name': builtin['name'],
                        'aliases': list(builtin['aliases']),
                        'hotels': list(builtin['hotels']),
                        'rates': [dict(r) for r in builtin['rates']],
                    })
                continue
            aliases = current.setdefault('aliases', [])
            aliases.extend(x for x in tl_aliases if x not in aliases)
        return agents
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
    # TravelLine source strings all share a long vendor prefix. Fuzzy-matching
    # on the token "travelline" would make an unknown code (e.g. CMB) look like
    # an unrelated known agent. TL labels must match a registered full alias.
    tl_source = s.startswith('travelline')

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
            elif not tl_source and min(len(nal), len(sl)) >= 5:
                r = 0.0
                for c in cands:
                    # для отдельных слов-токенов порог строже, чем для строки целиком
                    rr = _similarity(nal, c)
                    if rr >= (0.85 if c == sl else 0.9) and rr > r:
                        r = rr
                if r <= 0:
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
    return match_agent(source, agents, date=date, hotel='dna')


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


# ================================================================
# АГЕНТЫ ОТЕЛЯ «ЧАЙКОВСКИЙ» (из таблицы отеля)
# ================================================================
TCH_AGENTS = [
    ("Яндекс.Путешествия", [("до 31.01.2026", 15), ("с 01.02.2026", 18)],
     ["яндекс", "yandex", "яндекс.путешествия"]),
    ("ООО «Интернет Трэвел» (OZON)", [("base", 17)], ["ozon", "озон", "интернет трэвел", "интерент трэвел"]),
    ("ООО «ХОМЛИН ЛЭНД» (ПЕГАС)", [("base", 0)], ["пегас", "pegas", "пегасо", "pegaso", "хомлин", "khomlin"]),  # СКИДКА -> 0%
    ("ООО \"АЛЕАН ОТДЫХ\"", [("base", 0)], ["алеан", "alean"]),  # СКИДКА -> 0%
    ("ООО \"ОНЛАЙН ИННОВАЦИИ\" (101 Отель)", [("base", 17)], ["101 отель", "101hotel", "101hotels", "онлайн инновации"]),
    ("ООО \"БРОНИРОВАНИЕ ГОСТИНИЦ\" (OSTROVOK)", [("до 31.07.2026", 15), ("с 01.08.2026", 16)], ["островок", "ostrovok"]),  # +1 п.п. с 01.08.26
    ("ООО «12Т» (ранее ONE TWO TRIP)", [("base", 15)], ["12т", "one two trip", "onetwotrip", "вайт тревел", "white travel"]),
    ("ООО \"КОМПАНИЯ БРОНЕВИК\"", [("base", 20)], ["броневик", "bronevik"]),
    ("ООО \"Эй энд Эй\"", [("base", 18)], ["эй энд эй", "эйэндэй", "забронируй"]),
    ("ООО \"ТРИВИО СЕРВИС\"", [("base", 18)], ["тривио", "trivio"]),
    ("ООО \"ХБ-ПРО\"", [("base", 20)], ["хб-про", "хб про", "hb-pro", "hbpro", "хб трэвел"]),
    ("ООО \"АКАДЕМ-ОНЛАЙН\"", [("base", 15)], ["академ", "академсервис", "akadem", "академ-онлайн"]),
    ("ООО \"УК ГЦБ и Т\"", [("base", 20)], ["гцбит", "гцб", "городской центр бронирования"]),
    ("ООО \"Библио-Глобус Туроператор\"", [("base", 15)], ["библио", "библиоглобус", "biblio"]),
    ("ООО «Леонн»", [("base", 15)], ["леонн", "leonн", "leonn"]),
    ("ООО \"Кува\"", [("base", 15)], ["кува", "cuva", "лагун", "lagun"]),
    ("ООО \"Дельфин-Центр\"", [("base", 15)], ["дельфин-центр", "дельфин центр", "delfin-centr"]),
    ("ООО «Янтарный край»", [("base", 15)], ["янтарный край", "yantarny kray"]),
    ("ООО \"БАЛТМА ТУР\"", [("base", 20)], ["балтма", "baltma"]),
    ("ООО «Янтарная мозаика-Надежда»", [("base", 15)], ["янтарная мозаика", "мозаика", "mozaika"]),
    ("ООО «ТУРНЕ ТРАНС ЭКСПРЕСС»", [("base", 20)], ["турне транс", "turne trans", "турнетранс"]),
    ("Суточно.ру", [("base", 15)], ["суточно", "sutochno"]),
    ("ООО «Дубль Гис» (OTELLO)", [("до 31.03.2026", 15), ("01.04.2026-31.07.2026", 17), ("с 01.08.2026", 20)],
     ["otello", "отелло", "дубль гис", "dublgis", "2gis"]),
    ("ООО Тинькофф Путешествия", [("base", 15)], ["тинькофф", "tinkoff", "т-путешествия"]),
    ("ООО «Випсервис Трэвел»", [("base", 15)], ["випсервис", "vipservis", "бизнессервис решения"]),
    ("ООО «Компания Отели Онлайн»", [("base", 15)], ["avia-centre", "aviacentre", "отели онлайн", "avia center"]),
    ("ТТ-ТРЭВЕЛ (FUN&SUN)", [("base", 15)], ["fun&sun", "funsun", "тт-трэвел", "тт трэвел", "fun sun"]),
    ("ООО «Туроператор Дельфин»", [("base", 0)], ["туроператор дельфин", "delfin center tour"]),  # СКИДКА -> 0%
    ("ГАУК «МОСГОРТУР»", [("base", 15)], ["мосгортур", "mosgortur"]),
    ("ООО «ЦВЕТО»", [("base", 15)], ["цвето", "broni.travel", "брони.травел"]),
    ("ООО «Бюро Центр»", [("base", 10)], ["бюро центр", "byuro centr"]),
    ("\"Калининградские курорты\"", [("base", 10)], ["калининградские курорты"]),
    ("«В 39 регион»", [("base", 10)], ["в 39 регион", "в39регион"]),
    ("«ЖилиБыли»", [("до 26.06.2026", 0), ("с 27.06.2026", 15)], ["жилибыли", "жили-были"]),
    ("ООО «СБО» (Ростелеком)", [("base", 15)], ["сбо", "rostelecom", "ростелеком", "travel tech"]),
    ("Trip.com Group", [("base", 15)], ["trip.com", "tripcom", "trip com"]),
    # TravelLine Booking Engine public tariff is 4%; property-specific rate is
    # editable in Settings → Agents. Reception/PMS bookings are direct, 0%.
    ("TravelLine Booking Engine", [("base", 4)], ["TravelLine сайт"]),
    ("Прямые бронирования (стойка)", [("base", 0)],
     ["TravelLine ресепшен"]),
]


# ================================================================
# АГЕНТЫ ОТЕЛЯ «УСАДЬБА» (из таблицы отеля; спорные рукописные значения
# помечены комментарием — проверьте в «Настройки → Агенты»)
# ================================================================
# TravelLine Read Reservation source labels -> existing Chai agent records.
# These are aliases only: the agent's configured commission/rate schedule remains
# authoritative. CMB is intentionally omitted: Comfort Booking's total fee is
# 2% plus the underlying channel's commission and cannot be inferred from CMB.
TCH_TL_CHANNEL_ALIASES = {
    'ООО "Эй энд Эй"': ['TravelLine ANA'],       # Roomlink / Zabroniryi
    'ООО "ТРИВИО СЕРВИС"': ['TravelLine TIO'],
    'ООО «ХОМЛИН ЛЭНД» (ПЕГАС)': ['TravelLine PGS'],
    'ООО Тинькофф Путешествия': ['TravelLine TNK'],
    'ООО «СБО» (Ростелеком)': ['TravelLine RTT'],
    'ООО "АЛЕАН ОТДЫХ"': ['TravelLine ALN'],
    'ООО «Дубль Гис» (OTELLO)': ['TravelLine OL2'],
    'ООО "АКАДЕМ-ОНЛАЙН"': ['TravelLine ACS'],
    'ООО "Библио-Глобус Туроператор"': ['TravelLine BGL'],
    'Яндекс.Путешествия': ['TravelLine YAX'],
    'ООО "ХБ-ПРО"': ['TravelLine UTS'],         # Hotelbook
    'ООО «Интернет Трэвел» (OZON)': ['TravelLine OZ2'],
    'ООО "ОНЛАЙН ИННОВАЦИИ" (101 Отель)': ['TravelLine H1C'],
    'ООО «12Т» (ранее ONE TWO TRIP)': ['TravelLine OTT'],
    'Trip.com Group': ['TravelLine CTP'],
    'ООО "УК ГЦБ и Т"': ['TravelLine CBC'],
    'ООО "БРОНИРОВАНИЕ ГОСТИНИЦ" (OSTROVOK)': ['TravelLine OTK'],
    'ООО "КОМПАНИЯ БРОНЕВИК"': ['TravelLine BR2'],
    '«ЖилиБыли»': ['TravelLine JLB'],
}


USADBA_AGENTS = [
    ("ООО \"ХБ ПРО\"", 20, ["хб-про", "хб про", "hb-pro", "hbpro"]),
    ("ООО Компания Броневик", 20, ["броневик", "bronevik"]),
    ("Вайт Тревел (One Two Trip)", 15, ["one two trip", "onetwotrip", "вайт тревел", "white travel", "12t"]),
    ("ООО Бронирование гостиниц (Ostrovok)", 18, ["островок", "ostrovok"]),
    ("ИП Лагун (CUVA)", 10, ["cuva", "кува", "лагун", "lagun"]),
    ("ООО Яндекс Путешествия", 18, ["яндекс", "yandex"]),
    ("Эй энд Эй (Zabroniryi)", 18, ["забронируй", "zabroniryi", "эй энд эй", "эйэндэй"]),
    ("Академ - он лайн", 20, ["академ", "akadem", "аон.ру", "aon.ru"]),
    ("ООО Новые туристические технологии (TUTU.RU)", 15, ["tutu", "туту", "новые туристические технологии"]),
    ("РусАэро", 10, ["русаэро", "rusaero", "рус аэро"]),
    ("Свена Тур", 20, ["свена", "svena"]),
    ("Свой ТТ (Hotelbeds)", 20, ["hotelbeds", "хотелбедс", "свой тт"]),
    ("Trip.com", 15, ["trip.com", "tripcom"]),
    ("ГЦБиТ", 20, ["гцбит", "городской центр бронирования"]),
    ("\"Янтарный край\"", 20, ["янтарный край"]),
    ("\"Онлайн Инновации\" (all hotels, 101 hotel)", 17, ["all hotels", "101 hotel", "онлайн инновации"]),
    ("Юнона", 15, ["юнона", "yunona"]),
    ("Cbooking", 20, ["cbooking", "сибукинг"]),
    ("Атлас-Тур", 20, ["атлас-тур", "атлас тур", "atlas"]),
    ("Ноктюрн", 20, ["ноктюрн", "nocturne"]),
    ("Калининградские курорты", 20, ["калининградские курорты"]),
    ("Балтма-Тур/Балтма-Вояж", 20, ["балтма", "baltma"]),
    ("РЦБ (Краснодар)", 20, ["рцб", "краснодар"]),
    ("otelelo", 17, ["otelelo", "отелело", "otello"]),   # рукописная правка 15→17
    ("ООО \"ОТИС\" (seven travel)", 15, ["seven travel", "севен трэвел", "отис", "otis"]),
    ("SUTOCHNO.RU", 17, ["суточно", "sutochno"]),        # рукописная правка 15→17
    ("Русс Пасс", 4.5, ["русс пасс", "russ pass"]),
    ("Янтарная мозаика/Амбер Тур-Н", 15, ["янтарная мозаика", "амбер", "amber"]),
    ("Fun&Sun (ТТ-Трэвел)", 15, ["fun&sun", "funsun", "тт-трэвел"]),
    ("TRIVIO", 15, ["trivio", "тривио"]),
    ("ООО «Интернет тревел» (OZON)", 15, ["ozon", "озон", "интернет тревел"]),
    ("Комфорт Букинг", 1.5, ["комфорт букинг", "komfort booking"]),
    ("ООО ТУРМАШИНА", 15, ["турмашина", "tourmachine"]),
    ("ТУРНЕ-ТРАНС БИЗНЕС", 20, ["турне-транс", "turne-trans"]),
    ("Библиоглобус", 15, ["библиоглобус", "biblio"]),
    ("Сусанин", 15, ["сусанин", "susannin"]),
    ("Т-Путешествия (Тинькофф)", 15, ["тинькофф", "tinkoff", "т-путешествия"]),
    ("ООО Туроператор Дельфин", 0, ["дельфин", "delfin"]),  # СКИДКА -> 0%
    ("Компания Отели Онлайн (Авиацентр)", 15, ["авиацентр", "avia-centre", "отели онлайн"]),
    ("ООО \"Свой ТС\"", 15, ["свой тс", "svoy hotel"]),
    ("Жили-Были", 15, ["жили-были", "жилибыли"]),
    ("ZSVA/Абразив Эксперт", 10, ["абразив", "abraziv", "zsva", "эсва"]),
    ("ИП Стефание Е.А. (Актер)", 20, ["актер", "артисты", "стефание"]),
    ("Авиапартнер", 20, ["авиапартнер", "aviapartner"]),
    ("Фабрика обоев Прима Итальяна", 20, ["прима итальяна", "prima"]),
    ("Верона-Мебель", 10, ["верона", "verona"]),
    ("Казино Sobranie", 10, ["sobranie", "собрание", "юни гейминг"]),
    ("ООО \"СС-ОРМА\" (Сформа)", 10, ["сформа", "s-forma", "сс-орма"]),
    ("АО КРИОГАЗ", 10, ["криогаз", "kriogaz"]),
    ("ООО БЛК", 10, ["блк"]),
    ("Кат ТРЭВЕЛ", 15, ["кат трэвел", "kat travel", "киносамбия"]),  # в таблице 10–20%
    ("Фабрика \"ЯНТА\"", 12.5, ["янта", "yanta"]),                   # в таблице 10–15%
    ("ООО Система Алеан", 0, ["алеан", "alean"]),  # СКИДКА -> 0%
    ("PEGAS", 0, ["pegas", "пегас"]),  # СКИДКА -> 0%
    ("Универсалтур", 15, ["универсалтур", "universaltur"]),
    ("Авентураманяия", 15, ["авентураманяия", "авентура"]),
    ("Ростелеком (СБО)", 15, ["ростелеком", "сбо"]),
    ("Мультитур", 15, ["мультитур", "multitur"]),
    ("Интурист", 15, ["интурист", "intourist"]),
]


def _tch_rates(spec):
    """spec: [('base'|'до дд.мм.гггг'|'с дд.мм.гггг'|'dd.mm.yyyy-dd.mm.yyyy', pct)...]"""
    import re as _re
    rates = []
    for label, pct in spec:
        if label == 'base':
            rates.append({'from': '', 'to': '', 'pct': pct})
        elif label.startswith('до '):
            rates.append({'from': '', 'to': label[3:].strip(), 'pct': pct})
        elif label.startswith('с '):
            rates.append({'from': label[2:].strip(), 'to': '', 'pct': pct})
        else:
            f, t = label.split('-')
            rates.append({'from': f.strip(), 'to': t.strip(), 'pct': pct})
    return rates
