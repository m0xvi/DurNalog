# -*- coding: utf-8 -*-
"""Параметры и формулы турналога, общие для всех вкладок.

НК РФ, гл. 33.1: ставка растёт по годам (2025: 1%, 2026: 2%, 2027: 3% ...),
минимум — 100 руб/сутки. Налог считается включённым в стоимость проживания.
При смене года/ставки правятся только константы ниже.
"""
from datetime import date, datetime

RATE_PCT = 2                    # ставка налога, %
DIVISOR = 100 + RATE_PCT        # 102 (налог «в том числе»)
MIN_PER_NIGHT = 100             # минимальный налог, руб/сутки
# Суточная ставка, при которой процентный налог достигает минимума:
# 5100 * 2 / 102 = 100. Старое значение 5000 давало налог ниже минимума
# в диапазоне 5000–5099,99 руб/ночь.
SWITCH_THRESHOLD = 5100


def tax_formula_daily_rate(col_price, col_nights, r):
    """Сут.ставка = Цена / Ночей"""
    return f'={col_price}{r}/{col_nights}{r}'


def tax_formula_method(col_rate, r):
    return f'=IF({col_rate}{r}>={SWITCH_THRESHOLD},"ставка","мин.ставка")'


def tax_formula_tax(col_method, col_nights, col_price, r):
    return (f'=ROUND(IF({col_method}{r}="мин.ставка",'
            f'{col_nights}{r}*{MIN_PER_NIGHT},{col_price}{r}/{DIVISOR}*{RATE_PCT}),2)')


def parse_price(v):
    """Приведение стоимости к числу: числа проходят как есть, строки вида
    '5 000,00' / '5000.00' / '5000 руб.' очищаются. Не распознано -> None."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        return f if f == f else None  # NaN -> None
    s = str(v).strip()
    if not s:
        return None
    s = s.replace('\xa0', '').replace(' ', '').replace('руб.', '').replace('руб', '')
    s = s.replace(',', '.')
    s = ''.join(ch for ch in s if ch in '0123456789.-')
    try:
        return float(s)
    except ValueError:
        return None


def as_date(v):
    """datetime/date/строка 'дд.мм.гггг [чч:мм]' -> date или None."""
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if v is None:
        return None
    s = str(v).strip()
    for fmt in ('%d.%m.%Y %H:%M', '%d.%m.%Y %H:%M:%S', '%d.%m.%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def dates_overlap(s1, e1, s2, e2):
    """Строгое пересечение интервалов дат (смежные ночи не считаются)."""
    if None in (s1, e1, s2, e2):
        return False
    return s1 < e2 and s2 < e1


def quarter_of(d):
    """date -> '2026-Q2'."""
    return f'{d.year}-Q{(d.month - 1) // 3 + 1}'


# ================================================================
# НАСТРАИВАЕМЫЕ ПАРАМЕТРЫ НАЛОГА (окно «Настройки» в приложении)
# ================================================================
DEFAULT_TAX = {
    'rate': 2,            # ставка налога, %
    'min_per_night': 100, # минимальный налог, руб/сутки
    'threshold': 5100,    # порог суточной ставки для метода «ставка»
    'incl': True,         # налог включён в стоимость (делитель 100+rate)
    'use_full': True,     # налог от стоимости БЕЗ комиссии (пересчитанной в 100%)
    'auto_threshold': True,  # порог считается автоматически
    # Шаблоны Excel-формул. Плейсхолдеры: {P} цена, {N} ночей, {R} сут.ставка,
    # {M} метод, {T} порог, {min} минимум, {rate} ставка, {div} делитель.
    'tpl_daily': '={P}/{N}',
    'tpl_method': '=IF({R}>={T},"ставка","мин.ставка")',
    'tpl_tax': '=ROUND(IF({M}="мин.ставка",{N}*{min},{P}/{div}*{rate}),2)',
}


def _num(v):
    f = float(v)
    return str(int(f)) if f == int(f) else str(f)


class TaxParams:
    """Параметры налога отеля + генерация Excel-формул из шаблонов."""

    def __init__(self, d=None):
        d = dict(DEFAULT_TAX, **(d or {}))
        self.rate = float(d['rate'])
        self.min = float(d['min_per_night'])
        self.incl = bool(d['incl'])
        self.auto = bool(d['auto_threshold'])
        self.use_full = bool(d.get('use_full', True))
        self.threshold = (self.auto_threshold() if self.auto
                          else float(d['threshold']))
        self.tpl_daily = d['tpl_daily']
        self.tpl_method = d['tpl_method']
        self.tpl_tax = d['tpl_tax']

    @property
    def divisor(self):
        return 100.0 + self.rate if self.incl else 100.0

    def auto_threshold(self):
        """Суточная ставка, при которой процентный налог достигает минимума."""
        return (self.min * (100.0 + self.rate) / self.rate if self.incl
                else self.min * 100.0 / self.rate)

    # --- генерация формул для ячейки строки r ---
    def f_daily(self, P, N, r):
        return (self.tpl_daily.replace('{P}', f'{P}{r}').replace('{N}', f'{N}{r}'))

    def f_method(self, R, r):
        return self.tpl_method.replace('{R}', f'{R}{r}').replace('{T}', _num(self.threshold))

    def f_tax(self, M, N, P, r):
        return (self.tpl_tax
                .replace('{M}', f'{M}{r}').replace('{N}', f'{N}{r}')
                .replace('{P}', f'{P}{r}')
                .replace('{min}', _num(self.min))
                .replace('{rate}', _num(self.rate))
                .replace('{div}', _num(self.divisor)))

    def as_dict(self):
        return {'rate': self.rate, 'min_per_night': self.min, 'use_full': self.use_full,
                'threshold': self.threshold, 'incl': self.incl,
                'auto_threshold': self.auto, 'tpl_daily': self.tpl_daily,
                'tpl_method': self.tpl_method, 'tpl_tax': self.tpl_tax}

    def preview(self):
        return (f"Сут.ставка:  {self.f_daily('D', 'F', 2)}\n"
                f"Метод:       {self.f_method('G', 2)}\n"
                f"Налог:       {self.f_tax('H', 'F', 'D', 2)}")
