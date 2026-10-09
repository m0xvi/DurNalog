# -*- coding: utf-8 -*-
"""Палитра, шрифты и демо-данные GUI. Расчёта здесь нет."""

from __future__ import annotations

import customtkinter as ctk

COLORS = {
    "bg": "#E7EDF5",
    "sidebar": "#1B2A4A",
    "sidebar_hover": "#243656",
    "sidebar_act": "#31476A",
    "sidebar_text": "#C5D3E8",
    "sidebar_muted": "#7E93B0",
    "card": "#FFFFFF",
    "text": "#1A2332",
    "muted": "#5C6B80",
    "faint": "#8A97A8",
    "line": "#D7E0EA",
    "line2": "#EEF2F6",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "accent_soft": "#E8F0FE",
    "green": "#0B8F55",
    "green_hover": "#097A48",
    "green_soft": "#E5F7EE",
    "green_dis": "#A9CABB",
    "red": "#DC3D43",
    "red_soft": "#FDECEC",
    "amber": "#C47B12",
    "amber_soft": "#FFF4E0",
    "purple": "#6D28D9",
    "purple_soft": "#F3E8FF",
    "purple_text": "#4C1D95",
    "drop": "#F7FAFC",
    "drop_border": "#C5D2E0",
    "status_bg": "#F3F5F8",
}

HOTEL_DOT = {
    "chai": "#2BB3C8",
    "dona": "#E2A03A",
    "usadba": "#D46A8A",
}

HOTEL_PILL = {
    "chai": ("#E0F4FA", "#0E7490"),
    "dona": ("#FFF4E0", "#C47B12"),
    "usadba": ("#FCE8EF", "#B54A6A"),
}

HOTELS = {
    "chai": {
        "id": "chai",
        "name": "Чайковский",
        "title": "Отель «Чайковский»",
        "sub": "Турналог с разбивкой по корпусам",
        "chips": [
            ("#E0F4FA", "#0E7490", "корпуса"),
            ("#EEF2F6", "#5C6B80", "листы 1&2, 1, 2"),
            ("#EEF2F6", "#5C6B80", "«нет корпуса»"),
            ("#E8F0FE", "#2563EB", "ставка 2%"),
        ],
        "needs_rooms": True,
        "last": "08.09.2026",
        "last_sub": "III кв. · 842 строки",
        "tax": "128 450 ₽",
        "dup": "4 · 11",
        "demo_data": "выгрузка_чайковский_авг.xlsx",
        "demo_rooms": "номера_корпуса_авг.xlsx",
        "rate": "2.0",
        "min_night": "100.0",
        "threshold": "5100.0",
        "tax_in_price": True,
        "auto_threshold": True,
        "tax_no_commission": True,
        "formula_daily": "={P}/{N}",
        "formula_method": "=IF({R}>={T},\"ставка\",\"мин.ставка\")",
        "formula_tax": "=ROUND(IF({M}=\"мин.ставка\",{N}*{min},{P}/{div}*{rate}),2)",
    },
    "dona": {
        "id": "dona",
        "name": "Дона",
        "title": "Отель «Дона»",
        "sub": "Турналог без разбивки по корпусам",
        "chips": [
            ("#EEF2F6", "#5C6B80", "один объект"),
            ("#E8F0FE", "#2563EB", "ставка 2%"),
        ],
        "needs_rooms": False,
        "last": "04.07.2026",
        "last_sub": "II кв. · 516 строк",
        "tax": "74 210 ₽",
        "dup": "0 · 6",
        "demo_data": "dona_q3_брони.xlsx",
        "demo_rooms": None,
        "rate": "2.0",
        "min_night": "100.0",
        "threshold": "5100.0",
        "tax_in_price": True,
        "auto_threshold": True,
        "tax_no_commission": True,
        "formula_daily": "={P}/{N}",
        "formula_method": "=IF({R}>={T},\"ставка\",\"мин.ставка\")",
        "formula_tax": "=ROUND(IF({M}=\"мин.ставка\",{N}*{min},{P}/{div}*{rate}),2)",
    },
    "usadba": {
        "id": "usadba",
        "name": "Усадьба",
        "title": "Отель «Усадьба»",
        "sub": "Турналог без разбивки по корпусам",
        "chips": [
            ("#EEF2F6", "#5C6B80", "один объект"),
            ("#E8F0FE", "#2563EB", "ставка 2%"),
        ],
        "needs_rooms": False,
        "last": "03.07.2026",
        "last_sub": "II кв. · 291 строка",
        "tax": "41 880 ₽",
        "dup": "2 · 3",
        "demo_data": "usadba_лето.xls",
        "demo_rooms": None,
        "rate": "2.0",
        "min_night": "100.0",
        "threshold": "5100.0",
        "tax_in_price": True,
        "auto_threshold": True,
        "tax_no_commission": True,
        "formula_daily": "={P}/{N}",
        "formula_method": "=IF({R}>={T},\"ставка\",\"мин.ставка\")",
        "formula_tax": "=ROUND(IF({M}=\"мин.ставка\",{N}*{min},{P}/{div}*{rate}),2)",
    },
}

FORMULAS = {
    "daily": "={P}/{N}",
    "method": '=IF({R}>={T},"ставка","мин.ставка")',
    "tax": '=ROUND(IF({M}="мин.ставка",{N}*{min},{P}/{div}*{rate}),2)',
    "placeholders": (
        "Плейсхолдеры: {P} цена · {N} ночей · {R} сут. ставка · {M} метод · "
        "{T} порог · {min} минимум · {rate} ставка · {div} делитель"
    ),
}

AGENTS = [
    "Балтма тур",
    "Янтарный край",
    "УК ПЦБИТ",
    "Fun&Sun TT Трэвел",
    "Броневик",
    "Янт.Мозаика",
    "Анюта",
    "Академсервис",
    "ХБ Про (ХБ Трэвел)",
    "Сибукинг",
    "101 Отель ОНЛАЙН Инновации",
    "Алеан Отдых",
    "Вайт Тревел",
    "Калининградские курорты",
    "Бюро путешествий",
    "Seven Travel (OTIS)",
    "Островок",
    "Суточно.ру",
    "Эй энд Эй (ЦБП)",
    "Юнона",
    "Expedia",
]

# Демо-история. В боевом коде читать sqlite.
HISTORY_ROWS = [
    {
        "id": "3",
        "hotel_id": "chai",
        "hotel": "Чайковский",
        "quarter": "III 2026",
        "processed": "08.09.2026 14:22",
        "rows": "842",
        "dups": "4",
        "orig": "выгрузка_август.xlsx",
        "result": "Турналог_Чайковский_2026_III.xlsx",
        "cross": [
            "Петрова А. · 28.06–02.07.2026 · II→III",
            "Kim J. · 30.06–03.07.2026 · II→III",
        ],
    },
    {
        "id": "2",
        "hotel_id": "dona",
        "hotel": "Дона",
        "quarter": "II 2026",
        "processed": "04.07.2026 11:03",
        "rows": "516",
        "dups": "0",
        "orig": "dona_q2.xlsx",
        "result": "Турналог_Дона_2026_II.xlsx",
        "cross": [],
    },
    {
        "id": "1",
        "hotel_id": "usadba",
        "hotel": "Усадьба",
        "quarter": "II 2026",
        "processed": "03.07.2026 16:48",
        "rows": "291",
        "dups": "2",
        "orig": "usadba_июнь.xls",
        "result": "Турналог_Усадьба_2026_II.xlsx",
        "cross": ["Сидоров · 29.03–02.04.2026 · I→II"],
    },
]

USERS = [
    ("admin", "администратор", "—"),
    ("chai_buh", "отель", "Чайковский"),
    ("dona_buh", "отель", "Дона"),
]

DB_PATH_DEMO = r"F:\Desktop\TurNalog\turnalog_history.db"

COLUMN_MAP_DATA = [
    ("Дата заезда", "C — Arrival"),
    ("Дата выезда", "D — Departure"),
    ("Ночей", "E"),
    ("Сумма", "K — Total"),
    ("Источник / агент", "B — Agency"),
    ("ФИО гостя", "A — Guest"),
    ("Номер брони", "F — RecNo"),
]

COLUMN_MAP_ROOMS = [
    ("ФИО", "A"),
    ("Номер комнаты", "C"),
    ("Корпус", "D"),
    ("Дата", "B"),
]

FONTS: dict[str, ctk.CTkFont] = {}
FONT_FAMILY = "Segoe UI"


def setup_fonts() -> str:
    """Вызвать после создания CTk(). На Windows будет Segoe UI."""
    global FONT_FAMILY
    import tkinter.font as tkfont

    fams = set(tkfont.families())
    for cand in ("Segoe UI", "PT Sans", "Noto Sans", "DejaVu Sans", "Liberation Sans"):
        if cand in fams:
            FONT_FAMILY = cand
            break
    else:
        FONT_FAMILY = "TkDefaultFont"

    FONTS.update(
        {
            "ui": ctk.CTkFont(family=FONT_FAMILY, size=13),
            "ui_b": ctk.CTkFont(family=FONT_FAMILY, size=13, weight="bold"),
            "sm": ctk.CTkFont(family=FONT_FAMILY, size=12),
            "sm_b": ctk.CTkFont(family=FONT_FAMILY, size=12, weight="bold"),
            "xs": ctk.CTkFont(family=FONT_FAMILY, size=11),
            "xs_b": ctk.CTkFont(family=FONT_FAMILY, size=11, weight="bold"),
            "h1": ctk.CTkFont(family=FONT_FAMILY, size=22, weight="bold"),
            "h2": ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            "h3": ctk.CTkFont(family=FONT_FAMILY, size=14, weight="bold"),
            "stat": ctk.CTkFont(family=FONT_FAMILY, size=20, weight="bold"),
            "brand": ctk.CTkFont(family=FONT_FAMILY, size=16, weight="bold"),
            "mono": ctk.CTkFont(family="Consolas", size=12),
        }
    )
    return FONT_FAMILY
