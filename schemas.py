# -*- coding: utf-8 -*-
"""Схемы входных файлов: поля приложения + синонимы для автоподбора колонок.

key      — внутреннее имя поля
label    — название в диалоге сопоставления
required — обязательное поле (иначе обработка невозможна)
aliases  — варианты названий колонок в выгрузках (для автоподбора)
"""

TCH_DATA_SCHEMA = [
    dict(key='guest',    label='Гость',            required=True,
         aliases=['гость', 'гости', 'фио', 'фио гостя', 'имя', 'проживающий']),
    dict(key='price',    label='Стоимость',        required=True,
         aliases=['стоимость', 'цена', 'сумма', 'стоимость проживания', 'итого', 'цена за номер']),
    dict(key='checkin',  label='Заезд',            required=True,
         aliases=['заезд', 'прибытие', 'дата заезда', 'заселение', 'начало']),
    dict(key='checkout', label='Выезд',            required=True,
         aliases=['выезд', 'убытие', 'дата выезда', 'выселение', 'конец']),
    dict(key='source',   label='Источник',         required=True,
         aliases=['источник', 'источник бронирования', 'канал', 'агент', 'площадка', 'откуда']),
    dict(key='category', label='Категория номера', required=True,
         aliases=['категория номера', 'категория', 'тип номера', 'номер категория']),
    dict(key='bron',     label='№ брони',          required=False,
         aliases=['№ брони', 'номер брони', 'бронь', 'идентификатор', 'id брони', 'id']),
    dict(key='nights',   label='Ночей',            required=False,
         aliases=['ночей', 'ночи', 'кол-во ночей', 'количество ночей', 'суток']),
    dict(key='room',     label='№ комнаты',        required=False,
         aliases=['№ комнаты', 'номер комнаты', 'комната']),
    dict(key='corpus',   label='Корпус',           required=False,
         aliases=['корпус', 'здание', 'building']),
]

TCH_ROOMS_SCHEMA = [
    dict(key='fio',  label='ФИО',       required=True,
         aliases=['фио', 'гость', 'имя', 'проживающий']),
    dict(key='room', label='№ комнаты', required=True,
         aliases=['№ комнаты', 'номер комнаты', 'комната', 'номер', 'room']),
    dict(key='corpus', label='Корпус', required=False,
         aliases=['корпус', 'здание', 'building']),
]

# Дона: те же поля, но «Категория номера» необязательна
DNA_DATA_SCHEMA = [
    dict(f, required=(f['required'] and f['key'] != 'category'))
    for f in TCH_DATA_SCHEMA
]

# Поля без обязательных (для отображения «*» в диалоге)
def required_keys(schema):
    return [f['key'] for f in schema if f['required']]

# Слоты с настраиваемым сопоставлением колонок.
# Усадьбы здесь нет: структура её отчёта фиксированная.
SCHEMAS = {('tch', 'data'): TCH_DATA_SCHEMA, ('tch', 'rooms'): TCH_ROOMS_SCHEMA,
           ('dna', 'data'): DNA_DATA_SCHEMA}


def needs_mapping(code, slot):
    """Нужен ли диалог сопоставления колонок для (отель, слот)."""
    return (code, slot) in SCHEMAS
