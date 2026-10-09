# -*- coding: utf-8 -*-
"""Валидация «+ новый агент» в диалоге нераспознанных источников.

Регрессия: пользователь вписал комиссию в поле имени (поле % было за
краем нерастягиваемого окна) и получил «Для "17" укажите комиссию числом».
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from dialogs_common import validate_new_agent as v


def test_ok_plain():
    assert v('src', 'ООО Ромашка', '17') == (True, 'ООО Ромашка', 17.0, '')


def test_ok_variants():
    assert v('src', '  Агент  ', '17,5')[1:] == ('Агент', 17.5, '')
    assert v('src', 'Агент', '17 %')[2] == 17.0  # знак % прощается
    assert v('src', 'Агент', '0')[2] == 0.0
    assert v('src', 'Агент', '100')[2] == 100.0


def test_empty_name_mentions_source():
    ok, _, _, err = v('ООО "Интернет решения"', '   ', '17')
    assert ok is False and 'ООО "Интернет решения"' in err


def test_name_is_source_no_field_mentioned():
    # диалог передаёт источник как имя; в ошибках нет ссылок на поле имени
    ok, name, comm, _ = v('ООО Ромашка', 'ООО Ромашка', '17')
    assert (ok, name, comm) == (True, 'ООО Ромашка', 17.0)
    ok, _, _, err = v('17', '17', '')
    assert ok is False and 'укажите комиссию числом' in err
    assert 'Имя агента' not in err


def test_garbage_commission():
    ok, _, _, err = v('src', 'Агент', 'много')
    assert ok is False and 'укажите комиссию числом' in err
    ok, _, _, err = v('src', 'Агент', '')
    assert ok is False and 'укажите комиссию числом' in err


def test_commission_range():
    for bad in ('150', '-5', '101'):
        ok, _, _, err = v('src', 'Агент', bad)
        assert ok is False and 'от 0 до 100' in err
