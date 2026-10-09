# -*- coding: utf-8 -*-
"""Усадьба: фиксированная структура — сопоставление колонок не требуется.

Регрессия на KeyError ('usd', 'data') в предфазе обработки.
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from schemas import SCHEMAS, needs_mapping


def test_usadba_needs_no_mapping():
    assert not needs_mapping('usd', 'data')
    assert not needs_mapping('usd', 'rooms')
    assert ('usd', 'data') not in SCHEMAS
    assert ('usd', 'rooms') not in SCHEMAS


def test_tch_dna_still_mapped():
    assert needs_mapping('tch', 'data')
    assert needs_mapping('tch', 'rooms')
    assert needs_mapping('dna', 'data')
    assert not needs_mapping('dna', 'rooms')  # у Доны нет слота номеров
