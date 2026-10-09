# -*- coding: utf-8 -*-
"""Общая логика диалогов без GUI-зависимостей (можно тестировать без tkinter)."""
import re

NOT_SET = "— не указано —"
NEW_AGENT = "+ новый агент"


def filter_agent_options(query, agent_names, limit=30):
    """Живой поиск для диалога источников: подстрока без учёта регистра.
    Пустой запрос -> полный список (без лимита).
    Возвращает [NOT_SET] + совпадения (не более limit) + [NEW_AGENT]."""
    q = (query or '').strip().lower()
    if not q or q == NOT_SET.lower():
        hits = list(agent_names)
    else:
        hits = [n for n in agent_names if q in n.lower()][:limit]
    return [NOT_SET] + hits + [NEW_AGENT]


def resolve_agent_choice(choice, agent_names):
    """Текст из поля со списком ->
    ('none',) / ('agent', имя) / ('new',) / ('unknown', текст).
    Принимается точное совпадение, совпадение без учёта регистра
    и единственное совпадение по подстроке."""
    c = (choice or '').strip()
    if c in ('', NOT_SET):
        return ('none',)
    if c == NEW_AGENT:
        return ('new',)
    if c in agent_names:
        return ('agent', c)
    ci = next((n for n in agent_names if n.lower() == c.lower()), None)
    if ci is not None:
        return ('agent', ci)
    sub = [n for n in agent_names if c.lower() in n.lower()]
    if len(sub) == 1:
        return ('agent', sub[0])
    return ('unknown', c)


def _looks_number(text):
    return re.fullmatch(r'\d+([.,]\d+)?', (text or '').strip()) is not None


def validate_new_agent(src, name, comm_raw):
    """Проверка полей «+ новый агент».

    Возвращает (ok, имя, комиссия, текст_ошибки).
    ok=True  -> (True, имя, комиссия_float, '')
    ok=False -> (False, '', 0.0, сообщение для messagebox)
    """
    src = (src or '').strip()
    name = (name or '').strip()
    raw = (comm_raw or '').strip()
    if raw.endswith('%'):
        raw = raw[:-1].strip()
    if not name:
        return False, '', 0.0, f'Для «{src}» укажите имя нового агента.'
    try:
        comm = float(raw.replace(',', '.'))
    except ValueError:
        if raw == '' and _looks_number(name):
            return False, '', 0.0, (
                f'«{name}» похоже на комиссию, а не на название.\n'
                f'Введите название агента в поле «Имя агента», '
                f'а {name} — в поле «%».')
        return False, '', 0.0, f'Для «{name}» укажите комиссию числом.'
    if not 0 <= comm <= 100:
        return False, '', 0.0, f'Комиссия «{name}» должна быть от 0 до 100%.'
    return True, name, comm, ''
