# -*- coding: utf-8 -*-
"""Общая логика диалогов без GUI-зависимостей (можно тестировать без tkinter)."""

NOT_SET = "— не указано —"
NEW_AGENT = "+ новый агент"


def filter_agent_options(query, agent_names):
    """Живой поиск для диалога источников: подстрока без учёта регистра.
    Пустой запрос -> полный список.
    Возвращает [NOT_SET] + совпадения + [NEW_AGENT]."""
    q = (query or '').strip().lower()
    if not q or q == NOT_SET.lower():
        hits = list(agent_names)
    else:
        hits = [n for n in agent_names if q in n.lower()]
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
