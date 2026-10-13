# -*- coding: utf-8 -*-
"""Боевые диалоги: сопоставление колонок и нераспознанные источники."""
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

import ui_theme as ui

from dialogs_common import NOT_SET, NEW_AGENT  # noqa: F401  (реэкспорт)

class MappingDialog(ctk.CTkToplevel):
    """Выбор соответствия «поле приложения -> колонка файла» + доп. колонки."""

    def __init__(self, master, title, headers, schema, mapping=None, extra=None):
        super().__init__(master)
        self.title(title)
        self.geometry("640x600")
        self.resizable(False, True)
        self.transient(master)
        self.headers = [str(h) for h in headers]
        self.schema = schema
        self.result = None

        ctk.CTkLabel(self, text="Сопоставление колонок",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(12, 2))
        ctk.CTkLabel(self, text=f"{title}\n"
                                "Укажите, какой колонке файла соответствует каждое поле.\n"
                                "«*» — обязательные поля. Ненужные колонки просто не выбирайте.",
                     text_color="gray", justify="left").pack(padx=20, pady=(0, 8))

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=15, pady=5)

        scroll = ctk.CTkScrollableFrame(body, fg_color=("gray88", "gray17"))
        scroll.pack(fill="both", expand=True)

        self.vars = {}
        for f in schema:
            row = ctk.CTkFrame(scroll, fg_color="transparent")
            row.pack(fill="x", pady=3, padx=6)
            lbl = f['label'] + ("  *" if f['required'] else "")
            ctk.CTkLabel(row, text=lbl, width=170, anchor="w",
                         font=ctk.CTkFont(size=13)).pack(side="left")
            cur = (mapping or {}).get(f['key'])
            if cur not in self.headers:
                cur = None
            var = tk.StringVar(value=cur or NOT_SET)
            self.vars[f['key']] = var
            ctk.CTkOptionMenu(row, variable=var, width=300,
                              values=[NOT_SET] + self.headers).pack(side="left", padx=6)

        ctk.CTkLabel(self, text="Дополнительные колонки файла — перенести в результат:",
                     anchor="w", font=ctk.CTkFont(size=13, weight="bold")).pack(padx=22, pady=(8, 2), fill="x")
        ex_wrap = ctk.CTkFrame(self, fg_color="transparent")
        ex_wrap.pack(fill="x", padx=20)
        self.extra_vars = {}
        extra = set(extra or [])
        for i, h in enumerate(self.headers):
            v = tk.BooleanVar(value=h in extra)
            self.extra_vars[h] = v
            cb = ctk.CTkCheckBox(ex_wrap, text=h, variable=v,
                                 checkbox_width=16, checkbox_height=16,
                                 font=ctk.CTkFont(size=12))
            cb.grid(row=i // 3, column=i % 3, sticky="w", padx=6, pady=2)
        ex_wrap.columnconfigure((0, 1, 2), weight=1)

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(pady=12)
        ctk.CTkButton(bar, text="Автоподбор", width=110, fg_color="gray",
                      command=self.auto).pack(side="left", padx=4)
        ctk.CTkButton(bar, text="OK", width=110, fg_color="green",
                      command=self.ok).pack(side="left", padx=4)
        ctk.CTkButton(bar, text="Отмена", width=110, fg_color="gray",
                      command=self.destroy).pack(side="left", padx=4)
        self.grab_set()
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)

    def auto(self):
        from core import guess_mapping  # локально: core ничего не знает о диалогах
        g = guess_mapping(self.headers, self.schema)
        for key, var in self.vars.items():
            var.set(g.get(key) or NOT_SET)

    def ok(self):
        mapping = {k: (v.get() if v.get() != NOT_SET else None)
                   for k, v in self.vars.items()}
        missing = [f['label'] for f in self.schema
                   if f['required'] and not mapping.get(f['key'])]
        if missing:
            messagebox.showerror("Обязательные поля",
                                 "Не сопоставлены обязательные поля:\n• " +
                                 "\n• ".join(missing), parent=self)
            return
        extra = [h for h, v in self.extra_vars.items() if v.get()]
        self.result = (mapping, extra)
        self.destroy()


class UnrecognizedSourcesDialog(ctk.CTkToplevel):
    """Для каждого нераспознанного источника: сопоставить с агентом,
    создать нового агента (сохранится глобально) или оставить 0%.
    Поле со списком — с живым поиском: начните вводить название."""

    def __init__(self, master, items, agent_names):
        super().__init__(master)
        self.title("Нераспознанные источники")
        self.geometry("940x560")
        self.resizable(True, True)
        self.minsize(720, 420)
        self.transient(master)
        self.result = None
        self.agent_names = list(agent_names)

        ctk.CTkLabel(self, text="Нераспознанные источники",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(12, 2))
        help_text = ("Укажите, к кому отнести источники. Выбор сохранится "
                     "и в следующий раз вопрос не появится.\n«— не указано —» "
                     "= комиссия 0% без подсветки ошибки. "
                     "«+ новый агент»: имя возьмётся из левой колонки — впишите только %.\n"
                     "Начните вводить название и откройте список стрелкой или клавишей ↓ — "
                     "он уже отфильтрован.")
        if any(str(src).strip().casefold() == 'travelline cmb'
               for src, _ in items):
            help_text += (
                "\nTravelLine CMB — Comfort Booking: 2% относится к агрегатору; "
                "сверху может начисляться комиссия подканала. Полную ставку по "
                "одному коду определить нельзя — уточните подканал/договор.")
        ctk.CTkLabel(self, text=help_text, text_color="gray",
                     wraplength=900).pack(padx=20)

        sc = ctk.CTkScrollableFrame(self, fg_color=("gray88", "gray17"))
        sc.pack(fill="both", expand=True, padx=15, pady=8)

        self.rows = {}
        for src, cnt in items:
            fr = ctk.CTkFrame(sc, fg_color="transparent")
            fr.pack(fill="x", pady=3, padx=6)
            fr.grid_columnconfigure(0, weight=1)
            ctk.CTkLabel(fr, text=f"{src}   (строк: {cnt})", anchor="w",
                         justify="left",
                         font=ctk.CTkFont(size=12)).grid(row=0, column=0, sticky="ew")
            var = tk.StringVar(value=NOT_SET)
            combo = ctk.CTkComboBox(fr, variable=var, width=230,
                                    values=self._full_opts(),
                                    command=lambda v, s=src: self._picked(s))
            combo.grid(row=0, column=1, padx=6, sticky="e")
            # ввод идёт во внутреннее Entry (есть в 5.x и 6.x) — биндимся на него,
            # т.к. события клавиш не всплывают до фрейма комбобокса
            entry = getattr(combo, "_entry", combo)
            entry.bind("<KeyRelease>", lambda e, s=src: self._on_type(s, e))
            new_fr = ctk.CTkFrame(fr, fg_color="transparent")
            comm_var = tk.StringVar()
            comm_entry = ctk.CTkEntry(new_fr, textvariable=comm_var, width=60, height=26,
                                      placeholder_text="%",
                                      font=ctk.CTkFont(size=11))
            comm_entry.pack(side="left", padx=(6, 0))
            self.rows[src] = dict(var=var, combo=combo, new_fr=new_fr,
                                  comm_var=comm_var, comm_entry=comm_entry)

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(pady=(0, 12))
        ctk.CTkButton(bar, text="OK", width=120, fg_color="green",
                      command=self.ok).pack(side="left", padx=6)
        ctk.CTkButton(bar, text="Отмена", width=120, fg_color="gray",
                      command=self.destroy).pack(side="left", padx=6)
        self.grab_set()
        self.lift()
        self.focus_force()
        self.attributes("-topmost", True)

    def _full_opts(self):
        return [NOT_SET] + self.agent_names + [NEW_AGENT]

    def _picked(self, src):
        """Выбор пункта из списка: вернуть полный список и показать/скрыть
        поля нового агента."""
        self.rows[src]['combo'].configure(values=self._full_opts())
        self._toggle(src)
        if self.rows[src].get('new_shown'):
            try:
                self.rows[src]['comm_entry'].focus_set()
            except Exception:
                pass

    def _on_type(self, src, event):
        """Ввод в поле: ↓ открывает отфильтрованный список, остальные
        печатаемые клавиши — обновляют фильтр с задержкой (debounce).
        Меню НЕ переоткрывается при вводе: на Windows открытое меню
        перехватывает клавиатуру, и поле «виснет» после первого символа."""
        if event.keysym == 'Down':
            self._cancel_filter(src)
            self._open_menu(src)
            return
        if event.keysym in ('Up', 'Left', 'Right', 'Escape', 'Return',
                            'Tab', 'Shift_L', 'Shift_R', 'Control_L', 'Control_R',
                            'Alt_L', 'Alt_R', 'Caps_Lock', 'Num_Lock',
                            'Home', 'End', 'Prior', 'Next'):
            return
        r = self.rows[src]
        self._cancel_filter(src)
        try:
            r['filter_job'] = self.after(120, lambda: self._apply_filter(src))
        except Exception:
            pass
        self._toggle(src)

    def _cancel_filter(self, src):
        r = self.rows.get(src) or {}
        job = r.pop('filter_job', None)
        if job:
            try:
                self.after_cancel(job)
            except Exception:
                pass

    def _apply_filter(self, src):
        """Применить фильтр к списку, не трогая фокус и не открывая меню."""
        r = self.rows.get(src)
        if r is None:
            return
        r.pop('filter_job', None)
        try:
            from dialogs_common import filter_agent_options
            if not self.winfo_exists():
                return
            r['combo'].configure(
                values=filter_agent_options(r['var'].get(), self.agent_names))
        except Exception:
            pass

    def _open_menu(self, src):
        """Раскрыть список с текущим фильтром (для клавиши ↓)."""
        r = self.rows.get(src)
        if r is None:
            return
        try:
            from dialogs_common import filter_agent_options
            r['combo'].configure(
                values=filter_agent_options(r['var'].get(), self.agent_names))
            opener = getattr(r['combo'], '_open_dropdown_menu', None)
            if opener is not None:
                opener()  # приватный API, одинаковый в 5.x и 6.x
        except Exception:
            pass

    def _toggle(self, src):
        r = self.rows[src]
        show = r['var'].get() == NEW_AGENT
        if show == r.get('new_shown', False):
            return  # уже в нужном состоянии — не дёргаем layout
        r['new_shown'] = show
        if show:
            r['new_fr'].grid(row=0, column=2, sticky="e")
        else:
            r['new_fr'].grid_remove()

    def ok(self):
        from dialogs_common import resolve_agent_choice, validate_new_agent
        result = {}
        for src, r in self.rows.items():
            res = resolve_agent_choice(r['var'].get(), self.agent_names)
            if res[0] == 'none':
                result[src] = ('none',)
            elif res[0] == 'new':
                good, name, comm, err = validate_new_agent(
                    src, src, r['comm_var'].get())
                if not good:
                    messagebox.showerror("Новый агент", err, parent=self)
                    return
                result[src] = ('new', name, comm)
            elif res[0] == 'agent':
                result[src] = ('agent', res[1])
            else:
                messagebox.showerror(
                    "Нераспознанные источники",
                    f"«{res[1]}» нет в списке агентов.\n"
                    f"Выберите агента из списка или «{NEW_AGENT}».",
                    parent=self)
                return
        self.result = result
        self.destroy()


class TlConfigDialog(ctk.CTkToplevel):
    """Подключение TravelLine: client_id / client_secret / propertyId."""

    def __init__(self, master, preset=None, hotel_name=''):
        super().__init__(master)
        self.title('TravelLine — подключение')
        self.geometry('420x330')
        self.resizable(False, False)
        self.transient(master)
        self.result = None
        preset = preset or {}
        title = f'TravelLine — {hotel_name}' if hotel_name else 'TravelLine'
        ctk.CTkLabel(self, text=title,
                     font=ctk.CTkFont(size=16, weight='bold')).pack(pady=(14, 2))
        ctk.CTkLabel(
            self,
            text='Данные из кабинета TravelLine: Настройки гостиницы →\n'
                 'Подключения API. Сохраняются на этом компьютере.',
            text_color='gray').pack(pady=(0, 10))
        self.vars = {}
        for key, label, secret in (
                ('client_id', 'client_id (логин)', False),
                ('client_secret', 'client_secret (пароль)', True),
                ('property_id', 'propertyId (цифры)', False)):
            ctk.CTkLabel(self, text=label, anchor='w').pack(anchor='w', padx=24)
            var = tk.StringVar(value=preset.get(key, ''))
            ctk.CTkEntry(self, textvariable=var, width=370,
                         show='•' if secret else '').pack(padx=24, pady=(0, 8))
            self.vars[key] = var
        bar = ctk.CTkFrame(self, fg_color='transparent')
        bar.pack(pady=6)
        ctk.CTkButton(bar, text='OK', width=120, fg_color='green',
                      command=self.ok).pack(side='left', padx=6)
        ctk.CTkButton(bar, text='Отмена', width=120, fg_color='gray',
                      command=self.destroy).pack(side='left', padx=6)
        self.grab_set()
        self.lift()
        self.focus_force()

    def ok(self):
        res = {k: v.get().strip() for k, v in self.vars.items()}
        if not (res['client_id'] and res['client_secret'] and res['property_id']):
            messagebox.showerror('TravelLine', 'Заполните все три поля.',
                                 parent=self)
            return
        if not res['property_id'].isdigit():
            messagebox.showerror('TravelLine', 'propertyId — только цифры.',
                                 parent=self)
            return
        self.result = res
        self.destroy()


class TlCacheDialog(ctk.CTkToplevel):
    """Выбор: использовать локальный кэш, обновить его, удалить или отменить."""

    def __init__(self, master, quarter_label, cached_at, row_count):
        super().__init__(master)
        self.title('TravelLine — сохранённая выгрузка')
        self.geometry('540x260')
        self.resizable(False, False)
        self.transient(master)
        self.result = None
        self.protocol('WM_DELETE_WINDOW', self.cancel)
        ctk.CTkLabel(self, text='Для периода уже есть локальный кэш',
                     font=ctk.CTkFont(size=15, weight='bold')).pack(pady=(16, 5))
        ctk.CTkLabel(
            self,
            text=(f'{quarter_label}: {row_count} строк, сохранён {cached_at}.\n'
                  'Повторное использование не обращается к TravelLine.\n'
                  'Обновление заново скачает данные и заменит кэш.\n'
                  'В кэше только подготовленные строки (ФИО, № брони, даты и суммы), '
                  'он хранится в базе приложения turnalog_history.db.'),
            text_color='gray', justify='center', wraplength=490).pack(padx=20, pady=5)
        grid = ctk.CTkFrame(self, fg_color='transparent')
        grid.pack(pady=(8, 14))
        ctk.CTkButton(grid, text='Использовать кэш', width=190,
                      fg_color='green', command=lambda: self.finish('use')) \
            .grid(row=0, column=0, padx=5, pady=4)
        ctk.CTkButton(grid, text='Обновить из TravelLine', width=190,
                      command=lambda: self.finish('refresh')) \
            .grid(row=0, column=1, padx=5, pady=4)
        ctk.CTkButton(grid, text='Удалить кэш', width=190,
                      fg_color='#9b3333', command=lambda: self.finish('delete')) \
            .grid(row=1, column=0, padx=5, pady=4)
        ctk.CTkButton(grid, text='Отмена', width=190, fg_color='gray',
                      command=self.cancel).grid(row=1, column=1, padx=5, pady=4)
        self.grab_set()
        self.lift()
        self.focus_force()

    def finish(self, result):
        self.result = result
        self.destroy()

    def cancel(self):
        self.result = 'cancel'
        self.destroy()


class TlProgressDialog(ctk.CTkToplevel):
    """Модальный прогресс загрузки броней TravelLine с кнопкой «Отмена».

    cancel_event — threading.Event: кнопка/крестик устанавливают его,
    fetch_details между запросами бросает TLCancelled.
    Все методы вызывать из главного потока.
    """

    def __init__(self, master, cancel_event, title='TravelLine — загрузка'):
        super().__init__(master)
        self.title(title)
        self.geometry('440x210')
        self.resizable(False, False)
        self.transient(master)
        self.cancel_event = cancel_event
        self.protocol('WM_DELETE_WINDOW', self.cancel)
        ctk.CTkLabel(self, text='Загрузка броней из TravelLine…',
                     font=ctk.CTkFont(size=14, weight='bold')).pack(pady=(16, 4))
        self.stage_lbl = ctk.CTkLabel(self, text='Подключение…',
                                      text_color='gray')
        self.stage_lbl.pack(pady=(0, 8))
        self.bar = ctk.CTkProgressBar(self, width=380, height=14,
                                      mode='indeterminate')
        self.bar.pack(pady=(0, 4))
        self.bar.start()
        self.count_lbl = ctk.CTkLabel(self, text='', font=ctk.CTkFont(size=12))
        self.count_lbl.pack(pady=(0, 10))
        self.cancel_btn = ctk.CTkButton(self, text='Отмена', width=140,
                                        fg_color='gray', command=self.cancel)
        self.cancel_btn.pack()
        self.grab_set()
        self.lift()
        self.focus_force()

    def _alive(self):
        try:
            return bool(self.winfo_exists())
        except Exception:
            return False

    def set_stage(self, text):
        if self._alive():
            self.stage_lbl.configure(text=text)

    def set_determinate(self):
        if self._alive():
            try:
                self.bar.stop()
                self.bar.configure(mode='determinate')
                self.bar.set(0)
            except Exception:
                pass

    def set_progress(self, done, total):
        if not self._alive():
            return
        try:
            if total:
                self.bar.set(min(1.0, max(0.0, done / total)))
                self.count_lbl.configure(text=f'{done} / {total}')
        except Exception:
            pass

    def cancel(self):
        try:
            self.cancel_event.set()
        except Exception:
            pass
        if self._alive():
            try:
                self.cancel_btn.configure(state='disabled')
                self.stage_lbl.configure(text='Отмена…')
            except Exception:
                pass

    def close(self):
        if self._alive():
            try:
                self.grab_release()
            except Exception:
                pass
            try:
                self.destroy()
            except Exception:
                pass
