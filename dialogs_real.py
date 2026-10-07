# -*- coding: utf-8 -*-
"""Боевые диалоги: сопоставление колонок и нераспознанные источники."""
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

import ui_theme as ui

NOT_SET = "— не указано —"
NEW_AGENT = "+ новый агент"

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
    создать нового агента (сохранится глобально) или оставить 0%."""

    def __init__(self, master, items, agent_names):
        super().__init__(master)
        self.title("Нераспознанные источники")
        self.geometry("720x540")
        self.resizable(False, True)
        self.transient(master)
        self.result = None

        ctk.CTkLabel(self, text="Нераспознанные источники",
                     font=ctk.CTkFont(size=16, weight="bold")).pack(pady=(12, 2))
        ctk.CTkLabel(self, text="Укажите, к кому отнести источники. Выбор сохранится "
                                "и в следующий раз вопрос не появится.\n«— не указано —» "
                                "= комиссия 0% без подсветки ошибки.",
                     text_color="gray").pack(padx=20)

        sc = ctk.CTkScrollableFrame(self, fg_color=("gray88", "gray17"))
        sc.pack(fill="both", expand=True, padx=15, pady=8)

        opts = [NOT_SET] + agent_names + [NEW_AGENT]
        self.rows = {}
        for src, cnt in items:
            fr = ctk.CTkFrame(sc, fg_color="transparent")
            fr.pack(fill="x", pady=3, padx=6)
            ctk.CTkLabel(fr, text=f"{src}   (строк: {cnt})", width=330, anchor="w",
                         justify="left", font=ctk.CTkFont(size=12)).pack(side="left")
            var = tk.StringVar(value=NOT_SET)
            ctk.CTkOptionMenu(fr, variable=var, values=opts, width=230,
                              command=lambda v, s=src: self._toggle(s)).pack(side="left", padx=6)
            new_fr = ctk.CTkFrame(fr, fg_color="transparent")
            name_var = tk.StringVar()
            comm_var = tk.StringVar()
            ctk.CTkEntry(new_fr, textvariable=name_var, width=140, height=26,
                         placeholder_text="Имя агента",
                         font=ctk.CTkFont(size=11)).pack(side="left", padx=(6, 3))
            ctk.CTkEntry(new_fr, textvariable=comm_var, width=50, height=26,
                         placeholder_text="%",
                         font=ctk.CTkFont(size=11)).pack(side="left")
            self.rows[src] = dict(var=var, new_fr=new_fr,
                                  name_var=name_var, comm_var=comm_var)

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

    def _toggle(self, src):
        r = self.rows[src]
        if r['var'].get() == NEW_AGENT:
            r['new_fr'].pack(side="left")
        else:
            r['new_fr'].pack_forget()

    def ok(self):
        result = {}
        for src, r in self.rows.items():
            choice = r['var'].get()
            if choice == NOT_SET:
                result[src] = ('none',)
            elif choice == NEW_AGENT:
                name = r['name_var'].get().strip()
                if not name:
                    messagebox.showerror("Новый агент",
                                         f"Для «{src}» укажите имя нового агента.",
                                         parent=self)
                    return
                try:
                    comm = float(r['comm_var'].get().strip().replace(',', '.'))
                except ValueError:
                    messagebox.showerror("Новый агент",
                                         f"Для «{name}» укажите комиссию числом.",
                                         parent=self)
                    return
                result[src] = ('new', name, comm)
            else:
                result[src] = ('agent', choice)
        self.result = result
        self.destroy()

