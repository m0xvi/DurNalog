# -*- coding: utf-8 -*-
"""Окно настроек: левое меню вместо Tabview, с вкладками приложения."""

from __future__ import annotations

import os
import threading
import tkinter as tk
import tkinter.ttk as ttk
import webbrowser
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING, Dict, Optional

import customtkinter as ctk

from theme import (
    AGENTS,
    COLORS,
    DB_PATH_DEMO,
    FONTS,
    FORMULAS,
    HOTELS,
    USERS,
)
from widgets import GhostButton, GreenButton, PrimaryButton, labeled_entry, style_treeview
from widgets import DangerGhost

if TYPE_CHECKING:
    from app import TurNalogApp


try:  # в живом приложении корневые модули импортируются; в стенде GUI — нет
    from updater import (REPO, UpdateError, app_dir as _upd_app_dir,
                         check_for_updates, download_release, is_frozen,
                         launch_install, plan_install, staging_dir,
                         verify_package)
    _UPDATER_OK = True
except Exception:
    _UPDATER_OK = False
    REPO = "m0xvi/DurNalog"


def _f(name: str):
    return FONTS.get(name)


class HotelTaxForm(ctk.CTkScrollableFrame):
    def __init__(self, master, hotel_id: str):
        super().__init__(master, fg_color="transparent")
        self.hotel_id = hotel_id
        h = HOTELS[hotel_id]

        ctk.CTkLabel(
            self,
            text=f"{h['name']} — параметры налога",
            anchor="w",
            font=_f("h2"),
            text_color=COLORS["text"],
        ).pack(anchor="w", pady=(0, 14))

        grid = ctk.CTkFrame(self, fg_color="transparent")
        grid.pack(fill="x")
        grid.grid_columnconfigure(1, weight=1)

        self.rate = self._row(grid, 0, "Ставка налога, %", "туристический налог", h["rate"])
        self.min_night = self._row(grid, 1, "Минимум за ночь, ₽", "нижняя граница", h["min_night"])
        self.threshold = self._row(
            grid, 2, "Порог метода «ставка»", "суточная ставка, ₽", h["threshold"]
        )

        self.tax_in_price = tk.BooleanVar(value=h.get("tax_in_price", True))
        self.auto_threshold = tk.BooleanVar(value=h.get("auto_threshold", True))
        self.tax_no_commission = tk.BooleanVar(value=h.get("tax_no_commission", True))
        self.include_extras_in_tax = tk.BooleanVar(
            value=h.get("include_extras_in_tax", hotel_id in ("chai", "dona")))

        options = [
            (
                self.tax_in_price,
                "Налог включён в стоимость",
                "Делитель 100 + ставка",
            ),
            (
                self.auto_threshold,
                "Порог — автоматически",
                "минимум × (100 + ставка) / ставка",
            ),
            (
                self.tax_no_commission,
                "Считать налог от стоимости с комиссией",
                "База = 100% цены (нетто из файла + комиссия агента)",
            ),
        ]
        if hotel_id in ("chai", "dona"):
            if hotel_id == "dona":
                extras_title = "Включать «Продление номера» в налоговую базу"
                extras_sub = "Учитывать только допуслугу с названием «Продление номера»"
            else:
                extras_title = "Включать допуслуги в налоговую базу"
                extras_sub = "Проживание + комиссия, затем все допуслуги, затем налог"
            options.append((
                self.include_extras_in_tax, extras_title, extras_sub,
            ))

        for var, title, sub in options:
            row = ctk.CTkFrame(
                self,
                fg_color="#FFFFFF",
                corner_radius=10,
                border_width=1,
                border_color=COLORS["line"],
            )
            row.pack(fill="x", pady=(8, 0))
            inner = ctk.CTkFrame(row, fg_color="transparent")
            inner.pack(fill="x", padx=12, pady=8)
            ctk.CTkCheckBox(
                inner,
                text=title,
                variable=var,
                font=_f("ui_b"),
                text_color=COLORS["text"],
                fg_color=COLORS["accent"],
                hover_color=COLORS["accent_hover"],
                border_color=COLORS["line"],
            ).pack(anchor="w")
            ctk.CTkLabel(
                inner, text=sub, anchor="w", font=_f("xs"), text_color=COLORS["muted"]
            ).pack(anchor="w", padx=(28, 0))

        self.adv_open = False
        self.adv_btn = GhostButton(
            self,
            "Шаблоны Excel-формул  ▸  правки для опытных",
            command=self._toggle_adv,
            height=36,
        )
        self.adv_btn.pack(fill="x", pady=(16, 0))

        self.adv = ctk.CTkFrame(
            self,
            fg_color="#FFFFFF",
            corner_radius=10,
            border_width=1,
            border_color=COLORS["line"],
        )
        box, self.f_daily = labeled_entry(self.adv, "Сут. ставка", h.get("formula_daily", FORMULAS["daily"]), mono=True)
        box.pack(fill="x", padx=14, pady=(12, 6))
        box, self.f_method = labeled_entry(self.adv, "Метод", h.get("formula_method", FORMULAS["method"]), mono=True)
        box.pack(fill="x", padx=14, pady=6)
        box, self.f_tax = labeled_entry(self.adv, "Налог", h.get("formula_tax", FORMULAS["tax"]), mono=True)
        box.pack(fill="x", padx=14, pady=6)
        ctk.CTkLabel(
            self.adv,
            text=FORMULAS["placeholders"],
            wraplength=640,
            justify="left",
            anchor="w",
            font=_f("xs"),
            text_color=COLORS["muted"],
        ).pack(fill="x", padx=14, pady=(4, 8))
        rowb = ctk.CTkFrame(self.adv, fg_color="transparent")
        rowb.pack(fill="x", padx=14, pady=(0, 12))
        GhostButton(rowb, "Показать формулы", command=self._preview, width=160).pack(
            side="left", padx=(0, 8)
        )
        GhostButton(rowb, "Сбросить по умолчанию", command=self._reset, width=190).pack(side="left")

    def _row(self, grid, r, title, sub, value):
        lab = ctk.CTkFrame(grid, fg_color="transparent")
        lab.grid(row=r, column=0, sticky="w", pady=6, padx=(0, 16))
        ctk.CTkLabel(lab, text=title, anchor="w", font=_f("ui"), text_color=COLORS["text"]).pack(
            anchor="w"
        )
        ctk.CTkLabel(lab, text=sub, anchor="w", font=_f("xs"), text_color=COLORS["muted"]).pack(
            anchor="w"
        )
        ent = ctk.CTkEntry(
            grid,
            height=36,
            width=200,
            corner_radius=8,
            border_color=COLORS["line"],
            fg_color="#FFFFFF",
            font=_f("ui"),
        )
        ent.grid(row=r, column=1, sticky="w", pady=6)
        ent.insert(0, value)
        return ent

    def _toggle_adv(self) -> None:
        self.adv_open = not self.adv_open
        if self.adv_open:
            self.adv.pack(fill="x", pady=(8, 0))
            self.adv_btn.configure(text="Шаблоны Excel-формул  ▾  правки для опытных")
        else:
            self.adv.pack_forget()
            self.adv_btn.configure(text="Шаблоны Excel-формул  ▸  правки для опытных")

    def _preview(self) -> None:
        messagebox.showinfo(
            "Формулы",
            f"Сут. ставка:\n{self.f_daily.get()}\n\nМетод:\n{self.f_method.get()}\n\nНалог:\n{self.f_tax.get()}",
            parent=self.winfo_toplevel(),
        )

    def _reset(self) -> None:
        for ent, val in (
            (self.f_daily, FORMULAS["daily"]),
            (self.f_method, FORMULAS["method"]),
            (self.f_tax, FORMULAS["tax"]),
        ):
            ent.delete(0, "end")
            ent.insert(0, val)

    def collect(self) -> dict:
        return {
            "hotel_id": self.hotel_id,
            "rate": self.rate.get(),
            "min_night": self.min_night.get(),
            "threshold": self.threshold.get(),
            "tax_in_price": self.tax_in_price.get(),
            "auto_threshold": self.auto_threshold.get(),
            "tax_no_commission": self.tax_no_commission.get(),
            "include_extras_in_tax": (self.include_extras_in_tax.get()
                                      if self.hotel_id in ("chai", "dona") else False),
            "formula_daily": self.f_daily.get(),
            "formula_method": self.f_method.get(),
            "formula_tax": self.f_tax.get(),
        }


class AgentsPane(ctk.CTkFrame):
    def __init__(self, master, app: "TurNalogApp"):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self.agents = list(AGENTS)
        self.current: Optional[str] = self.agents[0] if self.agents else None
        self.item_btns: Dict[str, ctk.CTkButton] = {}

        ctk.CTkLabel(
            self,
            text="Агенты: комиссии, синонимы, отели",
            anchor="w",
            font=_f("h2"),
            text_color=COLORS["text"],
        ).pack(anchor="w", pady=(0, 10))

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 10))
        GreenButton(bar, "+  Добавить", command=self._add, width=120, height=32).pack(
            side="left", padx=(0, 6)
        )
        DangerGhost(bar, "Удалить", command=self._delete, width=90).pack(side="left", padx=6)
        GhostButton(bar, "Импорт…", command=lambda: self._stub("Импорт агентов"), width=90).pack(
            side="left", padx=6
        )
        GhostButton(bar, "Экспорт…", command=lambda: self._stub("Экспорт агентов"), width=90).pack(
            side="left", padx=6
        )
        GhostButton(bar, "Проверить", command=lambda: self._stub("Проверка синонимов"), width=100).pack(
            side="left", padx=6
        )

        split = ctk.CTkFrame(self, fg_color="transparent")
        split.pack(fill="both", expand=True)
        split.grid_columnconfigure(1, weight=1)
        split.grid_rowconfigure(0, weight=1)

        left = ctk.CTkFrame(
            split,
            width=260,
            fg_color="#FFFFFF",
            corner_radius=10,
            border_width=1,
            border_color=COLORS["line"],
        )
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        split.grid_columnconfigure(0, minsize=260)
        self.search = ctk.CTkEntry(
            left,
            placeholder_text="Найти агента…",
            height=32,
            border_color=COLORS["line"],
            fg_color="#FFFFFF",
        )
        self.search.pack(fill="x", padx=10, pady=10)
        self.search.bind("<KeyRelease>", lambda _e: self._rebuild_list())
        self.listbox = ctk.CTkScrollableFrame(left, fg_color="#FFFFFF")
        self.listbox.pack(fill="both", expand=True, padx=4, pady=(0, 8))

        right = ctk.CTkFrame(
            split,
            fg_color="#FFFFFF",
            corner_radius=10,
            border_width=1,
            border_color=COLORS["line"],
        )
        right.grid(row=0, column=1, sticky="nsew")
        form = ctk.CTkFrame(right, fg_color="transparent")
        form.pack(fill="both", expand=True, padx=18, pady=16)

        box, self.name_ent = labeled_entry(form, "Название", "")
        box.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            form, text="Применяется к", anchor="w", font=_f("xs_b"), text_color=COLORS["muted"]
        ).pack(anchor="w", pady=(0, 6))
        hotels_row = ctk.CTkFrame(form, fg_color="transparent")
        hotels_row.pack(fill="x", pady=(0, 10))
        self.hotel_vars = {}
        for hid, title in (("chai", "Чайковский"), ("dona", "Дона"), ("usadba", "Усадьба")):
            var = tk.BooleanVar(value=(hid == "dona"))
            self.hotel_vars[hid] = var
            cell = ctk.CTkFrame(
                hotels_row,
                fg_color="#F7F9FC",
                corner_radius=8,
                border_width=1,
                border_color=COLORS["line"],
            )
            cell.pack(side="left", padx=(0, 8))
            ctk.CTkCheckBox(
                cell,
                text=title,
                variable=var,
                font=_f("sm"),
                fg_color=COLORS["accent"],
                border_color=COLORS["line"],
            ).pack(padx=8, pady=6)

        box, self.syn_ent = labeled_entry(form, "Синонимы", "балтма, baltma")
        box.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(
            form,
            text="Комиссия по периодам     пустые даты = базовая ставка",
            anchor="w",
            font=_f("xs_b"),
            text_color=COLORS["muted"],
        ).pack(anchor="w", pady=(0, 6))
        prow = ctk.CTkFrame(form, fg_color="transparent")
        prow.pack(fill="x")
        self.p_from = ctk.CTkEntry(
            prow, placeholder_text="с", height=36, width=120, border_color=COLORS["line"]
        )
        self.p_from.pack(side="left", padx=(0, 6))
        self.p_to = ctk.CTkEntry(
            prow, placeholder_text="по", height=36, width=120, border_color=COLORS["line"]
        )
        self.p_to.pack(side="left", padx=6)
        self.p_pct = ctk.CTkEntry(prow, height=36, width=64, border_color=COLORS["line"], justify="center")
        self.p_pct.pack(side="left", padx=6)
        self.p_pct.insert(0, "20")
        ctk.CTkLabel(prow, text="%", font=_f("ui"), text_color=COLORS["muted"]).pack(side="left")
        GhostButton(form, "+ период", command=lambda: self._stub("Добавить период"), width=100).pack(
            anchor="w", pady=(8, 0)
        )

        self._rebuild_list()
        self._load_current()

    def _stub(self, title: str) -> None:
        messagebox.showinfo(title, "Макет GUI — логика не подключена.", parent=self.winfo_toplevel())

    def _rebuild_list(self) -> None:
        q = self.search.get().strip().lower()
        for w in self.listbox.winfo_children():
            w.destroy()
        self.item_btns.clear()
        for name in self.agents:
            if q and q not in name.lower():
                continue
            btn = ctk.CTkButton(
                self.listbox,
                text=name,
                anchor="w",
                height=34,
                corner_radius=6,
                fg_color=COLORS["accent_soft"] if name == self.current else "transparent",
                hover_color="#F5F8FC",
                text_color=COLORS["text"],
                font=_f("sm"),
                command=lambda n=name: self._select(n),
            )
            btn.pack(fill="x", pady=1)
            self.item_btns[name] = btn

    def _select(self, name: str) -> None:
        self.current = name
        self._rebuild_list()
        self._load_current()

    def _load_current(self) -> None:
        self.name_ent.delete(0, "end")
        if self.current:
            self.name_ent.insert(0, self.current)

    def _add(self) -> None:
        dlg = ctk.CTkInputDialog(text="Название агента:", title="Новый агент")
        name = dlg.get_input()
        if name:
            self.agents.append(name)
            self.current = name
            self._rebuild_list()
            self._load_current()

    def _delete(self) -> None:
        if not self.current:
            return
        if not messagebox.askyesno("Удалить", f"Удалить агента «{self.current}»?", parent=self.winfo_toplevel()):
            return
        self.agents = [a for a in self.agents if a != self.current]
        self.current = self.agents[0] if self.agents else None
        self._rebuild_list()
        self._load_current()


class UsersPane(ctk.CTkFrame):
    def __init__(self, master, app: "TurNalogApp"):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self.users = list(USERS)

        ctk.CTkLabel(
            self,
            text="Пользователи и права доступа",
            anchor="w",
            font=_f("h2"),
            text_color=COLORS["text"],
        ).pack(anchor="w")
        ctk.CTkLabel(
            self,
            text="Пользователь отеля видит только свой отель и его историю.",
            anchor="w",
            font=_f("sm"),
            text_color=COLORS["muted"],
        ).pack(anchor="w", pady=(4, 12))

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", pady=(0, 10))
        GreenButton(bar, "+  Добавить", command=self._add, width=120, height=32).pack(
            side="left", padx=(0, 6)
        )
        PrimaryButton(bar, "Сменить пароль", command=self._password, width=140).pack(
            side="left", padx=6
        )
        DangerGhost(bar, "Удалить", command=self._delete, width=90).pack(side="left", padx=6)

        wrap = ctk.CTkFrame(
            self, fg_color="#FFFFFF", corner_radius=10, border_width=1, border_color=COLORS["line"]
        )
        wrap.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(wrap, fg_color="#FFFFFF")
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        style_treeview()
        self.tree = ttk.Treeview(
            inner,
            columns=("login", "role", "hotel"),
            show="headings",
            style="Tur.Treeview",
            selectmode="browse",
        )
        for c, t, w in (("login", "Логин", 180), ("role", "Роль", 180), ("hotel", "Отель", 180)):
            self.tree.heading(c, text=t, anchor="w")
            self.tree.column(c, width=w, anchor="w")
        self.tree.pack(fill="both", expand=True)
        self._reload()

    def _reload(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for i, (login, role, hotel) in enumerate(self.users):
            self.tree.insert("", "end", iid=str(i), values=(login, role, hotel))
        if self.users:
            self.tree.selection_set("0")

    def _selected(self):
        sel = self.tree.selection()
        if not sel:
            return None
        return self.users[int(sel[0])]

    def _add(self) -> None:
        messagebox.showinfo(
            "Пользователь",
            "Макет GUI. Здесь будет диалог логина / роли / отеля / пароля из 2.6.",
            parent=self.winfo_toplevel(),
        )

    def _password(self) -> None:
        row = self._selected()
        if not row:
            return
        dlg = ctk.CTkInputDialog(text=f"Новый пароль для {row[0]}:", title="Сменить пароль")
        _ = dlg.get_input()

    def _delete(self) -> None:
        row = self._selected()
        if not row:
            return
        if row[0] == "admin":
            messagebox.showwarning("Удалить", "Нельзя удалить admin.", parent=self.winfo_toplevel())
            return
        if messagebox.askyesno("Удалить", f"Удалить {row[0]}?", parent=self.winfo_toplevel()):
            self.users = [u for u in self.users if u[0] != row[0]]
            self._reload()


class DataPane(ctk.CTkFrame):
    def __init__(self, master, app: "TurNalogApp"):
        super().__init__(master, fg_color="transparent")
        self.app = app
        ctk.CTkLabel(
            self, text="База данных: расположение", anchor="w", font=_f("h2"), text_color=COLORS["text"]
        ).pack(anchor="w", pady=(0, 10))
        self.path_ent = ctk.CTkEntry(
            self,
            height=38,
            corner_radius=8,
            border_color=COLORS["line"],
            fg_color="#F3F6FA",
            font=_f("mono"),
        )
        self.path_ent.pack(fill="x")
        self.path_ent.insert(0, DB_PATH_DEMO)
        self.path_ent.configure(state="disabled")
        ctk.CTkLabel(
            self,
            text="В базе — история отчётов, дубли кварталов, пользователи.\n"
            "Общая сетевая папка = общая история и пользователи для всех ПК.\n"
            "Смена пути применяется сразу, перезапуск не нужен.",
            justify="left",
            anchor="w",
            font=_f("sm"),
            text_color=COLORS["muted"],
        ).pack(anchor="w", pady=(12, 14))
        row = ctk.CTkFrame(self, fg_color="transparent")
        row.pack(anchor="w")
        PrimaryButton(row, "Сменить папку…", command=self._change, width=150).pack(side="left", padx=(0, 8))
        GhostButton(row, "Сбросить (рядом с приложением)", command=self._reset, width=250).pack(
            side="left"
        )

    def _change(self) -> None:
        path = filedialog.askdirectory(parent=self.winfo_toplevel(), title="Папка базы данных")
        if path:
            self.path_ent.configure(state="normal")
            self.path_ent.delete(0, "end")
            self.path_ent.insert(0, path + "\\turnalog_history.db")
            self.path_ent.configure(state="disabled")
            self.app.set_status("Путь БД обновлён (демо GUI)")

    def _reset(self) -> None:
        self.path_ent.configure(state="normal")
        self.path_ent.delete(0, "end")
        self.path_ent.insert(0, DB_PATH_DEMO)
        self.path_ent.configure(state="disabled")

class TlCampaignsPane(ctk.CTkScrollableFrame):
    """Список кампаний TravelLine, где скидка остаётся в налоговой базе."""

    def __init__(self, master, app: "TurNalogApp"):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self.boxes = {}
        cfg = getattr(app, "cfg", {}) or {}
        tl_cfg = cfg.get("travelline", {}) if isinstance(cfg, dict) else {}
        if not isinstance(tl_cfg, dict):
            tl_cfg = {}

        ctk.CTkLabel(
            self, text="Кампании TravelLine: считать налог со скидкой",
            anchor="w", font=_f("h2"), text_color=COLORS["text"]
        ).pack(anchor="w", pady=(0, 8))
        ctk.CTkLabel(
            self,
            text=("По правилу расчёта скидка обычно добавляется обратно к "
                  "налоговой базе (это не отдельная комиссия из API). "
                  "Для перечисленных ниже кампаний она остаётся вычтенной.\n"
                  "Вставляйте точное название из поля «discounts[].name», "
                  "по одному названию на строку. Список задаётся отдельно по отелям."),
            anchor="w", justify="left", wraplength=700, font=_f("sm"),
            text_color=COLORS["muted"]
        ).pack(anchor="w", pady=(0, 12))

        for code, hotel_name in (("tch", "Чайковский"), ("dna", "Дона")):
            card = ctk.CTkFrame(
                self, fg_color="#FFFFFF", corner_radius=10,
                border_width=1, border_color=COLORS["line"])
            card.pack(fill="x", pady=(0, 12))
            inner = ctk.CTkFrame(card, fg_color="transparent")
            inner.pack(fill="x", padx=14, pady=12)
            ctk.CTkLabel(
                inner, text=hotel_name, anchor="w", font=_f("ui_b"),
                text_color=COLORS["text"]
            ).pack(anchor="w", pady=(0, 5))
            textbox = ctk.CTkTextbox(
                inner, height=112, wrap="word", border_color=COLORS["line"],
                fg_color="#F8FAFC", font=_f("mono"))
            textbox.pack(fill="x")
            hotel_cfg = tl_cfg.get(code, {})
            if not isinstance(hotel_cfg, dict):
                hotel_cfg = {}
            saved = hotel_cfg.get("campaigns_with_discount", [])
            if isinstance(saved, str):
                saved = saved.splitlines()
            textbox.insert("1.0", "\n".join(str(x) for x in saved or []))
            self.boxes[code] = textbox

        ctk.CTkLabel(
            self,
            text=("Пустой список = скидки всех кампаний добавляются обратно. "
                  "Изменение списка создаёт отдельную версию кэша TravelLine."),
            anchor="w", justify="left", wraplength=700, font=_f("xs"),
            text_color=COLORS["muted"]
        ).pack(anchor="w", pady=(0, 8))

    def collect(self):
        result = {}
        for code, textbox in self.boxes.items():
            names, seen = [], set()
            for value in textbox.get("1.0", "end").splitlines():
                name = value.strip()
                key = " ".join(name.casefold().replace("ё", "е").split())
                if key and key not in seen:
                    seen.add(key)
                    names.append(name)
            result[code] = names
        return result


class UpdatesPane(ctk.CTkFrame):
    """Проверка и установка обновлений с GitHub Releases."""

    def __init__(self, master, app: "TurNalogApp"):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self._info = None
        self._plan = None
        self._busy = False

        ctk.CTkLabel(self, text="Обновления приложения", anchor="w",
                     font=_f("h2"), text_color=COLORS["text"]).pack(anchor="w", pady=(0, 10))

        card = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=10,
                            border_width=1, border_color=COLORS["line"])
        card.pack(fill="x", pady=(0, 12))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=16, pady=12)

        cur = getattr(app, "app_version", "?")
        frozen = _UPDATER_OK and is_frozen()
        build = "собранная версия (EXE)" if frozen else "запуск из исходников"
        ctk.CTkLabel(inner, text=f"Текущая версия:  {cur}   ·   {build}",
                     anchor="w", font=_f("ui"), text_color=COLORS["text"]).pack(anchor="w")
        self.status_lbl = ctk.CTkLabel(inner, text="Нажмите «Проверить сейчас».",
                                       anchor="w", font=_f("sm"), text_color=COLORS["muted"])
        self.status_lbl.pack(anchor="w", pady=(4, 0))
        last = (self._upd_cfg().get("last_check") or "")[:16].replace("T", " ")
        if last:
            ctk.CTkLabel(inner, text=f"Последняя проверка: {last}",
                         anchor="w", font=_f("xs"),
                         text_color=COLORS["muted"]).pack(anchor="w")

        ctk.CTkLabel(self, text="Что нового в найденной версии:", anchor="w",
                     font=_f("xs_b"), text_color=COLORS["muted"]).pack(anchor="w", pady=(0, 4))
        self.notes = ctk.CTkTextbox(self, height=90, wrap="word",
                                    border_color=COLORS["line"], fg_color="#FFFFFF")
        self.notes.pack(fill="x", pady=(0, 12))
        self.notes.insert("0.0", "—")
        self.notes.configure(state="disabled")

        self.prog = ctk.CTkProgressBar(self, height=10)
        self.prog.set(0)
        self.prog_lbl = ctk.CTkLabel(self, text="", anchor="w",
                                     font=_f("xs"), text_color=COLORS["muted"])

        self.btn_row = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_row.pack(fill="x", pady=(0, 12))
        self.check_btn = PrimaryButton(self.btn_row, "Проверить сейчас",
                                       command=self._check, width=170)
        self.check_btn.pack(side="left", padx=(0, 8))
        self.install_btn = GreenButton(self.btn_row, "Скачать и установить",
                                       command=self._install, width=210, height=32)
        self.install_btn.pack(side="left", padx=8)
        self.install_btn.configure(state="disabled")
        GhostButton(self.btn_row, "Страница релиза",
                    command=self._open_page, width=150).pack(side="left", padx=8)

        opts = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=10,
                            border_width=1, border_color=COLORS["line"])
        opts.pack(fill="x")
        oinner = ctk.CTkFrame(opts, fg_color="transparent")
        oinner.pack(fill="x", padx=16, pady=12)
        cfg = self._upd_cfg()
        self.auto_check_var = tk.BooleanVar(value=cfg.get("auto_check", True))
        self.auto_dl_var = tk.BooleanVar(value=cfg.get("auto_download", False))
        first = True
        for var, title, sub in (
            (self.auto_check_var, "Проверять обновления при запуске",
             "Тихая проверка; при находке — уведомление в приложении"),
            (self.auto_dl_var, "Скачивать обновление автоматически",
             "Файл загрузится сам, установка — кнопкой после перезапуска"),
        ):
            ctk.CTkCheckBox(oinner, text=title, variable=var, font=_f("ui_b"),
                            text_color=COLORS["text"], fg_color=COLORS["accent"],
                            border_color=COLORS["line"],
                            command=self._save_opts).pack(anchor="w", pady=(0 if first else 8, 0))
            first = False
            ctk.CTkLabel(oinner, text=sub, anchor="w", font=_f("xs"),
                         text_color=COLORS["muted"]).pack(anchor="w", padx=(28, 0))
        if not _UPDATER_OK:
            self._status("Модуль обновления недоступен в этом окружении.", err=True)
            self.check_btn.configure(state="disabled")
        elif not frozen:
            ctk.CTkLabel(self, text="Установка из приложения доступна только в собранной EXE-версии.",
                         anchor="w", font=_f("xs"),
                         text_color=COLORS["muted"]).pack(anchor="w", pady=(8, 0))

    # ----- конфиг -----
    def _upd_cfg(self):
        cfg = getattr(self.app, "cfg", None)
        if not isinstance(cfg, dict):
            return {}
        return cfg.setdefault("updates", {})

    def _save_cfg(self):
        cfg = getattr(self.app, "cfg", None)
        if not isinstance(cfg, dict):
            return
        try:
            from config_io import save_config
        except Exception:
            return
        try:
            save_config(cfg)
        except Exception:
            pass

    def _save_opts(self):
        if not isinstance(getattr(self.app, "cfg", None), dict):
            return
        uc = self._upd_cfg()
        uc["auto_check"] = bool(self.auto_check_var.get())
        uc["auto_download"] = bool(self.auto_dl_var.get())
        self._save_cfg()

    def _touch_last_check(self, version=None):
        if not isinstance(getattr(self.app, "cfg", None), dict):
            return
        from datetime import datetime
        uc = self._upd_cfg()
        uc["last_check"] = datetime.now().isoformat(timespec="seconds")
        if version:
            uc["last_version"] = version
        self._save_cfg()

    # ----- UI-хелперы (только из главного потока) -----
    def _status(self, text, err=False):
        try:
            self.status_lbl.configure(
                text=text, text_color=COLORS["red"] if err else COLORS["muted"])
        except Exception:
            pass

    def _set_notes(self, text):
        try:
            self.notes.configure(state="normal")
            self.notes.delete("0.0", "end")
            self.notes.insert("0.0", (text or "").strip()[:2000] or "—")
            self.notes.configure(state="disabled")
        except Exception:
            pass

    def _show_progress(self, show):
        try:
            if show:
                self.prog.pack(fill="x", pady=(0, 4), before=self.btn_row)
                self.prog_lbl.pack(fill="x", pady=(0, 8), before=self.btn_row)
            else:
                self.prog.pack_forget()
                self.prog_lbl.pack_forget()
        except Exception:
            pass

    def _ui(self, fn):
        def wrap():
            try:
                if self.winfo_exists():
                    fn()
            except Exception:
                pass
        try:
            self.after(0, wrap)
        except Exception:
            pass

    # ----- проверка -----
    def _check(self):
        if self._busy or not _UPDATER_OK:
            return
        self._busy = True
        self.check_btn.configure(state="disabled")
        self.install_btn.configure(state="disabled")
        self._status("Проверка обновлений…")
        cur = getattr(self.app, "app_version", "?")
        threading.Thread(target=self._check_worker, args=(cur,), daemon=True).start()

    def _check_worker(self, cur):
        try:
            info = check_for_updates(cur)
            err = None
        except UpdateError as e:
            info, err = None, str(e)
        except Exception as e:  # noqa: BLE001
            info, err = None, f"Неожиданная ошибка: {e}"
        self._ui(lambda: self._check_done(info, err))

    def _check_done(self, info, err):
        self._busy = False
        self.check_btn.configure(state="normal")
        if err:
            self._status(err, err=True)
            return
        self._touch_last_check(info.version if info else None)
        if info is None:
            self._status("У вас последняя версия.")
            self._set_notes("—")
            return
        self._info = info
        self._plan = None
        self.install_btn.configure(text="Скачать и установить")
        self._set_notes(info.notes)
        try:
            writable = os.access(_upd_app_dir(), os.W_OK)
        except Exception:
            writable = True
        if not (is_frozen() and os.name == "nt"):
            self._status(f"Доступна версия {info.version} ({info.tag}). "
                         "Установка — только в EXE-версии.")
        elif not writable:
            self._status(f"Доступна версия {info.version}, но нет прав на запись "
                         "в папку программы.", err=True)
        else:
            self._status(f"Доступна версия {info.version} ({info.tag}).")
            self.install_btn.configure(state="normal")

    # ----- установка -----
    def _install(self):
        if self._busy or self._info is None:
            return
        if self._plan is not None:
            self._ask_restart()
            return
        self._busy = True
        self.check_btn.configure(state="disabled")
        self.install_btn.configure(state="disabled")
        self._show_progress(True)
        self.prog.set(0)
        self.prog_lbl.configure(text="")
        threading.Thread(target=self._dl_worker, args=(self._info,), daemon=True).start()

    def _dl_worker(self, info):
        try:
            zp = download_release(
                info, staging_dir(info.tag),
                progress=lambda d, t: self._ui(lambda: self._on_progress(d, t)))
            verify_package(zp)
            plan = plan_install(zp)
            err = None
        except UpdateError as e:
            plan, err = None, str(e)
        except Exception as e:  # noqa: BLE001
            plan, err = None, f"Неожиданная ошибка: {e}"
        self._ui(lambda: self._dl_done(plan, err))

    def _on_progress(self, done, total):
        try:
            if total:
                self.prog.set(min(1.0, done / total))
                self.prog_lbl.configure(text=f"{done / 1e6:.1f} / {total / 1e6:.1f} МБ")
            else:
                self.prog_lbl.configure(text=f"{done / 1e6:.1f} МБ")
            self._status("Скачивание обновления…")
        except Exception:
            pass

    def _dl_done(self, plan, err):
        self._busy = False
        self.check_btn.configure(state="normal")
        self._show_progress(False)
        if err:
            self._status(err, err=True)
            return
        self._plan = plan
        self.install_btn.configure(text="Перезапустить и установить", state="normal")
        self._ask_restart()

    def _ask_restart(self):
        top = self.winfo_toplevel()
        ver = self._info.version if self._info else ""
        if not messagebox.askyesno(
                "Обновление загружено",
                f"Версия {ver} готова к установке.\n\n"
                "Приложение закроется, файлы обновятся и программа запустится снова.\n"
                "Ваши данные (настройки, история) сохранятся.\n\nУстановить сейчас?",
                parent=top):
            self._status(f"Версия {ver} загружена. Нажмите «Перезапустить и установить», "
                         "когда будете готовы.")
            return
        try:
            launch_install(self._plan)
        except UpdateError as e:
            messagebox.showerror("Обновление", str(e), parent=top)
            return
        try:
            top.destroy()
        except Exception:
            pass
        try:
            self.app.destroy()
        except Exception:
            pass

    def _open_page(self):
        url = self._info.page_url if self._info else f"https://github.com/{REPO}/releases"
        try:
            webbrowser.open(url)
        except Exception as e:
            self._status(f"Не удалось открыть браузер: {e}", err=True)


class SettingsWindow(ctk.CTkToplevel):
    def __init__(self, app: "TurNalogApp"):
        super().__init__(app)
        self.app = app
        self.title("Настройки")
        self.geometry("1020x680")
        self.minsize(900, 600)
        self.configure(fg_color=COLORS["bg"])
        self.transient(app)
        self.grab_set()
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        shell = ctk.CTkFrame(self, fg_color=COLORS["bg"], corner_radius=0)
        shell.pack(fill="both", expand=True)

        foot = ctk.CTkFrame(shell, fg_color="#FFFFFF", corner_radius=0)
        foot.pack(side="bottom", fill="x")
        inner = ctk.CTkFrame(foot, fg_color="transparent")
        inner.pack(fill="x", padx=22, pady=12)
        GreenButton(inner, "Сохранить", command=self._save, width=130, height=36).pack(
            side="right", padx=(8, 0)
        )
        GhostButton(inner, "Отмена", command=self._cancel, width=110, height=36).pack(side="right")

        header = ctk.CTkFrame(shell, fg_color="#FFFFFF", corner_radius=0)
        header.pack(side="top", fill="x")
        titles = ctk.CTkFrame(header, fg_color="transparent")
        titles.pack(fill="x", padx=22, pady=14)
        ctk.CTkLabel(
            titles, text="Настройки расчёта", anchor="w", font=_f("h2"), text_color=COLORS["text"]
        ).pack(anchor="w")
        ctk.CTkLabel(
            titles,
            text="Параметры и формулы налога — отдельно по каждому отелю. Применяются к новым обработкам.",
            anchor="w",
            font=_f("sm"),
            text_color=COLORS["muted"],
        ).pack(anchor="w")

        body = ctk.CTkFrame(shell, fg_color="transparent")
        body.pack(fill="both", expand=True)

        self.nav = ctk.CTkFrame(body, width=196, fg_color="#FFFFFF", corner_radius=0)
        self.nav.pack(side="left", fill="y")
        self.nav.pack_propagate(False)

        self.host = ctk.CTkFrame(body, fg_color=COLORS["bg"], corner_radius=0)
        self.host.pack(side="left", fill="both", expand=True)

        self.panes: Dict[str, ctk.CTkBaseClass] = {}
        self.nav_btns: Dict[str, ctk.CTkButton] = {}
        self._nav_label("ОТЕЛИ")
        self.panes["chai"] = HotelTaxForm(self.host, "chai")
        self.panes["dona"] = HotelTaxForm(self.host, "dona")
        self.panes["usadba"] = HotelTaxForm(self.host, "usadba")
        self._nav_btn("chai", "Чайковский")
        self._nav_btn("dona", "Дона")
        self._nav_btn("usadba", "Усадьба")
        self._nav_label("СПРАВОЧНИКИ")
        self.panes["agents"] = AgentsPane(self.host, app)
        self.panes["users"] = UsersPane(self.host, app)
        self._nav_btn("agents", "Агенты")
        self._nav_btn("users", "Пользователи")
        self._nav_label("СИСТЕМА")
        self.panes["data"] = DataPane(self.host, app)
        self._nav_btn("data", "База данных")
        self.panes["tl_campaigns"] = TlCampaignsPane(self.host, app)
        self._nav_btn("tl_campaigns", "Кампании TL")
        self.panes["updates"] = UpdatesPane(self.host, app)
        self._nav_btn("updates", "Обновления")

        self._show("chai")
        self.geometry("1020x680")
        self.after(40, self._center)

    def _center(self) -> None:
        self.update_idletasks()
        px, py = self.app.winfo_rootx(), self.app.winfo_rooty()
        pw, ph = self.app.winfo_width(), self.app.winfo_height()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

    def _nav_label(self, text: str) -> None:
        ctk.CTkLabel(
            self.nav,
            text=text,
            anchor="w",
            font=_f("xs_b"),
            text_color=COLORS["faint"],
        ).pack(fill="x", padx=16, pady=(14, 4))

    def _nav_btn(self, key: str, text: str) -> None:
        btn = ctk.CTkButton(
            self.nav,
            text=text,
            anchor="w",
            height=34,
            corner_radius=8,
            fg_color="transparent",
            hover_color="#F3F6FA",
            text_color=COLORS["text"],
            font=_f("ui"),
            command=lambda k=key: self._show(k),
        )
        btn.pack(fill="x", padx=10, pady=1)
        self.nav_btns[key] = btn

    def _show(self, key: str) -> None:
        for k, pane in self.panes.items():
            pane.pack_forget()
        for k, btn in self.nav_btns.items():
            if k == key:
                btn.configure(fg_color=COLORS["accent_soft"], text_color=COLORS["accent"], font=_f("ui_b"))
            else:
                btn.configure(fg_color="transparent", text_color=COLORS["text"], font=_f("ui"))
        self.panes[key].pack(fill="both", expand=True, padx=22, pady=18)

    def _save(self) -> None:
        payload = {
            "chai": self.panes["chai"].collect(),
            "dona": self.panes["dona"].collect(),
            "usadba": self.panes["usadba"].collect(),
            "db": self.panes["data"].path_ent.get(),
        }
        self.app.hook_save_settings(payload)
        self.destroy()

    def _cancel(self) -> None:
        self.destroy()
