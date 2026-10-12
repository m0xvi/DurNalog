# -*- coding: utf-8 -*-
"""Экраны: обработка отеля и история. Только GUI."""

from __future__ import annotations

import tkinter as tk
import tkinter.ttk as ttk
from tkinter import filedialog, messagebox
from typing import TYPE_CHECKING

import customtkinter as ctk

from theme import COLORS, FONTS, HISTORY_ROWS, HOTEL_PILL, HOTELS
from widgets import (
    Card,
    Chip,
    DropZone,
    GhostButton,
    GreenButton,
    PrimaryButton,
    StatCard,
    StepBadge,
    style_treeview,
)

if TYPE_CHECKING:
    from app import TurNalogApp


def _f(name: str):
    return FONTS.get(name)


class ProcessView(ctk.CTkFrame):
    def __init__(self, master, app: "TurNalogApp"):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self.hotel_id = "chai"

        self.header = ctk.CTkFrame(self, fg_color="transparent")
        self.header.pack(fill="x", pady=(0, 14))
        self.title_lbl = ctk.CTkLabel(
            self.header, text="", anchor="w", font=_f("h1"), text_color=COLORS["text"]
        )
        self.title_lbl.pack(anchor="w")
        self.sub_lbl = ctk.CTkLabel(
            self.header, text="", anchor="w", font=_f("sm"), text_color=COLORS["muted"]
        )
        self.sub_lbl.pack(anchor="w", pady=(4, 0))
        self.chips = ctk.CTkFrame(self.header, fg_color="transparent")
        self.chips.pack(anchor="w", pady=(10, 0))

        steps = ctk.CTkFrame(self, fg_color="transparent")
        steps.pack(fill="x")
        steps.grid_columnconfigure(0, weight=3, uniform="s")
        steps.grid_columnconfigure(1, weight=2, uniform="s")

        self.card1 = Card(steps)
        self.card1.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        self._build_files(self.card1)

        self.card2 = Card(steps)
        self.card2.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        self._build_process(self.card2)

        stats = ctk.CTkFrame(self, fg_color="transparent")
        stats.pack(fill="x", pady=(14, 0))
        stats.grid_columnconfigure((0, 1, 2), weight=1)
        self.stat_last = StatCard(stats, "Последняя обработка", "—", "")
        self.stat_tax = StatCard(stats, "Налог за квартал", "—", "по последнему результату")
        self.stat_dup = StatCard(stats, "Дубли / переход кв.", "—", "дубли · фиолетовые строки")
        self.stat_last.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        self.stat_tax.grid(row=0, column=1, sticky="nsew", padx=6)
        self.stat_dup.grid(row=0, column=2, sticky="nsew", padx=(6, 0))

        note = ctk.CTkLabel(
            self,
            text="Карточки статистики — демо GUI. Подключать только если сумма налога уже есть в БД.",
            font=_f("xs"),
            text_color=COLORS["faint"],
            anchor="w",
        )
        note.pack(anchor="w", pady=(8, 0))

    def _build_files(self, card: Card) -> None:
        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=16, pady=(14, 8))
        self.step1 = StepBadge(head, "1")
        self.step1.pack(side="left")
        ctk.CTkLabel(
            head, text="  Исходные файлы", font=_f("h3"), text_color=COLORS["text"]
        ).pack(side="left")
        ctk.CTkLabel(
            head, text="Excel · xlsx / xls", font=_f("xs"), text_color=COLORS["muted"]
        ).pack(side="right")

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        self.drop_data = DropZone(
            body,
            "Данные — брони",
            "Выгрузка броней: даты, сумма, источник, гости",
            on_pick=self._pick_data,
        )
        self.drop_data.pack(fill="x", pady=(0, 8))

        self.drop_rooms = DropZone(
            body,
            "Номера — расселение",
            "Гости по комнатам, для разбивки по корпусам",
            on_pick=self._pick_rooms,
        )
        self.drop_rooms.pack(fill="x", pady=(0, 10))

        btns = ctk.CTkFrame(body, fg_color="transparent")
        btns.pack(fill="x")
        GhostButton(
            btns,
            "Колонки «Данные»",
            command=lambda: self.app.open_column_map("data"),
            width=160,
        ).pack(side="left", padx=(0, 8))
        GhostButton(
            btns,
            "Колонки «Номера»",
            command=lambda: self.app.open_column_map("rooms"),
            width=160,
        ).pack(side="left")
        self.tl_btn = GhostButton(
            btns,
            "↓ TravelLine",
            command=lambda: self.app.hook_tl_import(self.hotel_id),
            width=160,
        )
        self.tl_btn.pack(side="left", padx=(8, 0))

    def _build_process(self, card: Card) -> None:
        head = ctk.CTkFrame(card, fg_color="transparent")
        head.pack(fill="x", padx=16, pady=(14, 8))
        self.step2 = StepBadge(head, "2")
        self.step2.pack(side="left")
        ctk.CTkLabel(
            head, text="  Обработка", font=_f("h3"), text_color=COLORS["text"]
        ).pack(side="left")

        body = ctk.CTkFrame(card, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        ctk.CTkLabel(
            body, text="Квартал отчёта", anchor="w", font=_f("xs_b"), text_color=COLORS["muted"]
        ).pack(anchor="w", pady=(0, 6))

        row = ctk.CTkFrame(body, fg_color="transparent")
        row.pack(fill="x")
        self.qseg = ctk.CTkSegmentedButton(
            row,
            values=["I", "II", "III", "IV"],
            fg_color="#D9E1EC",
            selected_color="#FFFFFF",
            selected_hover_color="#FFFFFF",
            unselected_color="#D9E1EC",
            unselected_hover_color="#CBD5E1",
            text_color=COLORS["text"],
            font=_f("sm_b"),
            height=34,
            width=260,
            dynamic_resizing=False,
        )
        self.qseg.pack(side="left", fill="x", expand=True, padx=(0, 0))
        self.qseg.set("III")
        self.year = ctk.CTkEntry(
            row,
            width=84,
            height=32,
            justify="center",
            corner_radius=8,
            border_color=COLORS["line"],
            fg_color="#FFFFFF",
            font=_f("ui"),
        )
        self.year.pack(side="left", padx=(8, 0))
        self.year.insert(0, "2026")

        self.dups_var = tk.BooleanVar(value=True)
        chk_row = ctk.CTkFrame(body, fg_color="transparent")
        chk_row.pack(fill="x", pady=(12, 0))
        self.dups_chk = ctk.CTkCheckBox(
            chk_row,
            text="Проверять дубли по истории",
            variable=self.dups_var,
            font=_f("ui"),
            text_color=COLORS["text"],
            fg_color=COLORS["accent"],
            hover_color=COLORS["accent_hover"],
            border_color=COLORS["line"],
        )
        self.dups_chk.pack(anchor="w")
        ctk.CTkLabel(
            chk_row,
            text="Совпадения — на отдельный лист",
            font=_f("xs"),
            text_color=COLORS["muted"],
            anchor="w",
            wraplength=300,
            justify="left",
        ).pack(anchor="w", padx=(28, 0), pady=(2, 0))

        note = ctk.CTkFrame(
            body, fg_color=COLORS["purple_soft"], corner_radius=8, border_width=0
        )
        note.pack(fill="x", pady=(12, 14))
        ctk.CTkLabel(
            note,
            text="Брони с переходом квартала\nвыделяются фиолетовым в результате.",
            font=_f("xs"),
            text_color=COLORS["purple_text"],
            justify="left",
            anchor="w",
        ).pack(padx=11, pady=9, anchor="w")

        self.run_btn = GreenButton(
            body, "▶   Обработать и сохранить", command=self._run
        )
        self.run_btn.pack(fill="x")
        self.run_btn.configure(state="disabled", fg_color=COLORS["green_dis"])

    def bind_hotel(self, hotel_id: str) -> None:
        self.hotel_id = hotel_id
        h = HOTELS[hotel_id]
        self.title_lbl.configure(text=h["title"])
        self.sub_lbl.configure(text=h["sub"])
        for ch in self.chips.winfo_children():
            ch.destroy()
        for fg, tc, text in h["chips"]:
            Chip(self.chips, text, fg, tc).pack(side="left", padx=(0, 6))
        self.stat_last.set_values(h["last"], h["last_sub"])
        self.stat_tax.set_values(h["tax"], "по последнему результату")
        self.stat_dup.set_values(h["dup"], "дубли · фиолетовые строки")
        self.drop_data.clear()
        self.drop_rooms.clear()
        self.step1.set_done(False)
        self.step2.set_done(False)
        self._sync_run()
        # TravelLine есть у Чайковского (и будет у Доны); у Усадьбы другая PMS.
        if hotel_id in ("chai", "dona"):
            self.tl_btn.pack(side="left", padx=(8, 0))
        else:
            self.tl_btn.pack_forget()
        # Файл номеров у Доны/Усадьбы не обязателен — зона остаётся, чуть бледнее.
        if h["needs_rooms"]:
            self.drop_rooms.configure(fg_color=COLORS["drop"])
        else:
            self.drop_rooms.configure(fg_color="#F3F5F8")

    def _pick_data(self, zone: DropZone) -> None:
        path = filedialog.askopenfilename(
            parent=self.app,
            title="Данные — брони",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Все файлы", "*.*")],
        )
        if path:
            zone.set_path(path)
            self.step1.set_done(True)
            self._sync_run()
            self.app.set_status("Файлы выбраны — можно обрабатывать")

    def _pick_rooms(self, zone: DropZone) -> None:
        path = filedialog.askopenfilename(
            parent=self.app,
            title="Номера — расселение",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Все файлы", "*.*")],
        )
        if path:
            zone.set_path(path)
            self.app.set_status("Файл номеров выбран")

    def set_tl_files(self, data_path: str, rooms_path: str) -> None:
        self.drop_data.set_path(data_path)
        self.drop_rooms.set_path(rooms_path)
        self.step1.set_done(True)
        self._sync_run()

    def _sync_run(self) -> None:
        ready = bool(self.drop_data.path)
        if ready:
            self.run_btn.configure(state="normal", fg_color=COLORS["green"])
        else:
            self.run_btn.configure(state="disabled", fg_color=COLORS["green_dis"])

    def _run(self) -> None:
        if not self.drop_data.path:
            messagebox.showwarning("Турналог", "Выберите файл «Данные».", parent=self.app)
            return
        dest = filedialog.asksaveasfilename(
            parent=self.app,
            title="Сохранить результат",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=self._suggest_name(),
        )
        if not dest:
            return
        self.app.hook_process(
            hotel_id=self.hotel_id,
            data_path=self.drop_data.path,
            rooms_path=self.drop_rooms.path,
            quarter=self.qseg.get(),
            year=self.year.get().strip(),
            check_dups=self.dups_var.get(),
            dest_path=dest,
        )
        self.step2.set_done(True)

    def _suggest_name(self) -> str:
        h = HOTELS[self.hotel_id]["name"]
        return f"Турналог_{h}_{self.year.get().strip()}_{self.qseg.get()}.xlsx"


class HistoryView(ctk.CTkFrame):
    def __init__(self, master, app: "TurNalogApp"):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self.rows = list(HISTORY_ROWS)

        head = ctk.CTkFrame(self, fg_color="transparent")
        head.pack(fill="x", pady=(0, 8))
        ctk.CTkLabel(
            head,
            text="История обработанных отчётов",
            anchor="w",
            font=_f("h1"),
            text_color=COLORS["text"],
        ).pack(anchor="w")
        ctk.CTkLabel(
            head,
            text="Оригиналы и результаты — в локальной БД. Дубли на переходе кварталов — ниже.",
            anchor="w",
            font=_f("sm"),
            text_color=COLORS["muted"],
        ).pack(anchor="w", pady=(4, 0))

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(fill="x", pady=(8, 10))
        GhostButton(bar, "Обновить", command=self.refresh, width=100).pack(side="left", padx=(0, 6))
        PrimaryButton(bar, "Открыть оригинал", command=self.open_orig, width=150).pack(
            side="left", padx=6
        )
        PrimaryButton(bar, "Открыть результат", command=self.open_result, width=160).pack(
            side="left", padx=6
        )
        from widgets import DangerGhost

        DangerGhost(bar, "Удалить", command=self.delete_row, width=100).pack(side="right")

        table_card = Card(self)
        table_card.pack(fill="both", expand=True)
        inner = ctk.CTkFrame(table_card, fg_color="#FFFFFF", corner_radius=12)
        inner.pack(fill="both", expand=True, padx=1, pady=1)

        style_treeview()
        cols = ("num", "hotel", "quarter", "processed", "rows", "dups", "orig", "result")
        self.tree = ttk.Treeview(
            inner, columns=cols, show="headings", style="Tur.Treeview", selectmode="browse"
        )
        headings = {
            "num": "№",
            "hotel": "Отель",
            "quarter": "Квартал",
            "processed": "Обработан",
            "rows": "Строк",
            "dups": "Дубли",
            "orig": "Оригинал",
            "result": "Результат",
        }
        widths = {
            "num": 48,
            "hotel": 120,
            "quarter": 90,
            "processed": 140,
            "rows": 70,
            "dups": 70,
            "orig": 180,
            "result": 240,
        }
        for c in cols:
            self.tree.heading(c, text=headings[c], anchor="w")
            self.tree.column(c, width=widths[c], anchor="w", stretch=c in ("orig", "result"))

        vsb = ttk.Scrollbar(inner, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self.tree.tag_configure("chai", foreground=HOTEL_PILL["chai"][1])
        self.tree.tag_configure("dona", foreground=HOTEL_PILL["dona"][1])
        self.tree.tag_configure("usadba", foreground=HOTEL_PILL["usadba"][1])

        self.tree.bind("<<TreeviewSelect>>", self._on_select)

        self.bottom = Card(self)
        self.bottom.pack(fill="x", pady=(12, 0))
        pad = ctk.CTkFrame(self.bottom, fg_color="transparent")
        pad.pack(fill="x", padx=16, pady=12)
        ctk.CTkLabel(
            pad,
            text="Брони с переходом кварталов в выбранном отчёте",
            anchor="w",
            font=_f("ui_b"),
            text_color=COLORS["text"],
        ).pack(anchor="w")
        self.cross_lbl = ctk.CTkLabel(
            pad,
            text="Выберите отчёт в таблице — здесь появятся строки, пересекающие границу квартала (в Excel они фиолетовые).",
            anchor="w",
            justify="left",
            font=_f("sm"),
            text_color=COLORS["faint"],
            wraplength=900,
        )
        self.cross_lbl.pack(anchor="w", pady=(6, 4))

        self.refresh()

    def refresh(self) -> None:
        self.tree.delete(*self.tree.get_children())
        for row in self.rows:
            self.tree.insert(
                "",
                "end",
                iid=row["id"],
                values=(
                    row["id"],
                    row["hotel"],
                    row["quarter"],
                    row["processed"],
                    row["rows"],
                    row["dups"],
                    row["orig"],
                    row["result"],
                ),
                tags=(row["hotel_id"],),
            )
        if self.rows:
            self.tree.selection_set(self.rows[0]["id"])
            self._on_select()
        self.app.set_status("История отчётов")

    def _selected(self):
        sel = self.tree.selection()
        if not sel:
            return None
        iid = sel[0]
        for row in self.rows:
            if row["id"] == iid:
                return row
        return None

    def _on_select(self, _e=None) -> None:
        row = self._selected()
        if not row:
            return
        if row["cross"]:
            self.cross_lbl.configure(
                text="\n".join(row["cross"]), text_color=COLORS["purple_text"]
            )
        else:
            self.cross_lbl.configure(
                text="В этом отчёте нет броней с переходом квартала.",
                text_color=COLORS["faint"],
            )

    def open_orig(self) -> None:
        row = self._selected()
        if not row:
            messagebox.showinfo("История", "Выберите строку.", parent=self.app)
            return
        self.app.hook_open_file("original", row)

    def open_result(self) -> None:
        row = self._selected()
        if not row:
            messagebox.showinfo("История", "Выберите строку.", parent=self.app)
            return
        self.app.hook_open_file("result", row)

    def delete_row(self) -> None:
        row = self._selected()
        if not row:
            return
        if not messagebox.askyesno(
            "Удалить", f"Удалить запись {row['result']} из истории?", parent=self.app
        ):
            return
        self.rows = [r for r in self.rows if r["id"] != row["id"]]
        self.refresh()
        self.app.toast("Запись удалена из истории (демо GUI)")
