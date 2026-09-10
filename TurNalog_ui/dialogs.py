# -*- coding: utf-8 -*-
"""Диалоги GUI: сопоставление колонок. Поля — заглушка, подменить из боевого кода."""

from __future__ import annotations

from typing import TYPE_CHECKING, List, Tuple

import customtkinter as ctk

from theme import COLORS, COLUMN_MAP_DATA, COLUMN_MAP_ROOMS, FONTS
from widgets import GhostButton, GreenButton, show_toast

if TYPE_CHECKING:
    from app import TurNalogApp


def _f(name: str):
    return FONTS.get(name)


class ColumnMapDialog(ctk.CTkToplevel):
    def __init__(self, app: "TurNalogApp", kind: str):
        super().__init__(app)
        self.app = app
        self.kind = kind
        title = (
            "Сопоставление колонок «Данные»"
            if kind == "data"
            else "Сопоставление колонок «Номера»"
        )
        self.title(title)
        self.geometry("560x480")
        self.resizable(False, True)
        self.configure(fg_color=COLORS["bg"])
        self.transient(app)
        self.grab_set()

        fields: List[Tuple[str, str]] = COLUMN_MAP_DATA if kind == "data" else COLUMN_MAP_ROOMS

        ctk.CTkLabel(
            self, text=title, font=_f("h2"), text_color=COLORS["text"], anchor="w"
        ).pack(fill="x", padx=20, pady=(16, 4))
        ctk.CTkLabel(
            self,
            text="Укажите, в каких колонках файла лежат поля расчёта. Сохраняется для этого отеля.\n"
            "Список полей ниже — из макета. В бою подставить текущий диалог 2.6.",
            font=_f("sm"),
            text_color=COLORS["muted"],
            justify="left",
            anchor="w",
        ).pack(fill="x", padx=20, pady=(0, 10))

        body = ctk.CTkScrollableFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=20)
        self.entries = []
        for label, value in fields:
            row = ctk.CTkFrame(body, fg_color="transparent")
            row.pack(fill="x", pady=4)
            ctk.CTkLabel(
                row, text=label, width=160, anchor="w", font=_f("ui"), text_color=COLORS["text"]
            ).pack(side="left")
            ent = ctk.CTkEntry(
                row,
                height=34,
                corner_radius=8,
                border_color=COLORS["line"],
                fg_color="#FFFFFF",
                font=_f("ui"),
            )
            ent.pack(side="left", fill="x", expand=True)
            ent.insert(0, value)
            self.entries.append((label, ent))

        foot = ctk.CTkFrame(self, fg_color="#FFFFFF", corner_radius=0)
        foot.pack(fill="x", side="bottom")
        inner = ctk.CTkFrame(foot, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=12)
        GreenButton(inner, "Сохранить", command=self._save, width=120, height=36).pack(
            side="right", padx=(8, 0)
        )
        GhostButton(inner, "Отмена", command=self.destroy, width=100, height=36).pack(side="right")

        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.after(50, self._center)

    def _center(self) -> None:
        self.update_idletasks()
        px, py = self.app.winfo_rootx(), self.app.winfo_rooty()
        pw, ph = self.app.winfo_width(), self.app.winfo_height()
        w, h = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{px + (pw - w) // 2}+{py + (ph - h) // 2}")

    def _save(self) -> None:
        mapping = {label: ent.get() for label, ent in self.entries}
        self.app.hook_save_columns(self.kind, mapping)
        show_toast(self.app, "Сопоставление колонок сохранено (демо GUI)")
        self.destroy()
