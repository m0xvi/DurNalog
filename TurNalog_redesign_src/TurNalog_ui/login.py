# -*- coding: utf-8 -*-
"""Окно входа. Поля как в 2.6; внешний вид — редизайн. Проверка пароля — заглушка."""

from __future__ import annotations

from typing import TYPE_CHECKING

import customtkinter as ctk
from tkinter import messagebox

from theme import COLORS, DB_PATH_DEMO, FONTS
from widgets import GreenButton

if TYPE_CHECKING:
    from app import TurNalogApp


def _f(name: str):
    return FONTS.get(name)


class LoginWindow(ctk.CTkToplevel):
    def __init__(self, app: "TurNalogApp"):
        super().__init__(app)
        self.app = app
        self.title("Вход")
        self.resizable(False, False)
        self.configure(fg_color=COLORS["bg"])
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        w, h = 440, 500
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")

        card = ctk.CTkFrame(self, fg_color="transparent")
        card.pack(fill="both", expand=True, padx=40, pady=28)

        ctk.CTkLabel(
            card,
            text="ТУРНАЛОГ",
            font=_f("h1"),
            text_color=COLORS["text"],
        ).pack(pady=(12, 0))
        ctk.CTkLabel(
            card,
            text="Вход в систему",
            font=_f("sm"),
            text_color=COLORS["muted"],
        ).pack(pady=(4, 22))

        self.login_ent = ctk.CTkEntry(
            card,
            height=42,
            corner_radius=10,
            border_color=COLORS["line"],
            fg_color="#FFFFFF",
            text_color=COLORS["text"],
            font=_f("ui"),
            placeholder_text="Логин",
        )
        self.login_ent.pack(fill="x", pady=(0, 10))

        self.pass_ent = ctk.CTkEntry(
            card,
            height=42,
            corner_radius=10,
            border_color=COLORS["line"],
            fg_color="#FFFFFF",
            text_color=COLORS["text"],
            font=_f("ui"),
            placeholder_text="Пароль",
            show="•",
        )
        self.pass_ent.pack(fill="x", pady=(0, 16))

        GreenButton(card, "Войти", command=self._submit, height=42).pack(fill="x")

        ctk.CTkLabel(
            card,
            text="БД: " + DB_PATH_DEMO,
            font=_f("xs"),
            text_color=COLORS["faint"],
        ).pack(pady=(18, 0))

        ctk.CTkLabel(
            card,
            text="Первый запуск: login admin, пароль admin — смените пароль в «Пользователи»",
            font=_f("xs"),
            text_color=COLORS["muted"],
            wraplength=360,
            justify="center",
        ).pack(pady=(8, 0))

        self.login_ent.bind("<Return>", lambda _e: self._submit())
        self.pass_ent.bind("<Return>", lambda _e: self._submit())
        self.after(80, self._ready)

    def _ready(self) -> None:
        try:
            self.grab_set()
        except Exception:
            pass
        self.lift()
        self.focus_force()
        self.login_ent.focus_set()

    def _submit(self) -> None:
        login = self.login_ent.get().strip()
        password = self.pass_ent.get()
        if not login or not password:
            messagebox.showwarning("Вход", "Введите логин и пароль.", parent=self)
            return
        user = self.app.hook_authenticate(login, password)
        if not user:
            messagebox.showerror("Вход", "Неверный логин или пароль.", parent=self)
            self.pass_ent.delete(0, "end")
            self.pass_ent.focus_set()
            return
        self.grab_release()
        self.destroy()
        self.app.start_main(user)

    def _on_close(self) -> None:
        self.app.destroy()
