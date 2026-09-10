# -*- coding: utf-8 -*-
"""Мелкие виджеты макета: карточки, чипы, зона файла, тост, стиль Treeview."""

from __future__ import annotations

import os
from typing import Callable, Optional

import customtkinter as ctk
from PIL import Image, ImageDraw

from theme import COLORS, FONTS


def _font(name: str):
    return FONTS.get(name)


def circle_image(color: str, size: int = 10) -> ctk.CTkImage:
    s = size * 2
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((1, 1, s - 2, s - 2), fill=color)
    return ctk.CTkImage(light_image=img, dark_image=img, size=(size, size))


def icon_image(kind: str, color: str = "#C5D3E8", size: int = 16) -> ctk.CTkImage:
    s = size * 2
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    w = 2
    if kind == "history":
        d.rounded_rectangle((4, 8, s - 4, s - 2), radius=4, outline=color, width=w)
        d.line((4, 16, s - 4, 16), fill=color, width=w)
        d.line((10, 2, 10, 10), fill=color, width=w)
        d.line((s - 10, 2, s - 10, 10), fill=color, width=w)
    elif kind == "settings":
        cx = cy = s // 2
        d.ellipse((cx - 5, cy - 5, cx + 5, cy + 5), outline=color, width=w)
        for a in range(0, 360, 45):
            import math

            rad = math.radians(a)
            x1 = cx + 8 * math.cos(rad)
            y1 = cy + 8 * math.sin(rad)
            x2 = cx + 13 * math.cos(rad)
            y2 = cy + 13 * math.sin(rad)
            d.line((x1, y1, x2, y2), fill=color, width=w)
    elif kind == "file":
        d.line((s // 2, 6, s // 2, s - 10), fill=color, width=w)
        d.line((10, 16, s // 2, s - 10), fill=color, width=w)
        d.line((s - 10, 16, s // 2, s - 10), fill=color, width=w)
        d.line((8, s - 4, s - 8, s - 4), fill=color, width=w)
    elif kind == "play":
        d.polygon([(10, 6), (10, s - 6), (s - 6, s // 2)], fill=color)
    else:
        d.ellipse((4, 4, s - 4, s - 4), outline=color, width=w)
    return ctk.CTkImage(light_image=img, dark_image=img, size=(size, size))


def style_treeview() -> None:
    import tkinter.ttk as ttk

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure(
        "Tur.Treeview",
        background="#FFFFFF",
        fieldbackground="#FFFFFF",
        foreground=COLORS["text"],
        rowheight=32,
        font=(FONTS["ui"].cget("family") if FONTS else "Segoe UI", 10),
        borderwidth=0,
        relief="flat",
    )
    style.configure(
        "Tur.Treeview.Heading",
        background="#F7F9FC",
        foreground=COLORS["muted"],
        font=(FONTS["ui"].cget("family") if FONTS else "Segoe UI", 9, "bold"),
        relief="flat",
        padding=4,
    )
    style.map(
        "Tur.Treeview",
        background=[("selected", "#EAF1FF")],
        foreground=[("selected", COLORS["text"])],
    )
    style.layout("Tur.Treeview", style.layout("Treeview"))


class Chip(ctk.CTkFrame):
    def __init__(self, master, text: str, fg: str, tc: str, **kwargs):
        super().__init__(master, fg_color=fg, corner_radius=12, height=24, **kwargs)
        ctk.CTkLabel(
            self, text=text, text_color=tc, font=_font("xs_b"), fg_color="transparent"
        ).pack(padx=9, pady=3)


class Card(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(
            master,
            fg_color=COLORS["card"],
            corner_radius=12,
            border_width=1,
            border_color=COLORS["line"],
            **kwargs,
        )


class StepBadge(ctk.CTkLabel):
    def __init__(self, master, text: str, **kwargs):
        super().__init__(
            master,
            text=text,
            width=26,
            height=26,
            corner_radius=8,
            fg_color=COLORS["sidebar"],
            text_color="#FFFFFF",
            font=_font("xs_b"),
            **kwargs,
        )

    def set_done(self, done: bool) -> None:
        self.configure(fg_color=COLORS["green"] if done else COLORS["sidebar"])


class DropZone(ctk.CTkFrame):
    """Кликабельная зона выбора файла. Путь хранится в .path (str | None)."""

    def __init__(
        self,
        master,
        title: str,
        hint: str,
        on_pick: Callable[["DropZone"], None],
        **kwargs,
    ):
        super().__init__(
            master,
            fg_color=COLORS["drop"],
            corner_radius=10,
            border_width=1,
            border_color=COLORS["drop_border"],
            height=74,
            **kwargs,
        )
        self.on_pick = on_pick
        self.path: Optional[str] = None
        self._hint_text = hint

        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill="both", expand=True, padx=12, pady=12)

        ico = ctk.CTkFrame(
            inner,
            width=40,
            height=40,
            corner_radius=10,
            fg_color="#FFFFFF",
            border_width=1,
            border_color=COLORS["line"],
        )
        ico.pack(side="left")
        ico.pack_propagate(False)
        ctk.CTkLabel(ico, text="xlsx", text_color=COLORS["muted"], font=_font("xs_b")).place(
            relx=0.5, rely=0.5, anchor="center"
        )

        txt = ctk.CTkFrame(inner, fg_color="transparent")
        txt.pack(side="left", fill="both", expand=True, padx=12)
        self.title_lbl = ctk.CTkLabel(
            txt, text=title, anchor="w", font=_font("ui_b"), text_color=COLORS["text"]
        )
        self.title_lbl.pack(anchor="w")
        self.hint_lbl = ctk.CTkLabel(
            txt, text=hint, anchor="w", font=_font("sm"), text_color=COLORS["muted"]
        )
        self.hint_lbl.pack(anchor="w")

        self.act = ctk.CTkButton(
            inner,
            text="Обзор…",
            width=88,
            height=28,
            corner_radius=7,
            fg_color="#FFFFFF",
            hover_color="#F3F7FF",
            text_color=COLORS["accent"],
            border_width=1,
            border_color="#C9D8F0",
            font=_font("sm_b"),
            command=self._pick,
        )
        self.act.pack(side="right")

        self._bind_empty(self)

    def _bind_empty(self, w) -> None:
        if w is self.act:
            return
        w.bind("<Button-1>", lambda _e: self._pick())
        for ch in w.winfo_children():
            self._bind_empty(ch)

    def _pick(self) -> None:
        self.on_pick(self)

    def set_path(self, path: Optional[str]) -> None:
        self.path = path
        if path:
            name = os.path.basename(path)
            self.configure(fg_color=COLORS["green_soft"], border_color="#B7E0C8")
            self.hint_lbl.configure(text=name, text_color=COLORS["text"])
            self.act.configure(text="Заменить", text_color=COLORS["muted"], border_color="#C5D8CC")
        else:
            self.configure(fg_color=COLORS["drop"], border_color=COLORS["drop_border"])
            self.hint_lbl.configure(text=self._hint_text, text_color=COLORS["muted"])
            self.act.configure(text="Обзор…", text_color=COLORS["accent"], border_color="#C9D8F0")

    def clear(self) -> None:
        self.set_path(None)


class StatCard(Card):
    def __init__(self, master, caption: str, value: str, detail: str, **kwargs):
        super().__init__(master, **kwargs)
        pad = ctk.CTkFrame(self, fg_color="transparent")
        pad.pack(fill="both", expand=True, padx=16, pady=14)
        ctk.CTkLabel(
            pad, text=caption, anchor="w", font=_font("xs_b"), text_color=COLORS["muted"]
        ).pack(anchor="w")
        self.value_lbl = ctk.CTkLabel(
            pad, text=value, anchor="w", font=_font("stat"), text_color=COLORS["text"]
        )
        self.value_lbl.pack(anchor="w", pady=(4, 0))
        self.detail_lbl = ctk.CTkLabel(
            pad, text=detail, anchor="w", font=_font("xs"), text_color=COLORS["faint"]
        )
        self.detail_lbl.pack(anchor="w")

    def set_values(self, value: str, detail: str) -> None:
        self.value_lbl.configure(text=value)
        self.detail_lbl.configure(text=detail)


class GhostButton(ctk.CTkButton):
    def __init__(self, master, text: str, command=None, **kwargs):
        kwargs.setdefault("fg_color", "#F4F7FA")
        kwargs.setdefault("hover_color", "#EEF2F6")
        kwargs.setdefault("text_color", COLORS["text"])
        kwargs.setdefault("border_width", 1)
        kwargs.setdefault("border_color", COLORS["line"])
        kwargs.setdefault("corner_radius", 8)
        kwargs.setdefault("height", 32)
        kwargs.setdefault("font", _font("sm_b"))
        super().__init__(master, text=text, command=command, **kwargs)


class PrimaryButton(ctk.CTkButton):
    def __init__(self, master, text: str, command=None, **kwargs):
        kwargs.setdefault("fg_color", COLORS["accent"])
        kwargs.setdefault("hover_color", COLORS["accent_hover"])
        kwargs.setdefault("text_color", "#FFFFFF")
        kwargs.setdefault("corner_radius", 8)
        kwargs.setdefault("height", 32)
        kwargs.setdefault("font", _font("sm_b"))
        super().__init__(master, text=text, command=command, **kwargs)


class GreenButton(ctk.CTkButton):
    def __init__(self, master, text: str, command=None, **kwargs):
        kwargs.setdefault("fg_color", COLORS["green"])
        kwargs.setdefault("hover_color", COLORS["green_hover"])
        kwargs.setdefault("text_color", "#FFFFFF")
        kwargs.setdefault("text_color_disabled", "#F4FFF8")
        kwargs.setdefault("corner_radius", 10)
        kwargs.setdefault("height", 44)
        kwargs.setdefault("font", _font("ui_b"))
        super().__init__(master, text=text, command=command, **kwargs)


class DangerGhost(ctk.CTkButton):
    def __init__(self, master, text: str, command=None, **kwargs):
        kwargs.setdefault("fg_color", COLORS["red_soft"])
        kwargs.setdefault("hover_color", "#F8D4D4")
        kwargs.setdefault("text_color", COLORS["red"])
        kwargs.setdefault("corner_radius", 8)
        kwargs.setdefault("height", 32)
        kwargs.setdefault("font", _font("sm_b"))
        super().__init__(master, text=text, command=command, **kwargs)


def show_toast(anchor: ctk.CTk, text: str, ms: int = 2400) -> None:
    toast = ctk.CTkToplevel(anchor)
    toast.overrideredirect(True)
    toast.attributes("-topmost", True)
    frame = ctk.CTkFrame(toast, fg_color=COLORS["sidebar"], corner_radius=10)
    frame.pack()
    ctk.CTkLabel(
        frame, text=text, text_color="#FFFFFF", font=_font("ui"), wraplength=340, justify="left"
    ).pack(padx=16, pady=12)
    toast.update_idletasks()
    aw = anchor.winfo_width()
    ah = anchor.winfo_height()
    ax = anchor.winfo_rootx()
    ay = anchor.winfo_rooty()
    tw = toast.winfo_reqwidth()
    th = toast.winfo_reqheight()
    toast.geometry(f"+{ax + aw - tw - 24}+{ay + ah - th - 48}")
    toast.after(ms, toast.destroy)


def labeled_entry(parent, label: str, value: str = "", mono: bool = False, **entry_kw):
    box = ctk.CTkFrame(parent, fg_color="transparent")
    ctk.CTkLabel(
        box, text=label, anchor="w", font=_font("xs_b"), text_color=COLORS["muted"]
    ).pack(anchor="w", pady=(0, 4))
    ent = ctk.CTkEntry(
        box,
        height=36,
        corner_radius=8,
        border_color=COLORS["line"],
        fg_color="#FFFFFF",
        text_color=COLORS["text"],
        font=_font("mono") if mono else _font("ui"),
        **entry_kw,
    )
    ent.pack(fill="x")
    if value:
        ent.insert(0, value)
    return box, ent
