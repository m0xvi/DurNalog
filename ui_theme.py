# -*- coding: utf-8 -*-
"""Дизайн-система Турналог 3.0 (по макету дизайнера).
Единственное место, где живут цвета/шрифты. Светлая тема фиксирована."""
import customtkinter as ctk
from tkinter import ttk

COLORS = {
    "bg":           "#E7EDF5",
    "sidebar":      "#1B2A4A",
    "sidebar_hover": "#243656",
    "sidebar_act":  "#31476A",
    "card":         "#FFFFFF",
    "text":         "#1A2332",
    "muted":        "#5C6B80",
    "faint":        "#8A97A8",
    "line":         "#D7E0EA",
    "line2":        "#EEF2F6",
    "accent":       "#2563EB",
    "accent_hover": "#1D4ED8",
    "accent_soft":  "#E8F0FE",
    "green":        "#0B8F55",
    "green_hover":  "#097A48",
    "green_soft":   "#E5F7EE",
    "green_border": "#B7E0C8",
    "green_dis":    "#A9CABB",
    "red":          "#DC3D43",
    "red_soft":     "#FDECEC",
    "amber":        "#C47B12",
    "amber_soft":   "#FFF4E0",
    "purple":       "#6D28D9",
    "purple_soft":  "#F3E8FF",
    "purple_text":  "#4C1D95",
    "zone":         "#F7FAFC",
    "zone_border":  "#C5D2E0",
}

HOTEL_NAME = {'tch': 'Чайковский', 'dna': 'Дона', 'usd': 'Усадьба'}
HOTEL_DOT = {'tch': '#2BB3C8', 'dna': '#E2A03A', 'usd': '#D46A8A'}
HOTEL_PILL = {'tch': ('#E0F4FA', '#0E7490'),
              'dna': ('#FFF4E0', '#C47B12'),
              'usd': ('#FCE8EF', '#B54A6A')}

F_UI = ("Segoe UI", 13)
F_SMALL = ("Segoe UI", 11)
F_TINY = ("Segoe UI", 10)
F_PAGE = ("Segoe UI", 20, "bold")
F_CARD = ("Segoe UI", 14, "bold")
F_MONO = ("Consolas", 12)


def card(parent):
    """Белая карточка с бордером — базовый блок контента."""
    return ctk.CTkFrame(parent, fg_color=COLORS["card"], corner_radius=12,
                        border_width=1, border_color=COLORS["line"])


def apply_treeview_style():
    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Treeview", background="#FFFFFF", fieldbackground="#FFFFFF",
                    foreground=COLORS["text"], rowheight=30,
                    font=("Segoe UI", 10), borderwidth=0)
    style.configure("Treeview.Heading", background="#F7F9FC",
                    foreground=COLORS["muted"], font=("Segoe UI", 9, "bold"),
                    relief="flat")
    style.map("Treeview", background=[("selected", "#EAF1FF")],
              foreground=[("selected", COLORS["text"])])


def toast(master, text):
    """Маленький тост в правом нижнем углу, 2.5 с."""
    try:
        t = ctk.CTkToplevel(master)
        t.overrideredirect(True)
        t.attributes("-topmost", True)
        fr = ctk.CTkFrame(t, fg_color=COLORS["sidebar"], corner_radius=10)
        fr.pack()
        ctk.CTkLabel(fr, text="  ✓  " + text + "  ", text_color="white",
                     font=F_SMALL, height=38).pack()
        t.update_idletasks()
        x = master.winfo_rootx() + master.winfo_width() - t.winfo_width() - 24
        y = master.winfo_rooty() + master.winfo_height() - t.winfo_height() - 40
        t.geometry(f"+{max(0, x)}+{max(0, y)}")
        t.after(2500, t.destroy)
    except Exception:
        pass
