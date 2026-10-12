# -*- coding: utf-8 -*-
"""
Турналог — графическая часть редизайна на customtkinter.

Только GUI: навигация, окна, выбор файлов. Расчёт / sqlite / Excel не подключены.
Точки для боевой логики: hook_* внизу класса TurNalogApp.
"""

from __future__ import annotations

from tkinter import messagebox

import customtkinter as ctk

from dialogs import ColumnMapDialog
from login import LoginWindow
from settings_window import SettingsWindow
from theme import COLORS, FONTS, HOTEL_DOT, HOTELS, DB_PATH_DEMO, setup_fonts
from views import HistoryView, ProcessView
from widgets import circle_image, icon_image, show_toast


class TurNalogApp(ctk.CTk):
    def __init__(self):
        super().__init__(fg_color=COLORS["bg"])
        setup_fonts()
        self.withdraw()
        self.user = None
        self.current_hotel = "chai"
        self._settings = None
        self.nav_items = {}
        LoginWindow(self)

    def start_main(self, user: dict) -> None:
        """После успешного входа — собрать главное окно."""
        self.user = user
        self.title("Турналог — обработка данных гостиниц")
        self.minsize(1100, 700)
        self.geometry("1240x780")
        self._center()
        self._build_sidebar()
        self._build_status()
        self._build_main()
        first = user.get("hotel_id") or "chai"
        self.show_process(first)
        self.set_status("Готово к работе")
        self.deiconify()
        self.lift()
        self.focus_force()

    def _center(self) -> None:
        self.update_idletasks()
        w, h = 1240, 780
        x = max(0, (self.winfo_screenwidth() - w) // 2)
        y = max(0, (self.winfo_screenheight() - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ----- сайдбар -----
    def _build_sidebar(self) -> None:
        self.sidebar = ctk.CTkFrame(
            self, width=232, corner_radius=0, fg_color=COLORS["sidebar"]
        )
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        brand = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        brand.pack(fill="x", padx=20, pady=(22, 8))
        ctk.CTkLabel(
            brand,
            text="ТУРНАЛОГ",
            anchor="w",
            font=FONTS["brand"],
            text_color="#FFFFFF",
        ).pack(anchor="w")
        ctk.CTkLabel(
            brand,
            text="налог по выгрузкам отелей",
            anchor="w",
            font=FONTS["xs"],
            text_color="#9AADC8",
        ).pack(anchor="w", pady=(2, 0))

        foot = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        foot.pack(side="bottom", fill="x", padx=14, pady=16)
        av = ctk.CTkLabel(
            foot,
            text="A",
            width=32,
            height=32,
            corner_radius=16,
            fg_color="#3D5A80",
            text_color="#FFFFFF",
            font=FONTS["sm_b"],
        )
        av.pack(side="left")
        meta = ctk.CTkFrame(foot, fg_color="transparent")
        meta.pack(side="left", padx=10)
        login = (self.user or {}).get("login", "admin")
        role = (self.user or {}).get("role", "администратор")
        ctk.CTkLabel(
            meta, text=login, anchor="w", font=FONTS["sm_b"], text_color="#FFFFFF"
        ).pack(anchor="w")
        ctk.CTkLabel(
            meta,
            text=f"{role} · GUI",
            anchor="w",
            font=FONTS["xs"],
            text_color="#8FA3C0",
        ).pack(anchor="w")

        ctk.CTkLabel(
            self.sidebar,
            text="ОТЕЛИ",
            anchor="w",
            font=FONTS["xs_b"],
            text_color=COLORS["sidebar_muted"],
        ).pack(fill="x", padx=20, pady=(16, 6))

        allowed = (self.user or {}).get("hotel_id")
        for hid, h in HOTELS.items():
            if allowed and hid != allowed:
                continue
            self._nav_hotel(hid, h["name"], HOTEL_DOT[hid])

        sep = ctk.CTkFrame(self.sidebar, height=1, fg_color="#2C3E5C")
        sep.pack(fill="x", padx=20, pady=12)

        self._nav_action(
            "history",
            "История",
            icon_image("history", COLORS["sidebar_text"]),
            self.show_history,
        )
        if (self.user or {}).get("role") == "администратор":
            self._nav_action(
                "settings",
                "Настройки",
                icon_image("settings", COLORS["sidebar_text"]),
                self.open_settings,
            )

    def _nav_hotel(self, hid: str, name: str, color: str) -> None:
        btn = ctk.CTkButton(
            self.sidebar,
            text="  " + name,
            image=circle_image(color, 8),
            compound="left",
            anchor="w",
            height=36,
            corner_radius=8,
            fg_color="transparent",
            hover_color=COLORS["sidebar_hover"],
            text_color=COLORS["sidebar_text"],
            font=FONTS["ui"],
            command=lambda: self.show_process(hid),
        )
        btn.pack(fill="x", padx=10, pady=1)
        self.nav_items[hid] = btn

    def _nav_action(self, key: str, text: str, image, command) -> None:
        btn = ctk.CTkButton(
            self.sidebar,
            text="  " + text,
            image=image,
            compound="left",
            anchor="w",
            height=36,
            corner_radius=8,
            fg_color="transparent",
            hover_color=COLORS["sidebar_hover"],
            text_color=COLORS["sidebar_text"],
            font=FONTS["ui"],
            command=command,
        )
        btn.pack(fill="x", padx=10, pady=1)
        self.nav_items[key] = btn

    def _set_active_nav(self, key: str) -> None:
        for k, btn in self.nav_items.items():
            if k == key:
                btn.configure(
                    fg_color=COLORS["sidebar_act"],
                    text_color="#FFFFFF",
                    font=FONTS["ui_b"],
                )
            else:
                btn.configure(
                    fg_color="transparent",
                    text_color=COLORS["sidebar_text"],
                    font=FONTS["ui"],
                )

    # ----- статус + контент -----
    def _build_status(self) -> None:
        self.status = ctk.CTkFrame(
            self, height=28, corner_radius=0, fg_color=COLORS["status_bg"]
        )
        self.status.pack(side="bottom", fill="x")
        self.status.pack_propagate(False)
        dot = ctk.CTkLabel(
            self.status, text="●", text_color="#22A06B", font=FONTS["xs"], width=18
        )
        dot.pack(side="left", padx=(10, 0))
        self.status_lbl = ctk.CTkLabel(
            self.status,
            text="Готово к работе",
            anchor="w",
            font=FONTS["xs"],
            text_color="#5A6778",
        )
        self.status_lbl.pack(side="left")
        self.db_lbl = ctk.CTkLabel(
            self.status,
            text="БД · " + DB_PATH_DEMO,
            anchor="e",
            font=FONTS["xs"],
            text_color="#5A6778",
        )
        self.db_lbl.pack(side="right", padx=12)

    def _build_main(self) -> None:
        self.main = ctk.CTkFrame(self, fg_color=COLORS["bg"], corner_radius=0)
        self.main.pack(side="left", fill="both", expand=True)

        self.process_view = ProcessView(self.main, self)
        self.history_view = HistoryView(self.main, self)

    def set_status(self, text: str) -> None:
        self.status_lbl.configure(text=text)

    def toast(self, text: str) -> None:
        show_toast(self, text)

    def show_process(self, hotel_id: str) -> None:
        self.current_hotel = hotel_id
        self._set_active_nav(hotel_id)
        self.history_view.pack_forget()
        self.process_view.pack(fill="both", expand=True, padx=26, pady=22)
        self.process_view.bind_hotel(hotel_id)
        self.set_status("Готово к работе")

    def show_history(self) -> None:
        self._set_active_nav("history")
        self.process_view.pack_forget()
        self.history_view.pack(fill="both", expand=True, padx=26, pady=22)
        self.history_view.refresh()

    def open_settings(self) -> None:
        if self._settings is not None and self._settings.winfo_exists():
            self._settings.lift()
            self._settings.focus()
            return
        self._settings = SettingsWindow(self)

    def open_column_map(self, kind: str) -> None:
        ColumnMapDialog(self, kind)

    # ------------------------------------------------------------------
    # Точки подключения боевой логики. Сейчас — заглушки.
    # Не удалять сигнатуры: сюда вставлять pandas / openpyxl / sqlite.
    # ------------------------------------------------------------------
    def hook_process(
        self,
        hotel_id: str,
        data_path: str,
        rooms_path: str | None,
        quarter: str,
        year: str,
        check_dups: bool,
        dest_path: str,
    ) -> None:
        self.set_status("Обработка… (GUI-заглушка)")
        messagebox.showinfo(
            "Турналог — макет GUI",
            "Обработка не подключена.\n\n"
            f"Отель: {HOTELS[hotel_id]['name']}\n"
            f"Данные: {data_path}\n"
            f"Номера: {rooms_path or '—'}\n"
            f"Квартал: {quarter} {year}\n"
            f"Дубли: {'да' if check_dups else 'нет'}\n"
            f"Сохранить как: {dest_path}\n\n"
            "Подставьте сюда текущий пайплайн из 2.6 (TurNalogApp.hook_process).",
            parent=self,
        )
        self.set_status("Готово к работе")
        self.toast("Файл не сформирован — это только GUI")

    def hook_tl_import(self, hotel_id: str) -> None:
        messagebox.showinfo(
            "TravelLine",
            "Импорт из TravelLine доступен только в боевом приложении.",
            parent=self,
        )

    def hook_open_file(self, kind: str, row: dict) -> None:
        key = "orig" if kind == "original" else "result"
        messagebox.showinfo(
            "Открыть файл",
            f"Макет GUI. Здесь os.startfile / subprocess файла:\n{row.get(key)}",
            parent=self,
        )

    def hook_save_settings(self, payload: dict) -> None:
        self.toast("Настройки сохранены (демо GUI)")
        self.set_status("Настройки записаны в память макета")

    def hook_save_columns(self, kind: str, mapping: dict) -> None:
        # В бою — сериализация в json/sqlite по отелю.
        _ = (kind, mapping)

    def hook_authenticate(self, login: str, password: str) -> dict | None:
        """Сюда проверка sqlite/хеша из 2.6. Сейчас: admin/admin."""
        if login == "admin" and password == "admin":
            return {"login": "admin", "role": "администратор", "hotel_id": None}
        if login == "chai_buh" and password == "chai":
            return {"login": "chai_buh", "role": "отель", "hotel_id": "chai"}
        if login == "dona_buh" and password == "dona":
            return {"login": "dona_buh", "role": "отель", "hotel_id": "dona"}
        return None


def main() -> None:
    ctk.set_appearance_mode("light")
    try:
        ctk.set_default_color_theme("blue")
    except Exception:
        pass
    app = TurNalogApp()
    app.mainloop()


if __name__ == "__main__":
    main()
