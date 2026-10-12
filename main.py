# -*- coding: utf-8 -*-
"""Турналог 3.0 — дизайнерский GUI (TurNalog_ui) + боевая логика (core/db)."""
import os
import sys
import tempfile
import threading
import copy
from datetime import datetime

APP_DIR = os.path.dirname(os.path.abspath(__file__))
if getattr(sys, 'frozen', False):
    APP_DIR = os.path.dirname(sys.executable)
_GUI_DIR = os.path.join(APP_DIR, 'TurNalog_ui')
if not os.path.isdir(_GUI_DIR) and getattr(sys, 'frozen', False):
    # в собранном EXE папка лежит внутри _internal
    _GUI_DIR = os.path.join(getattr(sys, '_MEIPASS', APP_DIR), 'TurNalog_ui')
sys.path.insert(0, _GUI_DIR)

import logging
logging.basicConfig(filename=os.path.join(APP_DIR, 'turnalog.log'),
                    level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger('turnalog')

import customtkinter as ctk
from tkinter import messagebox, filedialog
import tkinter as tk

from app import TurNalogApp
import theme
from settings_window import (SettingsWindow, AgentsPane, UsersPane, DataPane)
from views import ProcessView, HistoryView

from agents import migrate_agents, match_agent
from tax import TaxParams
from db import HistoryDB, set_db_path, default_db_path
from schemas import SCHEMAS, needs_mapping
from tl_client import TL_DATA_MAP, TL_EXTRA, TL_ROOMS_MAP, is_tl_export
from core import (read_table, read_headers, guess_mapping, mapping_valid,
                  pack_originals, open_in_system, process_tchaikovsky,
                  process_dona, process_usadba, usadba_sources,
                  repair_notes_for)
from dialogs_real import MappingDialog, UnrecognizedSourcesDialog, TlConfigDialog
from updater import (UpdateError, check_for_updates, download_release,
                     is_frozen, launch_install, plan_install, staging_dir,
                     verify_package)
from ui_theme import COLORS as UI  # палитра боевых диалогов

APP_VERSION = "3.16"

ID2CODE = {'chai': 'tch', 'dona': 'dna', 'usadba': 'usd'}
CODE2ID = {'tch': 'chai', 'dna': 'dona', 'usadba': 'usadba'}
NAME2ID = {'Чайковский': 'chai', 'Дона': 'dona', 'Усадьба': 'usadba'}
ROMAN = {'I': 'Q1', 'II': 'Q2', 'III': 'Q3', 'IV': 'Q4'}
ROMAN_REV = {'Q1': 'I', 'Q2': 'II', 'Q3': 'III', 'Q4': 'IV'}

from config_io import load_config, save_config  # noqa: E402


# ================================================================
# Реальные панели настроек (данные из БД/конфига вместо демо)
# ================================================================
class RealAgentsPane(AgentsPane):
    def __init__(self, master, app):
        super().__init__(master, app)
        self.data = migrate_agents(app.cfg)
        self.agents = [a['name'] for a in self.data]
        self.current = self.agents[0] if self.agents else None
        # контейнер доп. периодов — внутри правой формы
        split = [c for c in self.winfo_children() if isinstance(c, ctk.CTkFrame)][-1]
        right = split.winfo_children()[1]
        form = right.winfo_children()[0]
        self.extra_pf = ctk.CTkFrame(form, fg_color="transparent")
        self.extra_pf.pack(fill="x", padx=0, pady=(4, 0))
        self._extra_rows = []
        self._rebuild_list()
        self._select_first()

    def _select_first(self):
        if self.current:
            self._select(self.current)

    def _by_name(self, name):
        for a in getattr(self, 'data', None) or []:
            if a['name'] == name:
                return a
        return None

    def _load_current(self):
        if not getattr(self, 'data', None):
            return  # базовый __init__ вызвал метод раньше, чем мы заполнили data
        a = self._by_name(self.current) if self.current else None
        self.name_ent.delete(0, "end")
        if a:
            self.name_ent.insert(0, a['name'])
            self.syn_ent.delete(0, "end")
            self.syn_ent.insert(0, ", ".join(a.get('aliases', [])))
            for hid, var in self.hotel_vars.items():
                var.set(CODE2ID.get(hid and ID2CODE.get(hid, hid), hid) in a.get('hotels', [])
                        if False else (ID2CODE.get(hid, hid) in a.get('hotels', [])))
            rates = a.get('rates') or [{'from': '', 'to': '', 'pct': 0}]
            r0 = rates[0]
            for ent, key in ((self.p_from, 'from'), (self.p_to, 'to'), (self.p_pct, 'pct')):
                ent.delete(0, "end")
                ent.insert(0, str(r0.get(key, '')))
            for w in getattr(self, '_extra_rows', []):
                w[3].destroy()
            self._extra_rows = []
            for r in rates[1:]:
                self._add_period_row(r)

    def _add_period_row(self, r=None):
        r = r or {'from': '', 'to': '', 'pct': ''}
        fr = ctk.CTkFrame(self.extra_pf, fg_color="transparent")
        fr.pack(fill="x", pady=2)
        fv = ctk.CTkEntry(fr, placeholder_text="с", width=120, height=32,
                          border_color=UI['line'])
        fv.pack(side="left", padx=(0, 6))
        fv.insert(0, r.get('from', '') or '')
        tv = ctk.CTkEntry(fr, placeholder_text="по", width=120, height=32,
                          border_color=UI['line'])
        tv.pack(side="left", padx=(0, 6))
        tv.insert(0, r.get('to', '') or '')
        pv = ctk.CTkEntry(fr, width=64, height=32, justify="center",
                          border_color=UI['line'])
        pv.pack(side="left")
        pv.insert(0, str(r.get('pct', '')))
        ctk.CTkLabel(fr, text="%", font=theme.FONTS.get("ui")).pack(side="left")
        self._extra_rows.append((fv, tv, pv, fr))

    def _flush_current(self):
        a = self._by_name(self.current) if self.current else None
        if not a:
            return
        new_name = self.name_ent.get().strip() or a['name']
        a['name'] = new_name
        a['aliases'] = [x.strip() for x in self.syn_ent.get().split(',') if x.strip()]
        a['hotels'] = [ID2CODE[h] for h, v in self.hotel_vars.items() if v.get()] or ['dna']
        rates = [{'from': self.p_from.get().strip(), 'to': self.p_to.get().strip(),
                  'pct': self._pct(self.p_pct.get())}]
        for fv, tv, pv, fr in getattr(self, '_extra_rows', []):
            rates.append({'from': fv.get().strip(), 'to': tv.get().strip(),
                          'pct': self._pct(pv.get())})
        a['rates'] = rates
        self.agents = [x['name'] for x in self.data]

    @staticmethod
    def _pct(s):
        try:
            return float(s.strip().replace(',', '.'))
        except ValueError:
            return 0.0

    def _select(self, name):
        self._flush_current()
        self.current = name if name in [a['name'] for a in self.data] else name
        self._rebuild_list()
        self._load_current()

    def _add(self):
        self._flush_current()
        dlg = ctk.CTkInputDialog(text="Название агента:", title="Новый агент")
        name = (dlg.get_input() or '').strip()
        if not name:
            return
        self.data.append({'name': name, 'aliases': [], 'hotels': ['dna'],
                          'rates': [{'from': '', 'to': '', 'pct': 0}]})
        self.agents = [a['name'] for a in self.data]
        self.current = name
        self._rebuild_list()
        self._load_current()

    def _delete(self):
        if not self.current:
            return
        if not messagebox.askyesno("Удалить", f"Удалить агента «{self.current}»?",
                                   parent=self.winfo_toplevel()):
            return
        self.data = [a for a in self.data if a['name'] != self.current]
        self.agents = [a['name'] for a in self.data]
        self.current = self.agents[0] if self.agents else None
        self._rebuild_list()
        self._load_current()

    def _stub(self, title):
        top = self.winfo_toplevel()
        if title == "Импорт агентов":
            path = filedialog.askopenfilename(parent=top, title="Импорт агентов",
                                              filetypes=[("Таблицы", "*.xlsx *.csv")])
            if not path:
                return
            try:
                from agents import import_agents
                self._flush_current()
                imp = import_agents(path)
                by = {a['name']: a for a in self.data}
                add = rep = 0
                for a in imp:
                    if a['name'] in by:
                        by[a['name']].update(a); rep += 1
                    else:
                        self.data.append(a); add += 1
                self.agents = [a['name'] for a in self.data]
                self._rebuild_list()
                messagebox.showinfo("Импорт", f"Добавлено: {add}, заменено: {rep}", parent=top)
            except Exception as e:
                messagebox.showerror("Импорт", f"Не удалось импортировать:\n{e}", parent=top)
        elif title == "Экспорт агентов":
            path = filedialog.asksaveasfilename(parent=top, title="Экспорт агентов",
                                                defaultextension=".xlsx",
                                                filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
            if not path:
                return
            try:
                from agents import export_agents
                self._flush_current()
                export_agents(self.data, path)
                messagebox.showinfo("Экспорт", f"Сохранено:\n{path}", parent=top)
            except Exception as e:
                messagebox.showerror("Экспорт", f"Не удалось сохранить:\n{e}", parent=top)
        elif title == "Проверка синонимов":
            dlg2 = ctk.CTkToplevel(top)
            dlg2.title("Проверка распознавания")
            dlg2.geometry("460x150")
            dlg2.transient(top)
            ev = tk.StringVar()
            ctk.CTkEntry(dlg2, textvariable=ev, height=34,
                         placeholder_text="вставьте «Источник» из отчёта…").pack(
                padx=20, pady=(16, 8), fill="x")
            res = ctk.CTkLabel(dlg2, text="", font=UI and ("Segoe UI", 12, "bold"))
            res.pack(padx=20)

            def go():
                self._flush_current()
                c, n = match_agent(ev.get(), self.data)
                res.configure(text=f"→ {n}, {c:g}%" if n else "→ не распознан (0%)")
            ctk.CTkButton(dlg2, text="Проверить", command=go, height=30,
                          fg_color=UI['accent']).pack(pady=(0, 12))
        elif title == "Добавить период":
            self._add_period_row()

    def collect_agents(self):
        self._flush_current()
        return copy.deepcopy(self.data)


class RealUsersPane(UsersPane):
    def _reload(self):
        self.tree.delete(*self.tree.get_children())
        with HistoryDB() as db:
            self._db_users = db.list_users()
        for u in self._db_users:
            self.tree.insert("", "end", iid=str(u['id']), values=(
                u['login'],
                "администратор" if u['role'] == 'admin' else "пользователь",
                theme.HOTELS[CODE2ID[u['hotel']]]['name'] if u['hotel'] else "—"))
        if self._db_users:
            self.tree.selection_set(str(self._db_users[0]['id']))

    def _selected(self):
        sel = self.tree.selection()
        if not sel:
            return None
        for u in self._db_users:
            if str(u['id']) == sel[0]:
                return u
        return None

    def _add(self):
        top = self.winfo_toplevel()
        dlg = ctk.CTkToplevel(top)
        dlg.title("Новый пользователь")
        dlg.geometry("340x300")
        dlg.transient(top)
        lv, pv = tk.StringVar(), tk.StringVar()
        ctk.CTkLabel(dlg, text="Логин").pack(padx=40, anchor="w", pady=(14, 0))
        ctk.CTkEntry(dlg, textvariable=lv, height=32).pack(padx=40, fill="x")
        ctk.CTkLabel(dlg, text="Пароль").pack(padx=40, anchor="w", pady=(8, 0))
        ctk.CTkEntry(dlg, textvariable=pv, height=32, show="•").pack(padx=40, fill="x")
        ctk.CTkLabel(dlg, text="Роль").pack(padx=40, anchor="w", pady=(8, 0))
        rv = tk.StringVar(value="пользователь отеля")
        ctk.CTkOptionMenu(dlg, variable=rv, height=32,
                          values=["пользователь отеля", "администратор"]).pack(padx=40, fill="x")
        ctk.CTkLabel(dlg, text="Отель").pack(padx=40, anchor="w", pady=(8, 0))
        hv = tk.StringVar(value="Чайковский")
        ctk.CTkOptionMenu(dlg, variable=hv, height=32,
                          values=[h['name'] for h in theme.HOTELS.values()]).pack(padx=40, fill="x")

        def ok():
            role = 'admin' if rv.get() == 'администратор' else 'hotel'
            hotel = None if role == 'admin' else ID2CODE[NAME2ID[hv.get()]]
            with HistoryDB() as db:
                if not db.add_user(lv.get().strip(), pv.get(), role, hotel):
                    messagebox.showerror("Пользователи", "Логин занят или поля пустые.",
                                         parent=dlg)
                    return
            self._reload()
            dlg.destroy()

        ctk.CTkButton(dlg, text="Создать", fg_color=UI['green'], height=32,
                      command=ok).pack(pady=14, padx=40, fill="x")

    def _password(self):
        u = self._selected()
        if not u:
            return
        dlg = ctk.CTkInputDialog(text=f"Новый пароль для {u['login']}:", title="Сменить пароль")
        pw = dlg.get_input()
        if pw:
            with HistoryDB() as db:
                db.set_password(u['id'], pw)
            self.app.toast("Пароль изменён")

    def _delete(self):
        u = self._selected()
        if not u:
            return
        if not messagebox.askyesno("Удалить", f"Удалить {u['login']}?",
                                   parent=self.winfo_toplevel()):
            return
        with HistoryDB() as db:
            if not db.del_user(u['id']):
                messagebox.showwarning("Удалить", "Нельзя удалить последнего администратора.",
                                       parent=self.winfo_toplevel())
                return
        self._reload()


class RealDataPane(DataPane):
    def __init__(self, master, app):
        super().__init__(master, app)
        self.path_ent.configure(state="normal")
        self.path_ent.delete(0, "end")
        self.path_ent.insert(0, default_db_path())
        self.path_ent.configure(state="disabled")

    def _apply(self, new_path):
        cur = default_db_path()
        if os.path.normcase(new_path) == os.path.normcase(cur):
            return
        if not os.path.exists(new_path) and os.path.exists(cur):
            if messagebox.askyesno("Копирование",
                                   "Скопировать текущую базу (историю и пользователей)\n"
                                   "в новую папку?", parent=self.winfo_toplevel()):
                import shutil
                try:
                    shutil.copy(cur, new_path)
                except Exception as e:
                    messagebox.showerror("База данных", f"Не удалось скопировать:\n{e}",
                                         parent=self.winfo_toplevel())
                    return
        try:
            t = HistoryDB(new_path)
            t.ensure_default_admin()
            t.close()
        except Exception as e:
            messagebox.showerror("База данных", f"Не удалось открыть/создать базу:\n{new_path}\n{e}",
                                 parent=self.winfo_toplevel())
            return
        set_db_path(new_path)
        self.app.cfg['db_path'] = new_path
        save_config(self.app.cfg)
        self.path_ent.configure(state="normal")
        self.path_ent.delete(0, "end")
        self.path_ent.insert(0, new_path)
        self.path_ent.configure(state="disabled")
        self.app.db_lbl.configure(text="БД · " + new_path)
        self.app.toast("Путь БД изменён — применяется сразу")

    def _change(self):
        d = filedialog.askdirectory(parent=self.winfo_toplevel(), title="Папка базы данных")
        if d:
            self._apply(os.path.join(d, 'turnalog_history.db'))

    def _reset(self):
        self.app.cfg.pop('db_path', None)
        save_config(self.app.cfg)
        set_db_path(None)
        self.path_ent.configure(state="normal")
        self.path_ent.delete(0, "end")
        self.path_ent.insert(0, default_db_path())
        self.path_ent.configure(state="disabled")
        self.app.db_lbl.configure(text="БД · " + default_db_path())


class RealSettingsWindow(SettingsWindow):
    def __init__(self, app):
        super().__init__(app)
        for key, cls in (('agents', RealAgentsPane), ('users', RealUsersPane),
                         ('data', RealDataPane)):
            old = self.panes[key]
            was_shown = old.winfo_manager()
            old.destroy()
            self.panes[key] = cls(self.host, app)
            if was_shown:
                self._show(key)

    def _save(self):
        payload = {
            "chai": self.panes["chai"].collect(),
            "dona": self.panes["dona"].collect(),
            "usadba": self.panes["usadba"].collect(),
            "db": self.panes["data"].path_ent.get(),
            "agents": self.panes["agents"].collect_agents(),
        }
        self.app.hook_save_settings(payload)
        self.destroy()


class RealHistoryView(HistoryView):
    def refresh(self):
        self.rows = self.app.history_rows()
        super().refresh()

    def delete_row(self):
        row = self._selected()
        if not row:
            return
        if not messagebox.askyesno("Удалить", f"Удалить запись {row['result']} из истории?",
                                   parent=self.app):
            return
        with HistoryDB() as db:
            db.delete_report(int(row['id']))
        self.refresh()
        self.app.toast("Запись удалена из истории")


class RealProcessView(ProcessView):
    def bind_hotel(self, hotel_id):
        super().bind_hotel(hotel_id)
        rep = self.app.last_report(hotel_id)
        if rep:
            y, qn = rep['quarter'].split('-')
            self.stat_last.set_values(
                rep['processed_at'][:16].replace('T', ' '),
                f"{ROMAN_REV[qn]} кв. · {rep['rows']} строк")
            self.stat_dup.set_values(str(rep['dups']), "дубли · фиолетовые строки")
        else:
            self.stat_last.set_values("ещё не было", "")
            self.stat_dup.set_values("—", "дубли · фиолетовые строки")
        self.stat_tax.set_values("—", "сумма не хранится в БД")


# ================================================================
# Приложение
# ================================================================
class RealApp(TurNalogApp):
    def __init__(self):
        super().__init__()
        self.cfg = load_config()
        self.app_version = APP_VERSION
        self._update_info = None
        self._update_plan = None
        self.slot_state = {}
        self._sync_theme()

    # ---------- тема/демо-данные → реальные значения ----------
    def _sync_theme(self):
        theme.DB_PATH_DEMO = default_db_path()
        for hid, code in ID2CODE.items():
            tp = TaxParams(self.cfg.get('tax', {}).get(code))
            h = theme.HOTELS[hid]
            h['rate'] = f"{tp.rate:g}"
            h['min_night'] = f"{tp.min:g}"
            h['threshold'] = f"{tp.threshold:g}"
            h['tax_in_price'] = tp.incl
            h['auto_threshold'] = tp.auto
            h['tax_no_commission'] = tp.use_full
            h['formula_daily'] = tp.tpl_daily
            h['formula_method'] = tp.tpl_method
            h['formula_tax'] = tp.tpl_tax
        theme.AGENTS[:] = [a['name'] for a in migrate_agents(self.cfg)]

    def _build_main(self):
        self.main = ctk.CTkFrame(self, fg_color=theme.COLORS["bg"], corner_radius=0)
        self.main.pack(side="left", fill="both", expand=True)
        self.process_view = RealProcessView(self.main, self)
        self.history_view = RealHistoryView(self.main, self)

    # ---------- автообновление ----------
    def start_main(self, user):
        super().start_main(user)
        try:
            if self.cfg.get("updates", {}).get("auto_check", True):
                self.after(2500, self._auto_update_check)
        except Exception:
            pass

    def _auto_update_check(self):
        threading.Thread(target=self._auto_update_worker, daemon=True).start()

    def _auto_update_worker(self):
        try:
            info = check_for_updates(self.app_version)
            err = None
        except UpdateError as e:
            info, err = None, str(e)
        except Exception as e:  # noqa: BLE001
            info, err = None, f"Неожиданная ошибка: {e}"
        try:
            self.after(0, lambda: self._auto_update_done(info, err))
        except Exception:
            pass

    def _auto_update_done(self, info, err):
        if err or info is None:
            return  # тихая проверка: молчим и при ошибке, и при «нет обновлений»
        self._update_info = info
        uc = self.cfg.setdefault("updates", {})
        uc["last_check"] = datetime.now().isoformat(timespec="seconds")
        uc["last_version"] = info.version
        save_config(self.cfg)
        self.set_status(f"Доступна версия {info.version} — Настройки → Обновления")
        try:
            self.toast(f"Доступна версия {info.version}")
        except Exception:
            pass
        if uc.get("auto_download") and is_frozen() and os.name == "nt":
            threading.Thread(target=self._auto_download_worker,
                             args=(info,), daemon=True).start()

    def _auto_download_worker(self, info):
        try:
            zp = download_release(info, staging_dir(info.tag))
            verify_package(zp)
            plan = plan_install(zp)
            err = None
        except UpdateError as e:
            plan, err = None, str(e)
        except Exception as e:  # noqa: BLE001
            plan, err = None, f"Неожиданная ошибка: {e}"
        try:
            self.after(0, lambda: self._auto_download_done(info, plan, err))
        except Exception:
            pass

    def _auto_download_done(self, info, plan, err):
        if err or plan is None:
            self.set_status(f"Автообновление не удалось: {err}")
            return
        self._update_plan = plan
        self.set_status(f"Версия {info.version} загружена — Настройки → Обновления")
        if messagebox.askyesno(
                "Обновление загружено",
                f"Версия {info.version} загружена автоматически.\n\n"
                "Установить сейчас? Приложение перезапустится.",
                parent=self):
            try:
                launch_install(plan)
            except UpdateError as e:
                messagebox.showerror("Обновление", str(e), parent=self)
                return
            self.destroy()

    # ---------- данные истории ----------

    def history_rows(self):
        hotel = None if self.user.get('role') == 'администратор' \
            else theme.HOTELS[self.user['hotel_id']]['name']
        rows = []
        with HistoryDB() as db:
            for r in db.list_reports(hotel):
                cross = [f"{b['guest'] or '—'} · {b['checkin'] or '—'}…{b['checkout'] or '—'}"
                         for b in db.dup_bookings(r['id'])]
                y, qn = r['quarter'].split('-')
                rows.append({
                    'id': str(r['id']), 'hotel_id': NAME2ID[r['hotel']],
                    'hotel': r['hotel'], 'quarter': f"{ROMAN_REV[qn]} {y}",
                    'processed': r['processed_at'][:16].replace('T', ' '),
                    'rows': str(r['rows']), 'dups': str(r['dups']),
                    'orig': r['original_name'], 'result': r['output_name'],
                    'cross': cross})
        return rows

    def last_report(self, hotel_id):
        name = theme.HOTELS[hotel_id]['name']
        with HistoryDB() as db:
            reps = db.list_reports(name)
        return reps[0] if reps else None

    # ---------- хуки ----------
    def hook_authenticate(self, login, password):
        with HistoryDB() as db:
            u = db.check_user(login, password)
        if not u:
            return None
        return {'login': u['login'],
                'role': 'администратор' if u['role'] == 'admin' else 'отель',
                'hotel_id': CODE2ID.get(u['hotel'])}

    def hook_open_file(self, kind, row):
        key = 'orig' if kind == 'original' else 'out'
        with HistoryDB() as db:
            name, blob = db.get_file(int(row['id']), key)
        if not name or blob is None:
            messagebox.showwarning("История", "Файл не найден в БД.", parent=self)
            return
        import tempfile, zipfile as _zip
        tmp = tempfile.mkdtemp(prefix="turnalog_")
        target = os.path.join(tmp, name)
        with open(target, 'wb') as f:
            f.write(blob)
        if name.lower().endswith('.zip'):
            with _zip.ZipFile(target) as z:
                z.extractall(tmp)
            open_in_system(tmp)
        else:
            open_in_system(target)

    def hook_save_settings(self, payload):
        for hid, code in ID2CODE.items():
            p = payload[hid]
            try:
                self.cfg.setdefault('tax', {})[code] = {
                    'rate': float(p['rate']), 'min_per_night': float(p['min_night']),
                    'threshold': float(p['threshold']), 'incl': p['tax_in_price'],
                    'auto_threshold': p['auto_threshold'],
                    'use_full': p['tax_no_commission'],
                    'tpl_daily': p['formula_daily'], 'tpl_method': p['formula_method'],
                    'tpl_tax': p['formula_tax']}
            except ValueError:
                messagebox.showerror("Настройки",
                                     f"«{theme.HOTELS[hid]['name']}»: числа содержат нечисловые значения.")
                return
        if payload.get('agents'):
            self.cfg['agents'] = payload['agents']
        save_config(self.cfg)
        self._sync_theme()
        self.toast("Настройки сохранены")
        self.set_status("Настройки сохранены")

    def hook_save_columns(self, kind, mapping):
        pass  # сохраняется в open_column_map

    # ---------- импорт из TravelLine ----------
    def hook_tl_import(self, hotel_id):
        from tl_client import (bookings_to_rows, fetch_details,
                               tl_export_paths, write_data_workbook,
                               write_rooms_stub)
        code = ID2CODE[hotel_id]
        hname = theme.HOTELS[hotel_id]['name']
        pv = self.process_view
        quarter, year = pv.qseg.get(), pv.year.get().strip()
        if quarter not in ROMAN or not year.isdigit():
            messagebox.showerror('TravelLine', 'Укажите корректные квартал и год.',
                                 parent=self)
            return
        tl_cfg = self.cfg.setdefault('travelline', {}).setdefault(code, {})
        dlg = TlConfigDialog(self, tl_cfg, hname)
        self.wait_window(dlg)
        if not dlg.result:
            return
        tl_cfg.update(dlg.result)
        save_config(self.cfg)
        qn, y = {'I': 1, 'II': 2, 'III': 3, 'IV': 4}[quarter], int(year)
        qs = f'{y}-{(qn - 1) * 3 + 1:02d}-01'
        qe = f'{y + 1}-01-01' if qn == 4 else f'{y}-{qn * 3 + 1:02d}-01'
        depth = f'{y - 2}-01-01'
        self.set_status('TravelLine: получение токена…')

        def work():
            def progress(stage, done, total):
                if stage == 'details' and total:
                    self.after(0, lambda: self.set_status(
                        f'TravelLine: {done}/{total}…'))
                elif stage == 'list':
                    self.after(0, lambda: self.set_status(
                        f'TravelLine: список ({done})…'))
            details = fetch_details(tl_cfg['client_id'], tl_cfg['client_secret'],
                                    tl_cfg['property_id'], depth, progress)
            rows, rep = bookings_to_rows(details, qs, qe)
            data_path, rooms_path = tl_export_paths(tempfile.gettempdir(),
                                                    code, y, quarter)
            write_data_workbook(rows, data_path)
            write_rooms_stub(rooms_path)
            return rows, rep, data_path, rooms_path

        def done(result, err):
            if err:
                self.set_status('TravelLine: ошибка')
                messagebox.showerror(
                    'TravelLine', f'Не удалось загрузить брони:\n{err}',
                    parent=self)
                return
            rows, rep, data_path, rooms_path = result
            pv.set_tl_files(data_path, rooms_path)
            self.set_status(f'TravelLine: {rep["rows"]} строк за '
                            f'{quarter}/{year} → нажмите «Обработать»')
            lines = [f'Броней в квартале: {rep["in_quarter"]} '
                     f'(строк: {rep["rows"]})',
                     f'Пропущено отменённых: {rep["cancelled"]}']
            if rep['mismatch']:
                lines.append(f'⚠ Расхождений арифметики: {len(rep["mismatch"])} '
                             '(номера в логе) — проверьте!')
                for m in rep['mismatch'][:20]:
                    log.warning('TL mismatch %s', m)
            lines.append('Файлы подставлены — нажмите «Обработать и сохранить».')
            messagebox.showinfo('TravelLine', '\n'.join(lines), parent=self)

        def runner():
            try:
                r, e = work(), None
            except Exception as exc:
                log.exception('tl_import')
                r, e = None, exc
            self.after(0, lambda: done(r, e))

        threading.Thread(target=runner, daemon=True).start()

    # ---------- сопоставление колонок (боевое) ----------

    def open_column_map(self, kind):
        pid = self.current_hotel
        code = ID2CODE[pid]
        slot = 'data' if kind == 'data' else 'rooms'
        if not needs_mapping(code, slot):
            messagebox.showinfo("Колонки",
                                "У этого отеля структура отчёта фиксированная.", parent=self)
            return
        pv = self.process_view
        path = pv.drop_data.path if slot == 'data' else pv.drop_rooms.path
        if not path:
            messagebox.showinfo("Колонки", "Сначала выберите файл.", parent=self)
            return
        st = self.slot_state.setdefault(code, {}).setdefault(slot, {})
        try:
            headers = read_headers(path)
        except Exception as e:
            messagebox.showerror("Колонки", f"Файл не читается:\n{e}", parent=self)
            return
        dlg = MappingDialog(self, os.path.basename(path), headers, SCHEMAS[(code, slot)],
                            mapping=st.get('mapping'), extra=st.get('extra'))
        self.wait_window(dlg)
        if dlg.result:
            st['mapping'], st['extra'] = dlg.result
            self.cfg[f"{code}_{slot}"] = {"mapping": dlg.result[0], "extra": dlg.result[1]}
            save_config(self.cfg)
            self.toast("Сопоставление колонок сохранено")

    def _resolve_unrecognized(self, code, agents, sources):
        agents = copy.deepcopy(agents)
        counts = {}
        for s in sources:
            s = (s or '').strip()
            if s:
                counts[s] = counts.get(s, 0) + 1
        decided_none = set(self.cfg.get('source_none', []))
        items = [(s, c) for s, c in counts.items()
                 if s not in decided_none
                 and match_agent(s, agents, hotel=code)[1] is None]
        # ранее оставленные «без подсветки» — сразу в карту с комиссией 0%
        source_map = {s: (0, '') for s in counts if s in decided_none}
        if not items:
            return agents, source_map
        dlg = UnrecognizedSourcesDialog(self, items, [a['name'] for a in agents])
        self.wait_window(dlg)
        if dlg.result is None:
            return None, None
        changed = False
        for s, res in dlg.result.items():
            if res[0] == 'agent':
                for a in agents:
                    if a['name'] == res[1]:
                        if s not in a['aliases']:
                            a['aliases'].append(s)
                        if code not in a['hotels']:
                            a['hotels'].append(code)
                        changed = True
                        break
            elif res[0] == 'new':
                agents.append({'name': res[1], 'aliases': [s], 'hotels': [code],
                               'rates': [{'from': '', 'to': '', 'pct': res[2]}]})
                source_map[s] = (res[2], res[1])
                changed = True
            elif res[0] == 'none':
                # «— не указано —» = комиссия 0% без подсветки; запоминаем,
                # чтобы диалог больше не спрашивал про этот источник
                source_map[s] = (0, '')
                if s not in decided_none:
                    decided_none.add(s)
                    changed = True
        if changed:
            self.cfg['agents'] = agents
            self.cfg['source_none'] = sorted(decided_none)
            save_config(self.cfg)
        return agents, source_map

    # ---------- обработка ----------
    def hook_process(self, hotel_id, data_path, rooms_path, quarter, year,
                     check_dups, dest_path):
        code = ID2CODE[hotel_id]
        hname = theme.HOTELS[hotel_id]['name']
        q = f"{year}-{ROMAN[quarter]}"
        tax = self.cfg.get('tax', {}).get(code)
        agents = migrate_agents(self.cfg)
        try:
            maps = {}
            for slot, path in (('data', data_path), ('rooms', rooms_path)):
                if not path:
                    continue
                if not needs_mapping(code, slot):
                    continue  # фиксированная структура (Усадьба)
                if is_tl_export(path):
                    tlm = TL_DATA_MAP if slot == 'data' else TL_ROOMS_MAP
                    maps[slot] = (dict(tlm),
                                  list(TL_EXTRA) if slot == 'data' else [])
                    continue
                headers = read_headers(path)
                saved = self.cfg.get(f"{code}_{slot}", {})
                if mapping_valid(saved.get('mapping'), headers, SCHEMAS[(code, slot)]):
                    maps[slot] = (saved['mapping'], saved.get('extra', []))
                else:
                    self.current_hotel = hotel_id
                    self.open_column_map(slot)
                    st = self.slot_state.get(code, {}).get(slot, {})
                    if not st.get('mapping'):
                        return
                    maps[slot] = (st['mapping'], st.get('extra', []))
            if code in ('tch', 'dna'):
                dfc = read_table(data_path)
                scol = maps['data'][0].get('source')
                sources = list(dfc[scol].fillna('').astype(str)) \
                    if scol and scol in dfc.columns else []
            else:
                sources = usadba_sources(data_path)
            res = self._resolve_unrecognized(code, agents, sources)
            if res[0] is None:
                return
            agents, source_map = res
        except Exception as e:
            log.exception("hook_process pre")
            messagebox.showerror("Ошибка", f"Не удалось прочитать файлы:\n{e}", parent=self)
            return

        with HistoryDB() as db:
            old = db.find_same(hname, q)
        replace = False
        if old:
            replace = messagebox.askyesno(
                "Отчёт уже есть в истории",
                f"За квартал {q} уже обработан отчёт от {old['processed_at'][:10]}\n"
                f"(файл: {old['original_name']}).\n\nЗаменить его?", parent=self)

        self.set_status("Обработка…")
        pv = self.process_view
        pv.run_btn.configure(state="disabled", fg_color=theme.COLORS["green_dis"])

        def work():
            with HistoryDB() as db:
                if old and replace:
                    db.delete_report(old['id'])
                checker = (lambda rows: db.check_duplicates(hname, rows)) if check_dups else None
                if code == 'tch':
                    summary, bookings = process_tchaikovsky(
                        data_path, rooms_path, dest_path,
                        maps['data'][0], maps.get('rooms', (None, None))[0],
                        extra_data=maps['data'][1], dup_checker=checker, tax=tax,
                        agents=agents, source_map=source_map)
                elif code == 'dna':
                    summary, bookings = process_dona(
                        data_path, dest_path, maps['data'][0],
                        extra_data=maps['data'][1], dup_checker=checker, tax=tax,
                        agents=agents, source_map=source_map)
                else:
                    summary, bookings = process_usadba(
                        data_path, dest_path, dup_checker=checker, tax=tax,
                        agents=agents, source_map=source_map)
                orig_name, orig_blob = pack_originals([p for p in (data_path, rooms_path) if p])
                with open(dest_path, 'rb') as f:
                    out_blob = f.read()
                dups = sum(1 for b in bookings if b.get('dup_note'))
                rid = db.add_report(hname, q, orig_name, orig_blob,
                                    os.path.basename(dest_path), out_blob,
                                    len(bookings), dups, "")
                db.add_bookings(rid, bookings)
            return summary, dups

        def done(result, err):
            pv.run_btn.configure(state="normal", fg_color=theme.COLORS["green"])
            if err:
                self.set_status("Ошибка обработки")
                messagebox.showerror("Ошибка", f"В процессе обработки произошла ошибка:\n{err}",
                                     parent=self)
                return
            summary, dups = result
            self.history_view.refresh()
            self.set_status(f"{hname}: готово → {os.path.basename(dest_path)}")
            total = sum(summary.values()) if isinstance(summary, dict) else 0
            rep = repair_notes_for([p for p in (data_path, rooms_path) if p])
            if total > 0 or rep:
                msg = f"Файл сохранён, замечаний: {total}.\n\n" if total > 0 \
                    else "Файл сохранён.\n\n"
                if total > 0:
                    msg += "\n".join(f"• {k}: {v}" for k, v in summary.items() if v)
                if rep:
                    msg += ("\n\n" if total > 0 else "") + "\n".join(f"• {n}" for n in rep)
                messagebox.showwarning("Готово (есть замечания)", msg, parent=self)
            else:
                self.toast("Файл сохранён без замечаний")

        def runner():
            try:
                r = work()
                e = None
            except Exception as exc:
                log.exception("processing")
                r, e = None, exc
            self.after(0, lambda: done(r, e))

        threading.Thread(target=runner, daemon=True).start()

    # ---------- настройки ----------
    def open_settings(self):
        if self._settings is not None and self._settings.winfo_exists():
            self._settings.lift()
            self._settings.focus()
            return
        self._sync_theme()
        self._settings = RealSettingsWindow(self)


def main():
    with HistoryDB() as db:
        db.ensure_default_admin()
    theme.DB_PATH_DEMO = default_db_path()
    ctk.set_appearance_mode("light")
    try:
        ctk.set_default_color_theme("blue")
    except Exception:
        pass
    app = RealApp()
    app.mainloop()


if __name__ == "__main__":
    main()
