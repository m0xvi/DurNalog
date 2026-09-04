# -*- coding: utf-8 -*-
"""Локальная база истории отчётов (SQLite, входит в Python).

Хранит по каждому обработанному отчёту: отель, квартал, дату обработки,
КОПИИ исходного и итогового файлов и распознанные брони. Нужна, чтобы
находить брони на переходе кварталов (один гость с теми же датами попадает
в два квартальных отчёта с полной стоимостью) и помечать их.

Файл БД лежит рядом с приложением: turnalog_history.db
"""
import hashlib
import os
import secrets
import sqlite3
import sys
from datetime import datetime

from tax import dates_overlap


def default_db_path():
    if getattr(sys, 'frozen', False):           # собранный EXE
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, 'turnalog_history.db')


class HistoryDB:
    def __init__(self, path=None):
        self.path = path or default_db_path()
        self.con = sqlite3.connect(self.path)
        self.con.row_factory = sqlite3.Row
        self._init_schema()

    # ------------------------------------------------------------------
    def _init_schema(self):
        self.con.executescript("""
        CREATE TABLE IF NOT EXISTS reports (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            hotel         TEXT NOT NULL,
            quarter       TEXT NOT NULL,
            processed_at  TEXT NOT NULL,
            original_name TEXT,
            original_blob BLOB,
            output_name   TEXT,
            output_blob   BLOB,
            rows          INTEGER DEFAULT 0,
            dups          INTEGER DEFAULT 0,
            warnings      TEXT
        );
        CREATE TABLE IF NOT EXISTS bookings (
            id         INTEGER PRIMARY KEY AUTOINCREMENT,
            report_id  INTEGER NOT NULL,
            guest      TEXT,
            guest_norm TEXT,
            bron       TEXT,
            checkin    TEXT,
            checkout   TEXT,
            price      REAL,
            price_full REAL,
            nights     INTEGER,
            dup_flag   INTEGER DEFAULT 0,
            dup_note   TEXT
        );
        CREATE INDEX IF NOT EXISTS idx_bookings_report ON bookings(report_id);
        CREATE INDEX IF NOT EXISTS idx_bookings_hotel ON bookings(guest_norm);
        CREATE TABLE IF NOT EXISTS users (
            id     INTEGER PRIMARY KEY AUTOINCREMENT,
            login  TEXT UNIQUE NOT NULL,
            salt   TEXT NOT NULL,
            hash   TEXT NOT NULL,
            role   TEXT NOT NULL,
            hotel  TEXT
        );
        """)
        self.con.commit()

    # контекстный менеджер: with HistoryDB() as db: ...
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        self.con.commit()
        self.con.close()
        return False

    def close(self):
        self.con.commit()
        self.con.close()

    # ------------------------------------------------------------------
    # Отчёты
    # ------------------------------------------------------------------
    def list_reports(self, hotel=None):
        if hotel:
            cur = self.con.execute(
                "SELECT * FROM reports WHERE hotel=? "
                "ORDER BY processed_at DESC, id DESC", (hotel,))
        else:
            cur = self.con.execute(
                "SELECT * FROM reports ORDER BY processed_at DESC, id DESC")
        return cur.fetchall()

    def find_same(self, hotel, quarter):
        cur = self.con.execute(
            "SELECT * FROM reports WHERE hotel=? AND quarter=?", (hotel, quarter))
        return cur.fetchone()

    def delete_report(self, report_id):
        self.con.execute("DELETE FROM bookings WHERE report_id=?", (report_id,))
        self.con.execute("DELETE FROM reports WHERE id=?", (report_id,))
        self.con.commit()

    def add_report(self, hotel, quarter, original_name, original_blob,
                   output_name, output_blob, rows, dups, warnings=''):
        cur = self.con.execute(
            "INSERT INTO reports (hotel, quarter, processed_at, original_name,"
            " original_blob, output_name, output_blob, rows, dups, warnings)"
            " VALUES (?,?,?,?,?,?,?,?,?,?)",
            (hotel, quarter, datetime.now().isoformat(timespec='seconds'),
             original_name, original_blob, output_name, output_blob,
             rows, dups, warnings))
        return cur.lastrowid

    def add_bookings(self, report_id, bookings):
        """bookings: список dict(guest, guest_norm, bron, checkin, checkout,
        price, price_full, nights, dup_flag, dup_note); даты — ISO-строки или None."""
        self.con.executemany(
            "INSERT INTO bookings (report_id, guest, guest_norm, bron, checkin,"
            " checkout, price, price_full, nights, dup_flag, dup_note)"
            " VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            [(report_id, b.get('guest'), b.get('guest_norm'), b.get('bron'),
              b.get('checkin'), b.get('checkout'), b.get('price'),
              b.get('price_full'), b.get('nights'),
              1 if b.get('dup_note') else 0, b.get('dup_note'))
             for b in bookings])

    def get_file(self, report_id, kind):
        col = 'original_name, original_blob' if kind == 'orig' else 'output_name, output_blob'
        cur = self.con.execute(f"SELECT {col} FROM reports WHERE id=?", (report_id,))
        row = cur.fetchone()
        return (row[0], row[1]) if row else (None, None)

    # ------------------------------------------------------------------
    # Дубли (переход кварталов)
    # ------------------------------------------------------------------
    def _hotel_bookings(self, hotel):
        cur = self.con.execute(
            """SELECT b.*, r.quarter AS r_quarter, r.processed_at AS r_at,
                      r.original_name AS r_file
               FROM bookings b JOIN reports r ON r.id = b.report_id
               WHERE r.hotel = ?""", (hotel,))
        return cur.fetchall()

    def check_duplicates(self, hotel, rows):
        """rows: список dict(bron, guest_norm, checkin(date|None), checkout(date|None)).
        Возвращает список примечаний той же длины (None, если дубль не найден).
        Ключ: точное совпадение № брони ИЛИ (то же ФИО + строгое пересечение дат)."""
        prev = self._hotel_bookings(hotel)
        notes = []
        for r in rows:
            note = None
            bron = (r.get('bron') or '').strip()
            if bron:
                for p in prev:
                    if p['bron'] and p['bron'].strip() == bron:
                        note = self._note(p)
                        break
            if note is None and r.get('guest_norm') and r.get('checkin') and r.get('checkout'):
                for p in prev:
                    if (p['guest_norm'] and p['guest_norm'] == r['guest_norm']
                            and dates_overlap(r['checkin'], r['checkout'],
                                              _iso(p['checkin']), _iso(p['checkout']))):
                        note = self._note(p)
                        break
            notes.append(note)
        return notes

    @staticmethod
    def _note(p):
        when = (p['r_at'] or '')[:10]
        return (f"Переход кварталов: эта бронь уже учтена в отчёте "
                f"{p['r_quarter']} от {when} (файл: {p['r_file']})")

    # ------------------------------------------------------------------
    # Пользователи и права
    # ------------------------------------------------------------------
    @staticmethod
    def _pw_hash(pw, salt):
        return hashlib.sha256((salt + pw).encode('utf-8')).hexdigest()

    def ensure_default_admin(self):
        """Первый запуск: создаёт admin/admin (смените пароль в «Пользователи»)."""
        n = self.con.execute("SELECT COUNT(*) FROM users").fetchone()[0]
        if n == 0:
            self.add_user('admin', 'admin', 'admin', None)

    def add_user(self, login, password, role, hotel=None):
        if not login or not password:
            return False
        if self.con.execute("SELECT 1 FROM users WHERE login=?",
                            (login,)).fetchone():
            return False
        salt = secrets.token_hex(8)
        self.con.execute(
            "INSERT INTO users (login, salt, hash, role, hotel) VALUES (?,?,?,?,?)",
            (login, salt, self._pw_hash(password, salt), role, hotel))
        self.con.commit()
        return True

    def check_user(self, login, password):
        row = self.con.execute(
            "SELECT * FROM users WHERE login=?", (login,)).fetchone()
        if not row:
            return None
        if self._pw_hash(password, row['salt']) == row['hash']:
            return {'id': row['id'], 'login': row['login'],
                    'role': row['role'], 'hotel': row['hotel']}
        return None

    def list_users(self):
        return self.con.execute(
            "SELECT id, login, role, hotel FROM users ORDER BY id").fetchall()

    def del_user(self, user_id):
        """Нельзя удалить последнего админа."""
        admins = self.con.execute(
            "SELECT COUNT(*) FROM users WHERE role='admin'").fetchone()[0]
        target = self.con.execute(
            "SELECT role FROM users WHERE id=?", (user_id,)).fetchone()
        if target and target['role'] == 'admin' and admins <= 1:
            return False
        self.con.execute("DELETE FROM users WHERE id=?", (user_id,))
        self.con.commit()
        return True

    def set_password(self, user_id, password):
        salt = secrets.token_hex(8)
        self.con.execute("UPDATE users SET salt=?, hash=? WHERE id=?",
                         (salt, self._pw_hash(password, salt), user_id))
        self.con.commit()

    # ------------------------------------------------------------------
    def dup_bookings(self, report_id):
        cur = self.con.execute(
            "SELECT * FROM bookings WHERE report_id=? AND dup_flag=1", (report_id,))
        return cur.fetchall()


def _iso(s):
    """'2026-04-02' -> date | None"""
    from tax import as_date
    return as_date(s) if s else None
