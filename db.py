# -*- coding: utf-8 -*-
"""Локальная база истории отчётов (SQLite, входит в Python).

Хранит по каждому обработанному отчёту: отель, квартал, дату обработки,
КОПИИ исходного и итогового файлов и распознанные брони. Нужна, чтобы
находить брони на переходе кварталов (один гость с теми же датами попадает
в два квартальных отчёта с полной стоимостью) и помечать их.

Файл БД лежит рядом с приложением: turnalog_history.db
"""
import hashlib
import json
import os
import secrets
import sqlite3
import sys
import zlib
from datetime import date, datetime, timezone

from tax import dates_overlap


def default_db_path():
    if _DB_PATH_OVERRIDE:
        return _DB_PATH_OVERRIDE
    if getattr(sys, 'frozen', False):           # собранный EXE
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, 'turnalog_history.db')


_DB_PATH_OVERRIDE = None


def set_db_path(path):
    """Путь к БД из конфига/диалога первого запуска. None = по умолчанию."""
    global _DB_PATH_OVERRIDE
    _DB_PATH_OVERRIDE = path or None


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
        CREATE TABLE IF NOT EXISTS travelline_cache (
            cache_key    TEXT PRIMARY KEY,
            hotel_code   TEXT NOT NULL,
            property_id  TEXT NOT NULL,
            period_start TEXT NOT NULL,
            period_end   TEXT NOT NULL,
            modified_since TEXT NOT NULL,
            fetched_at   TEXT NOT NULL,
            row_count    INTEGER NOT NULL DEFAULT 0,
            payload_json TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS travelline_archive (
            hotel_code      TEXT NOT NULL,
            property_id     TEXT NOT NULL,
            booking_number  TEXT NOT NULL,
            arrival_date    TEXT,
            departure_date  TEXT,
            modified_at     TEXT,
            fetched_at      TEXT NOT NULL,
            payload_blob    BLOB NOT NULL,
            PRIMARY KEY (hotel_code, property_id, booking_number)
        );
        CREATE INDEX IF NOT EXISTS idx_travelline_archive_period
            ON travelline_archive(hotel_code, property_id, arrival_date, departure_date);
        CREATE TABLE IF NOT EXISTS travelline_archive_state (
            hotel_code       TEXT NOT NULL,
            property_id      TEXT NOT NULL,
            full_sync_complete INTEGER NOT NULL DEFAULT 0,
            last_sync_at     TEXT,
            last_full_sync_at TEXT,
            sync_started_at  TEXT,
            sync_kind        TEXT,
            last_error       TEXT,
            PRIMARY KEY (hotel_code, property_id)
        );
        CREATE TABLE IF NOT EXISTS travelline_inventory (
            hotel_code   TEXT NOT NULL,
            property_id  TEXT NOT NULL,
            fetched_at   TEXT NOT NULL,
            inventory_json TEXT NOT NULL,
            PRIMARY KEY (hotel_code, property_id)
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
    # TravelLine cache (только готовые к импорту строки, не сырой API JSON)
    # ------------------------------------------------------------------
    def get_travelline_cache(self, cache_key):
        row = self.con.execute(
            "SELECT * FROM travelline_cache WHERE cache_key=?", (cache_key,)
        ).fetchone()
        return dict(row) if row else None

    def put_travelline_cache(self, cache_key, hotel_code, property_id,
                             period_start, period_end, modified_since,
                             row_count, payload_json):
        self.con.execute(
            """INSERT INTO travelline_cache
               (cache_key, hotel_code, property_id, period_start, period_end,
                modified_since, fetched_at, row_count, payload_json)
               VALUES (?,?,?,?,?,?,?,?,?)
               ON CONFLICT(cache_key) DO UPDATE SET
                 hotel_code=excluded.hotel_code,
                 property_id=excluded.property_id,
                 period_start=excluded.period_start,
                 period_end=excluded.period_end,
                 modified_since=excluded.modified_since,
                 fetched_at=excluded.fetched_at,
                 row_count=excluded.row_count,
                 payload_json=excluded.payload_json""",
            (cache_key, hotel_code, str(property_id), period_start, period_end,
             modified_since, datetime.now().isoformat(timespec='seconds'),
             int(row_count), payload_json))

    def delete_travelline_cache(self, cache_key):
        self.con.execute(
            "DELETE FROM travelline_cache WHERE cache_key=?", (cache_key,))

    def delete_travelline_caches_for_property(self, hotel_code, property_id):
        self.con.execute(
            "DELETE FROM travelline_cache WHERE hotel_code=? AND property_id=?",
            (hotel_code, str(property_id)))

    @staticmethod
    def _tl_archive_now():
        return (datetime.now(timezone.utc).replace(microsecond=0)
                .isoformat().replace('+00:00', 'Z'))

    @staticmethod
    def _tl_archive_day(value):
        if isinstance(value, datetime):
            value = value.date()
        if isinstance(value, date):
            return value.isoformat()
        try:
            return date.fromisoformat(str(value or '')[:10]).isoformat()
        except (TypeError, ValueError):
            return None

    @classmethod
    def _tl_booking_period(cls, booking):
        room_stays = booking.get('roomStays') or []
        if isinstance(room_stays, dict):
            room_stays = [room_stays]
        arrivals, departures = [], []
        for stay in room_stays:
            if not isinstance(stay, dict):
                continue
            dates = stay.get('stayDates') or {}
            if not isinstance(dates, dict):
                continue
            arrival = cls._tl_archive_day(dates.get('arrivalDateTime'))
            departure = cls._tl_archive_day(dates.get('departureDateTime'))
            if arrival:
                arrivals.append(arrival)
            if departure:
                departures.append(departure)
        if not arrivals and not departures:
            dates = booking.get('stayDates') or {}
            if isinstance(dates, dict):
                arrival = cls._tl_archive_day(dates.get('arrivalDateTime'))
                departure = cls._tl_archive_day(dates.get('departureDateTime'))
                if arrival:
                    arrivals.append(arrival)
                if departure:
                    departures.append(departure)
        return (min(arrivals) if arrivals else None,
                max(departures) if departures else None)

    def upsert_travelline_archive_bookings(self, hotel_code, property_id,
                                           bookings):
        """Сохраняет детали брони целиком в сжатом локальном архиве."""
        now = self._tl_archive_now()
        records = []
        skipped = 0
        for booking in bookings or []:
            if not isinstance(booking, dict):
                skipped += 1
                continue
            number = (booking.get('number') or booking.get('bookingNumber')
                      or booking.get('id') or booking.get('bookingId'))
            if number is None or str(number).strip() == '':
                skipped += 1
                continue
            arrival, departure = self._tl_booking_period(booking)
            modified = (booking.get('lastModification')
                        or booking.get('lastModified') or '')
            payload = json.dumps(
                booking, ensure_ascii=False, separators=(',', ':'),
                default=lambda value: value.isoformat()
                if hasattr(value, 'isoformat') else str(value))
            records.append((
                str(hotel_code), str(property_id), str(number), arrival,
                departure, str(modified) if modified else '', now,
                sqlite3.Binary(zlib.compress(payload.encode('utf-8'), level=6))))
        self.con.executemany(
            """INSERT INTO travelline_archive
               (hotel_code, property_id, booking_number, arrival_date,
                departure_date, modified_at, fetched_at, payload_blob)
               VALUES (?,?,?,?,?,?,?,?)
               ON CONFLICT(hotel_code, property_id, booking_number) DO UPDATE SET
                 arrival_date=excluded.arrival_date,
                 departure_date=excluded.departure_date,
                 modified_at=excluded.modified_at,
                 fetched_at=excluded.fetched_at,
                 payload_blob=excluded.payload_blob""",
            records)
        return {'saved': len(records), 'skipped': skipped}

    def get_travelline_archive_bookings(self, hotel_code, property_id,
                                        period_start, period_end_exclusive):
        """Возвращает архивные брони с ночами, пересекающими [start, end)."""
        rows = self.con.execute(
            """SELECT payload_blob FROM travelline_archive
               WHERE hotel_code=? AND property_id=?
                 AND arrival_date IS NOT NULL AND departure_date IS NOT NULL
                 AND departure_date>? AND arrival_date<?
               ORDER BY arrival_date, booking_number""",
            (str(hotel_code), str(property_id), str(period_start),
             str(period_end_exclusive))).fetchall()
        return [json.loads(zlib.decompress(row['payload_blob']).decode('utf-8'))
                for row in rows]

    def get_travelline_archive_stats(self, hotel_code, property_id):
        hotel_code, property_id = str(hotel_code), str(property_id)
        state = self.con.execute(
            """SELECT * FROM travelline_archive_state
               WHERE hotel_code=? AND property_id=?""",
            (hotel_code, property_id)).fetchone()
        totals = self.con.execute(
            """SELECT COUNT(*) AS booking_count,
                      COALESCE(SUM(length(payload_blob)), 0) AS archive_bytes
               FROM travelline_archive WHERE hotel_code=? AND property_id=?""",
            (hotel_code, property_id)).fetchone()
        inventory = self.con.execute(
            """SELECT COALESCE(length(CAST(inventory_json AS BLOB)), 0)
               FROM travelline_inventory WHERE hotel_code=? AND property_id=?""",
            (hotel_code, property_id)).fetchone()
        try:
            database_bytes = os.path.getsize(self.path)
        except (OSError, TypeError, ValueError):
            database_bytes = 0
        result = {
            'booking_count': int(totals['booking_count']),
            'archive_bytes': int(totals['archive_bytes']),
            'inventory_bytes': int(inventory[0]) if inventory else 0,
            'database_bytes': int(database_bytes),
            'database_path': self.path,
            'full_sync_complete': False,
            'last_sync_at': None,
            'last_full_sync_at': None,
            'sync_started_at': None,
            'sync_kind': None,
            'last_error': None,
        }
        if state:
            result.update({
                'full_sync_complete': bool(state['full_sync_complete']),
                'last_sync_at': state['last_sync_at'],
                'last_full_sync_at': state['last_full_sync_at'],
                'sync_started_at': state['sync_started_at'],
                'sync_kind': state['sync_kind'],
                'last_error': state['last_error'],
            })
        return result

    def get_travelline_archive_sync_state(self, hotel_code, property_id):
        row = self.con.execute(
            """SELECT * FROM travelline_archive_state
               WHERE hotel_code=? AND property_id=?""",
            (str(hotel_code), str(property_id))).fetchone()
        return dict(row) if row else None

    def begin_travelline_archive_sync(self, hotel_code, property_id, full=False):
        now = self._tl_archive_now()
        self.con.execute(
            """INSERT INTO travelline_archive_state
               (hotel_code, property_id, full_sync_complete, sync_started_at,
                sync_kind, last_error)
               VALUES (?,?,0,?,?,NULL)
               ON CONFLICT(hotel_code, property_id) DO UPDATE SET
                 sync_started_at=excluded.sync_started_at,
                 sync_kind=excluded.sync_kind, last_error=NULL""",
            (str(hotel_code), str(property_id), now,
             'full' if full else 'incremental'))

    def finish_travelline_archive_sync(self, hotel_code, property_id, full=False):
        now = self._tl_archive_now()
        self.con.execute(
            """INSERT INTO travelline_archive_state
               (hotel_code, property_id, full_sync_complete, last_sync_at,
                last_full_sync_at, sync_started_at, sync_kind, last_error)
               VALUES (?,?,?, ?, ?, NULL, NULL, NULL)
               ON CONFLICT(hotel_code, property_id) DO UPDATE SET
                 full_sync_complete=MAX(
                     travelline_archive_state.full_sync_complete,
                     excluded.full_sync_complete),
                 last_sync_at=excluded.last_sync_at,
                 last_full_sync_at=COALESCE(
                     excluded.last_full_sync_at,
                     travelline_archive_state.last_full_sync_at),
                 sync_started_at=NULL, sync_kind=NULL, last_error=NULL""",
            (str(hotel_code), str(property_id), 1 if full else 0, now,
             now if full else None))

    def fail_travelline_archive_sync(self, hotel_code, property_id, error):
        self.con.execute(
            """INSERT INTO travelline_archive_state
               (hotel_code, property_id, full_sync_complete, last_error)
               VALUES (?,?,0,?)
               ON CONFLICT(hotel_code, property_id) DO UPDATE SET
                 sync_started_at=NULL, sync_kind=NULL, last_error=excluded.last_error""",
            (str(hotel_code), str(property_id), str(error or 'Ошибка синхронизации')[:2000]))

    def put_travelline_inventory(self, hotel_code, property_id, inventory):
        payload = json.dumps(inventory or {}, ensure_ascii=False,
                             separators=(',', ':'))
        self.con.execute(
            """INSERT INTO travelline_inventory
               (hotel_code, property_id, fetched_at, inventory_json)
               VALUES (?,?,?,?)
               ON CONFLICT(hotel_code, property_id) DO UPDATE SET
                 fetched_at=excluded.fetched_at,
                 inventory_json=excluded.inventory_json""",
            (str(hotel_code), str(property_id), self._tl_archive_now(), payload))

    def get_travelline_inventory(self, hotel_code, property_id):
        row = self.con.execute(
            """SELECT inventory_json FROM travelline_inventory
               WHERE hotel_code=? AND property_id=?""",
            (str(hotel_code), str(property_id))).fetchone()
        return json.loads(row['inventory_json']) if row else None

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
