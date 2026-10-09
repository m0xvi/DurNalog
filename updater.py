# -*- coding: utf-8 -*-
"""Проверка и установка обновлений с GitHub Releases (только stdlib).

Схема: GitHub API releases/latest -> сравнение версий -> скачивание
TurNalog-win64.zip во временную папку -> проверка целостности ->
BAT-скрипт ждёт завершения приложения, распаковывает поверх
и запускает новую версию.

Файлы пользователя (turnalog_config.json, turnalog_history.db, логи),
лежащие рядом с EXE, обновлением не затрагиваются: архив содержит
только программные файлы, распаковка идёт поверх с перезаписью.
"""
import json
import os
import re
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import zipfile
from dataclasses import dataclass

REPO = "m0xvi/DurNalog"
ASSET_NAME = "TurNalog-win64.zip"
EXE_NAME = "TurNalog.exe"
API_URL = "https://api.github.com/repos/{repo}/releases/latest"
TIMEOUT = 20


class UpdateError(Exception):
    """Понятная пользователю ошибка обновления."""


@dataclass
class ReleaseInfo:
    version: str      # "3.11"
    tag: str          # "v3.11"
    page_url: str
    notes: str
    asset_url: str
    asset_size: int


# ------------------------------------------------------------------
# Версии
# ------------------------------------------------------------------
def parse_version(s):
    """'v3.10' -> (3, 10). Несравнимый мусор -> UpdateError."""
    nums = re.findall(r"\d+", str(s or ""))
    if not nums:
        raise UpdateError(f"Непонятный номер версии: {s!r}")
    return tuple(int(x) for x in nums[:4])


def _pad(a, b):
    n = max(len(a), len(b))
    return a + (0,) * (n - len(a)), b + (0,) * (n - len(b))


def is_newer(latest, current):
    """True, если latest строго новее current ('3.10' > '3.9')."""
    try:
        la, ca = _pad(parse_version(latest), parse_version(current))
    except UpdateError:
        return False
    return la > ca


# ------------------------------------------------------------------
# GitHub API
# ------------------------------------------------------------------
def _http_get(url, timeout=TIMEOUT, opener=None):
    req = urllib.request.Request(url, headers={
        "User-Agent": "TurNalog-updater",
        "Accept": "application/vnd.github+json",
    })
    do = opener or urllib.request.urlopen
    try:
        return do(req, timeout=timeout)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            raise UpdateError("Релизы не найдены (репозиторий недоступен).")
        if e.code == 403:
            raise UpdateError("GitHub отклонил запрос (лимит). Попробуйте позже.")
        raise UpdateError(f"GitHub вернул ошибку HTTP {e.code}.")
    except urllib.error.URLError as e:
        raise UpdateError(f"Нет соединения с GitHub: {e.reason}")
    except TimeoutError:
        raise UpdateError("GitHub не отвечает (таймаут).")
    except UpdateError:
        raise
    except Exception as e:  # noqa: BLE001 — любая сетевая ошибка -> текст
        raise UpdateError(f"Сетевая ошибка: {e}")


def pick_asset(assets):
    """Выбор файла из релиза: точное имя, иначе первый .zip, иначе None."""
    for a in assets or []:
        if a.get("name") == ASSET_NAME:
            return a
    for a in assets or []:
        if str(a.get("name", "")).lower().endswith(".zip"):
            return a
    return None


def fetch_latest(repo=REPO, timeout=TIMEOUT, opener=None):
    """Метаданные последнего релиза -> ReleaseInfo (или UpdateError)."""
    resp = _http_get(API_URL.format(repo=repo), timeout=timeout, opener=opener)
    try:
        try:
            data = json.load(resp)
        except Exception:
            raise UpdateError("GitHub вернул некорректный ответ.")
    finally:
        try:
            resp.close()
        except Exception:
            pass
    tag = str(data.get("tag_name") or "")
    if not tag:
        raise UpdateError("У релиза нет номера версии.")
    asset = pick_asset(data.get("assets"))
    if not asset or not asset.get("browser_download_url"):
        raise UpdateError(f"В релизе {tag} нет файла {ASSET_NAME}.")
    try:
        version = ".".join(str(x) for x in parse_version(tag))
    except UpdateError:
        raise UpdateError(f"Непонятный номер версии релиза: {tag}")
    return ReleaseInfo(
        version=version, tag=tag,
        page_url=str(data.get("html_url") or
                     f"https://github.com/{repo}/releases"),
        notes=str(data.get("body") or ""),
        asset_url=asset["browser_download_url"],
        asset_size=int(asset.get("size") or 0))


def check_for_updates(current_version, repo=REPO, timeout=TIMEOUT, opener=None):
    """ReleaseInfo, если на GitHub версия новее текущей, иначе None."""
    rel = fetch_latest(repo=repo, timeout=timeout, opener=opener)
    return rel if is_newer(rel.version, current_version) else None


# ------------------------------------------------------------------
# Скачивание и проверка
# ------------------------------------------------------------------
def download_release(info, dest_dir, progress=None, timeout=120, opener=None):
    """Скачивает asset релиза. progress(done, total). Возвращает путь к zip."""
    os.makedirs(dest_dir, exist_ok=True)
    dest = os.path.join(dest_dir, ASSET_NAME)
    tmp = dest + ".part"
    resp = _http_get(info.asset_url, timeout=timeout, opener=opener)
    total = info.asset_size or 0
    try:
        hdr = resp.getheader("Content-Length") if hasattr(resp, "getheader") else None
    except Exception:
        hdr = None
    if hdr:
        try:
            total = int(hdr)
        except ValueError:
            pass
    done = 0
    try:
        with open(tmp, "wb") as f:
            while True:
                chunk = resp.read(1024 * 256)
                if not chunk:
                    break
                f.write(chunk)
                done += len(chunk)
                if progress:
                    progress(done, total)
    except UpdateError:
        raise
    except Exception as e:  # noqa: BLE001
        raise UpdateError(f"Скачивание прервано: {e}")
    finally:
        try:
            resp.close()
        except Exception:
            pass
    if total and done != total:
        try:
            os.remove(tmp)
        except OSError:
            pass
        raise UpdateError("Файл скачался не полностью. Попробуйте ещё раз.")
    os.replace(tmp, dest)
    return dest


def verify_package(path):
    """Целостность zip + наличие TurNalog/TurNalog.exe. Иначе UpdateError."""
    try:
        with zipfile.ZipFile(path) as z:
            bad = z.testzip()
            if bad is not None:
                raise UpdateError(f"Архив повреждён (файл {bad}). Скачайте заново.")
            names = z.namelist()
    except zipfile.BadZipFile:
        raise UpdateError("Скачанный файл — не архив. Скачайте заново.")
    except UpdateError:
        raise
    except Exception as e:  # noqa: BLE001
        raise UpdateError(f"Не удалось проверить архив: {e}")
    if f"TurNalog/{EXE_NAME}" not in names:
        raise UpdateError("В архиве нет TurNalog.exe — неверный файл обновления.")


# ------------------------------------------------------------------
# Установка (Windows): скрипт ставится поверх после выхода
# ------------------------------------------------------------------
def is_frozen():
    """True — запущены как собранный EXE (обновление применимо)."""
    return bool(getattr(sys, "frozen", False))


def app_dir():
    """Папка приложения: рядом с EXE (сборка) или с исходниками (разработка)."""
    if is_frozen():
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def staging_dir(tag):
    """Папка для загрузки релиза."""
    base = os.path.join(tempfile.gettempdir(), "TurNalog_update", tag)
    os.makedirs(base, exist_ok=True)
    return base


def build_install_script(zip_path, target_dir, exe_name=EXE_NAME):
    """BAT-скрипт: ждёт выхода EXE, распаковывает архив поверх,
    запускает новую версию, удаляет временные файлы и сам себя."""
    tmp_dir = os.path.join(tempfile.gettempdir(), "TurNalog_update", "_x")
    lines = [
        "@echo off",
        "chcp 65001 >nul",
        f'set "TL_ZIP={zip_path}"',
        f'set "TL_DST={target_dir}"',
        f'set "TL_TMP={tmp_dir}"',
        "timeout /t 2 /nobreak >nul",
        ":waitloop",
        f'tasklist /FI "IMAGENAME eq {exe_name}" 2>nul | find /I "{exe_name}" >nul',
        "if not errorlevel 1 (timeout /t 1 /nobreak >nul & goto waitloop)",
        'if exist "%TL_TMP%" rmdir /S /Q "%TL_TMP%"',
        "powershell -NoProfile -ExecutionPolicy Bypass -Command "
        "\"$ErrorActionPreference='Stop'; "
        "Expand-Archive -LiteralPath \\\"$env:TL_ZIP\\\" "
        "-DestinationPath \\\"$env:TL_TMP\\\" -Force; "
        "Copy-Item -Path \\\"$env:TL_TMP\\TurNalog\\*\\\" "
        "-Destination \\\"$env:TL_DST\\\" -Recurse -Force\"",
        "if errorlevel 1 (echo UPDATE FAILED: see errors above & pause & exit /b 1)",
        f'start "" /D "%TL_DST%" "%TL_DST%\\{exe_name}"',
        'del "%TL_ZIP%" 2>nul',
        'rmdir /S /Q "%TL_TMP%" 2>nul',
        "(goto) 2>nul & del \"%~f0\"",
    ]
    return "\r\n".join(lines) + "\r\n"


def plan_install(zip_path, target_dir=None, exe_name=EXE_NAME):
    """Пишет BAT рядом с архивом. Возвращает {'script', 'exe'}."""
    target = os.path.abspath(target_dir or app_dir())
    script = os.path.join(os.path.dirname(os.path.abspath(zip_path)),
                          "install_update.bat")
    with open(script, "w", encoding="utf-8", newline="") as f:
        f.write(build_install_script(os.path.abspath(zip_path), target, exe_name))
    return {"script": script, "exe": os.path.join(target, exe_name)}


def launch_install(plan):
    """Запускает BAT независимо от приложения (дальше приложение
    должно закрыться, чтобы файлы разблокировались)."""
    if os.name != "nt":
        raise UpdateError("Автоустановка поддерживается только в Windows.")
    subprocess.Popen(["cmd", "/c", plan["script"]])  # noqa: S603,S607 — свой скрипт
