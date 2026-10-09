# -*- coding: utf-8 -*-
import io
import json
import os
import sys
import urllib.error
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import updater
from updater import (UpdateError, check_for_updates, download_release,
                     fetch_latest, is_newer, launch_install, parse_version,
                     pick_asset, plan_install, verify_package)


def test_versions():
    assert parse_version("v3.10") == (3, 10)
    assert parse_version("3.9") == (3, 9)
    assert is_newer("v3.10", "3.9")
    assert is_newer("3.10", "v3.9")
    assert not is_newer("v3.9", "3.10")
    assert not is_newer("v3.10", "3.10")
    assert not is_newer("v3.2", "3.10")
    assert is_newer("v4.0", "3.10")
    try:
        parse_version("noversion")
    except UpdateError:
        pass
    else:
        raise AssertionError("мусор обязан дать UpdateError")
    assert not is_newer("v3.11", "мусор")  # несравнимые версии — не новее


def test_pick_asset():
    za = {"name": "TurNalog-win64.zip", "browser_download_url": "http://x/z"}
    other = {"name": "notes.txt"}
    assert pick_asset([other, za]) is za
    assert pick_asset([{"name": "a.zip"}, {"name": "b.zip"}])["name"] == "a.zip"
    assert pick_asset([other]) is None
    assert pick_asset([]) is None
    assert pick_asset(None) is None


class FakeResp:
    """Минимальный urlopen-ответ: read()/getheader()/close()."""

    def __init__(self, data=b"", headers=None):
        self._buf = io.BytesIO(data)
        self._headers = headers or {}

    def read(self, n=-1):
        return self._buf.read(n)

    def getheader(self, name):
        return self._headers.get(name)

    def close(self):
        pass


def _release_json(**kw):
    d = {"tag_name": "v3.11", "html_url": "https://github.com/r/3.11",
         "body": "notes",
         "assets": [{"name": "TurNalog-win64.zip",
                     "browser_download_url": "https://x/dl.zip", "size": 10}]}
    d.update(kw)
    return json.dumps(d).encode()


def test_fetch_latest_ok():
    rel = fetch_latest(opener=lambda req, timeout=None: FakeResp(_release_json()))
    assert (rel.version, rel.tag) == ("3.11", "v3.11")
    assert rel.asset_url == "https://x/dl.zip" and rel.asset_size == 10


def test_fetch_latest_errors():
    bad_json = lambda req, timeout=None: FakeResp(b"not json")
    try:
        fetch_latest(opener=bad_json)
    except UpdateError:
        pass
    else:
        raise AssertionError("битый JSON обязан дать UpdateError")

    no_assets = lambda req, timeout=None: FakeResp(_release_json(assets=[]))
    try:
        fetch_latest(opener=no_assets)
    except UpdateError:
        pass
    else:
        raise AssertionError("релиз без ассета обязан дать UpdateError")

    def http404(req, timeout=None):
        raise urllib.error.HTTPError(req.full_url, 404, "x", {}, None)

    try:
        fetch_latest(opener=http404)
    except UpdateError as e:
        assert "не найдены" in str(e)
    else:
        raise AssertionError("404 обязан дать UpdateError")

    def no_net(req, timeout=None):
        raise urllib.error.URLError("boom")

    try:
        fetch_latest(opener=no_net)
    except UpdateError as e:
        assert "boom" in str(e)
    else:
        raise AssertionError("URLError обязан дать UpdateError")


def test_check_for_updates():
    ok = lambda req, timeout=None: FakeResp(_release_json())
    assert check_for_updates("3.10", opener=ok).tag == "v3.11"
    assert check_for_updates("3.11", opener=ok) is None
    assert check_for_updates("3.12", opener=ok) is None


def test_download_ok_and_progress(tmp_path):
    data = os.urandom(600 * 1024)
    seen = []

    def opener(req, timeout=None):
        return FakeResp(data, {"Content-Length": str(len(data))})

    info = updater.ReleaseInfo("3.11", "v3.11", "u", "", "https://x/dl.zip", len(data))
    p = download_release(info, str(tmp_path), progress=lambda d, t: seen.append((d, t)),
                         opener=opener)
    assert open(p, "rb").read() == data
    assert seen and seen[-1] == (len(data), len(data))
    assert [d for d, _ in seen] == sorted(d for d, _ in seen)  # монотонный прогресс


def test_download_incomplete(tmp_path):
    def opener(req, timeout=None):
        return FakeResp(b"short", {"Content-Length": "1000"})

    info = updater.ReleaseInfo("3.11", "v3.11", "u", "", "https://x/dl.zip", 1000)
    try:
        download_release(info, str(tmp_path), opener=opener)
    except UpdateError as e:
        assert "не полностью" in str(e)
    else:
        raise AssertionError("оборванное скачивание обязано дать UpdateError")


def _mk_zip(path, names):
    with zipfile.ZipFile(path, "w") as z:
        for n in names:
            z.writestr(n, b"x")
    return path


def test_verify_package(tmp_path):
    good = _mk_zip(str(tmp_path / "g.zip"), ["TurNalog/TurNalog.exe", "TurNalog/a.dll"])
    assert verify_package(good) is None
    bad = _mk_zip(str(tmp_path / "b.zip"), ["readme.txt"])
    try:
        verify_package(bad)
    except UpdateError as e:
        assert "TurNalog.exe" in str(e)
    else:
        raise AssertionError("архив без EXE обязан дать UpdateError")
    corrupt = str(tmp_path / "c.zip")
    open(corrupt, "wb").write(b"definitely not a zip")
    try:
        verify_package(corrupt)
    except UpdateError:
        pass
    else:
        raise AssertionError("битый архив обязан дать UpdateError")


def test_install_plan(tmp_path):
    zp = os.path.join(str(tmp_path), "TurNalog-win64.zip")
    open(zp, "wb").write(b"fake")
    plan = plan_install(zp, target_dir=os.path.join(str(tmp_path), "app"))
    assert os.path.exists(plan["script"])
    bat = open(plan["script"], encoding="utf-8").read()
    assert "tasklist" in bat and "TurNalog.exe" in bat
    assert "Expand-Archive" in bat and "Copy-Item" in bat
    assert "TurNalog\\*" in bat and "start" in bat and 'del "%~f0"' in bat
    assert zp in bat and plan["exe"].endswith("TurNalog.exe")


def test_launch_install_not_windows():
    if os.name == "nt":
        return  # на Windows не запускаем настоящий cmd в тестах
    try:
        launch_install({"script": "x.bat", "exe": "x.exe"})
    except UpdateError as e:
        assert "Windows" in str(e)
    else:
        raise AssertionError("вне Windows обязан дать UpdateError")
