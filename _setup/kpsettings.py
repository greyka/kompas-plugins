# -*- coding: utf-8 -*-
"""Общие настройки и версия плагинов КОМПАС.

Файл настроек: <корень>\\settings.ini (UTF-8). Его правит «Менеджер плагинов КОМПАС.exe»,
можно и Блокнотом. Чего нет в файле — берётся по умолчанию из settings_schema.json.
Обновление с GitHub файл настроек не трогает (его нет в репозитории).

    import kpsettings
    S = kpsettings.load()
    S.get("frame", "pdf")      # → True
    kpsettings.version()       # → "1.1.0"
"""
import configparser
import json
import os
import subprocess

SETUP = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SETUP)                     # C:\KompasPlugins (или папка «Плагины»)
SETTINGS_FILE = os.path.join(ROOT, "settings.ini")
SCHEMA_FILE = os.path.join(SETUP, "settings_schema.json")
MANAGER_EXE = os.path.join(ROOT, "Менеджер плагинов КОМПАС.exe")


def schema():
    with open(SCHEMA_FILE, encoding="utf-8") as f:
        return json.load(f)["sections"]


def version():
    try:
        with open(os.path.join(ROOT, "VERSION"), encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return "?"


class Settings:
    def __init__(self):
        self.defs = {(s["id"], it["key"]): it for s in schema() for it in s["items"]}
        self.cp = configparser.ConfigParser(interpolation=None, inline_comment_prefixes=(";",))
        if os.path.isfile(SETTINGS_FILE):
            self.cp.read(SETTINGS_FILE, encoding="utf-8-sig")

    def get(self, section, key):
        d = self.defs.get((section, key))
        if d is None:
            raise KeyError(f"нет настройки {section}.{key} в settings_schema.json")
        raw = self.cp.get(section, key, fallback=None)
        if raw is None or raw.strip() == "":
            return d["default"] if d["type"] != "text" or raw is None else ""
        raw = raw.strip()
        try:
            if d["type"] == "bool":
                return raw.lower() in ("1", "true", "yes", "да", "on")
            if d["type"] == "int":
                return int(raw)
            if d["type"] == "float":
                return float(raw.replace(",", "."))
        except ValueError:
            return d["default"]
        if d["type"] == "choice" and raw not in [c[0] for c in d["choices"]]:
            return d["default"]
        return raw

    def sides(self, section):
        """'both' | 'left' | 'right' → 'left,right' (формат ключа --sides)."""
        v = self.get(section, "sides")
        return {"both": "left,right"}.get(v, v)


def load():
    return Settings()


def open_manager(tab=None):
    """Открыть окно «Менеджер плагинов» (из пункта меню «Настройки…» в КОМПАС)."""
    args = [MANAGER_EXE] + ([f"--tab={tab}"] if tab else [])
    if os.path.isfile(MANAGER_EXE):
        subprocess.Popen(args, cwd=ROOT)
        return True
    os.startfile(SETTINGS_FILE if os.path.isfile(SETTINGS_FILE) else ROOT)
    return False
