# -*- coding: utf-8 -*-
"""Обновление плагинов КОМПАС-3D с GitHub (git на компьютере не нужен).

Запуск:  <корень>\\Python\\pythonw.exe <корень>\\_setup\\update.py --root <корень> [--auto]
  --auto  — тихий режим (автозагрузка Windows): окно только если что-то обновилось или ошибка.

1. Узнаёт последний коммит ветки (api.github.com). Совпадает с <корень>\\version.txt — выход.
2. Скачивает архив ветки (codeload.github.com), распаковывает во временную папку.
3. Копирует папки плагинов (с plugin.ini) и _setup в <корень>.
   .rtw, занятая открытым КОМПАС, переименовывается в *.old (Windows это позволяет)
   и заменяется — новая версия подхватится после перезапуска КОМПАС.
4. Новые плагины подключает в КОМПАС (если он закрыт; иначе — при следующем обновлении).
"""
import argparse
import configparser
import ctypes
import glob
import io
import json
import os
import shutil
import sys
import tempfile
import traceback
import urllib.request
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
TITLE = "Обновление плагинов КОМПАС"
AUTO = False


def msg(text, flags=0x40):
    if AUTO and not flags & 0x10 and not text.startswith("Плагины обновлены"):
        print(text)
        return 1
    return ctypes.windll.user32.MessageBoxW(None, text, TITLE, flags | 0x40000)  # поверх окон


def get(url, timeout=30):
    req = urllib.request.Request(url, headers={"User-Agent": "kompas-plugins-updater"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read()


def config():
    cp = configparser.ConfigParser(inline_comment_prefixes=(";",))
    cp.read(os.path.join(HERE, "update.ini"), encoding="utf-8")
    return cp["update"]["repo"].strip(), cp["update"].get("branch", "main").strip()


def latest_commit(repo, branch):
    data = json.loads(get(f"https://api.github.com/repos/{repo}/commits/{branch}"))
    return data["sha"], data["commit"]["message"].splitlines()[0]


def read_ini(path):
    cp = configparser.ConfigParser(inline_comment_prefixes=(";",))
    cp.read(path, encoding="utf-8")
    return cp["plugin"]


def same_file(a, b):
    if not os.path.isfile(b) or os.path.getsize(a) != os.path.getsize(b):
        return False
    with open(a, "rb") as fa, open(b, "rb") as fb:
        return fa.read() == fb.read()


def put_file(src, dst):
    """Копирует файл; занятый (загруженная КОМПАС .rtw) — переименовывает в .old и кладёт новый."""
    if same_file(src, dst):
        return False
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    try:
        shutil.copy2(src, dst)
    except PermissionError:
        old = dst + ".old"
        n = 1
        while os.path.exists(old):
            try:
                os.remove(old)
            except OSError:
                old = f"{dst}.old{n}"
                n += 1
        os.replace(dst, old)
        shutil.copy2(src, dst)
    return True


def put_tree(src, dst):
    changed = 0
    for base, dirs, files in os.walk(src):
        dirs[:] = [d for d in dirs if d != "__pycache__"]
        for f in files:
            rel = os.path.relpath(os.path.join(base, f), src)
            changed += put_file(os.path.join(base, f), os.path.join(dst, rel))
    return changed


def cleanup_old(root):
    for f in glob.glob(os.path.join(root, "*", "*.old*")):
        try:
            os.remove(f)
        except OSError:
            pass   # ещё занят КОМПАС — удалим в следующий раз


def main():
    global AUTO
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=os.path.dirname(HERE))
    ap.add_argument("--auto", action="store_true")
    ap.add_argument("--force", action="store_true", help="обновить, даже если версия та же")
    a = ap.parse_args()
    AUTO = a.auto
    root = os.path.abspath(a.root)
    cleanup_old(root)

    repo, branch = config()
    ver_file = os.path.join(root, "version.txt")
    current = open(ver_file, encoding="utf-8").read().split()[0] if os.path.isfile(ver_file) else ""
    try:
        sha, title = latest_commit(repo, branch)
    except Exception as e:
        if AUTO:
            return               # нет интернета при входе в Windows — молча
        msg(f"Не удалось связаться с GitHub ({repo}):\n{e}", 0x10)
        return
    pending = os.path.isfile(os.path.join(root, "register_pending.txt"))
    if sha == current and not a.force and not pending:
        msg(f"Плагины уже последней версии ({sha[:7]}).")
        return

    tmp = tempfile.mkdtemp(prefix="kp_update_")
    try:
        zipfile.ZipFile(io.BytesIO(get(f"https://codeload.github.com/{repo}/zip/refs/heads/{branch}", 120))).extractall(tmp)
        src = os.path.join(tmp, os.listdir(tmp)[0])          # <repo>-<branch>\

        updated, new_items, all_items = [], [], []
        for name in sorted(os.listdir(src)):
            ini = os.path.join(src, name, "plugin.ini")
            if not os.path.isfile(ini):
                continue
            p = read_ini(ini)
            dst = os.path.join(root, name)
            is_new = not os.path.isdir(dst)
            if put_tree(os.path.join(src, name), dst):
                updated.append(name)
            item = (p.get("name", name), os.path.join(dst, p["library"]), p.get("folder", "Приложения"))
            all_items.append(item)
            if is_new:
                new_items.append(item)
        put_tree(os.path.join(src, "_setup"), os.path.join(root, "_setup"))
        for f in ("ПРОЧТИ.txt", "Обновить.bat"):
            if os.path.isfile(os.path.join(src, f)):
                put_file(os.path.join(src, f), os.path.join(root, f))

        # подключить новые плагины в КОМПАС (нужно, чтобы КОМПАС был закрыт)
        reg_note = ""
        if new_items or pending:
            sys.path.insert(0, os.path.join(root, "_setup"))
            import install
            if install.kompas_running():
                open(os.path.join(root, "register_pending.txt"), "w", encoding="utf-8").write(
                    "\n".join(i[0] for i in new_items))
                reg_note = ("\n\nНовые плагины появятся в меню после того, как вы закроете КОМПАС "
                            "и снова запустите «Обновить.bat».")
            else:
                reg_note = "\n\n" + "\n".join(install.register_in_kompas(all_items))
                try:
                    os.remove(os.path.join(root, "register_pending.txt"))
                except OSError:
                    pass

        with open(ver_file, "w", encoding="utf-8") as f:
            f.write(f"{sha}\n{title}\n")
        changes = ", ".join(updated) if updated else "служебные файлы"
        msg(f"Плагины обновлены до версии {sha[:7]}:\n«{title}»\n\nОбновлено: {changes}{reg_note}"
            "\n\nЕсли КОМПАС открыт — перезапустите его.")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        msg("Ошибка обновления:\n\n" + traceback.format_exc(), 0x10)
        raise
