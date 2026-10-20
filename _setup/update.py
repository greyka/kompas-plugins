# -*- coding: utf-8 -*-
"""Обновление плагинов КОМПАС-3D с GitHub (git на компьютере не нужен).

Запуск:  <корень>\\Python\\pythonw.exe <корень>\\_setup\\update.py --root <корень> [--auto]
  --auto   — автозагрузка Windows: молча, окно только если обновилось или ошибка;
             ничего не делает, если в настройках выключено «Проверять обновления».
  --check  — только узнать версии (JSON в stdout) — для «Менеджера плагинов».
  --quiet  — без окон (вывод в консоль).
  --force  — поставить заново, даже если версия та же.

Канал (настройки → Обновление → «Что устанавливать»):
  releases — только выпущенные версии (релизы GitHub, тег vX.Y.Z) — по умолчанию;
  main     — последние изменения ветки main.

Что делает: скачивает архив версии, копирует папки плагинов (с plugin.ini), _setup и файлы
корня (VERSION, CHANGELOG.md, менеджер, …) в <корень>. Файл настроек settings.ini не трогает.
.rtw / .exe, занятые открытым КОМПАС или менеджером, переименовываются в *.old и заменяются —
новая версия заработает после перезапуска. Новые плагины подключаются в КОМПАС, если он закрыт.
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
import time
import traceback
import urllib.error
import urllib.request
import zipfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import net  # noqa: E402  HTTPS с запасными путями

HERE = os.path.dirname(os.path.abspath(__file__))
TITLE = "Обновление плагинов КОМПАС"
ROOT_FILES = ("VERSION", "CHANGELOG.md", "ПРОЧТИ.txt", "Обновить.bat", "Менеджер плагинов КОМПАС.exe")
AUTO = QUIET = False


def msg(text, flags=0x40):
    if QUIET or (AUTO and not flags & 0x10 and not text.startswith("Плагины обновлены")):
        print(text)
        return 1
    return ctypes.windll.user32.MessageBoxW(None, text, TITLE, flags | 0x40000)  # поверх окон


def get(url, timeout=30):
    return net.get(url, timeout)     # сертификаты Windows + certifi, при ошибке SSL — curl.exe


def repo_config():
    cp = configparser.ConfigParser(inline_comment_prefixes=(";",))
    cp.read(os.path.join(HERE, "update.ini"), encoding="utf-8")
    return cp["update"]["repo"].strip(), cp["update"].get("branch", "main").strip()


def user_setting(root, key, default):
    cp = configparser.ConfigParser(interpolation=None, inline_comment_prefixes=(";",))
    cp.read(os.path.join(root, "settings.ini"), encoding="utf-8-sig")
    return cp.get("update", key, fallback=default).strip() or default


def installed(root):
    """(id, версия) установленной копии: id — тег или коммит из version.txt."""
    vid = ""
    if os.path.isfile(os.path.join(root, "version.txt")):
        vid = open(os.path.join(root, "version.txt"), encoding="utf-8").read().split("\n")[0].strip()
    ver = "?"
    if os.path.isfile(os.path.join(root, "VERSION")):
        ver = open(os.path.join(root, "VERSION"), encoding="utf-8").read().strip()
    return vid, ver


LOG_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "KompasPlugins", "logs")


def log(text):
    """Журнал обновлений — его можно отправить кнопкой «Сообщить о проблеме…» в менеджере."""
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        with open(os.path.join(LOG_DIR, "Обновление.log"), "a", encoding="utf-8") as f:
            f.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')}  {text}\n")
    except OSError:
        pass


def latest(repo, branch, channel):
    """Последняя версия: через API GitHub, а если он недоступен (лимит запросов, блокировка
    api.github.com у провайдера) — через обычные страницы github.com / raw.githubusercontent.com."""
    try:
        return latest_api(repo, branch, channel)
    except Exception as e_api:
        log(f"api.github.com: {type(e_api).__name__}: {e_api} — пробую через github.com")
        try:
            return latest_web(repo, branch, channel)
        except Exception as e_web:
            log(f"github.com: {type(e_web).__name__}: {e_web}")
            raise RuntimeError(f"api.github.com: {e_api}\ngithub.com: {e_web}") from None


def changelog_section(text, ver):
    import re
    m = re.search(rf"^## {re.escape(ver)}\b.*?$\n(.*?)(?=^## |\Z)", text, re.S | re.M)
    return m.group(1).strip() if m else ""


def latest_web(repo, branch, channel):
    if channel == "releases":
        final = net.final_url(f"https://github.com/{repo}/releases/latest")   # …/releases/tag/v1.3.0
        if "/releases/tag/" in final:
            tag = final.rsplit("/", 1)[-1]
            ver = tag.lstrip("vV")
            try:
                notes = changelog_section(
                    get(f"https://raw.githubusercontent.com/{repo}/{tag}/CHANGELOG.md").decode("utf-8"), ver)
            except Exception:
                notes = ""
            return dict(id=tag, version=ver, title=f"Версия {ver}", notes=notes, channel="releases",
                        zip=f"https://codeload.github.com/{repo}/zip/refs/tags/{tag}")
    ver = get(f"https://raw.githubusercontent.com/{repo}/{branch}/VERSION").decode("utf-8").strip()
    return dict(id=f"{branch}@{ver}", version=f"{ver} ({branch})", title=f"Ветка {branch}, версия {ver}",
                notes="", channel="main", zip=f"https://codeload.github.com/{repo}/zip/refs/heads/{branch}")


def latest_api(repo, branch, channel):
    """→ dict(id, version, title, notes, zip, channel). Нет релизов — берём ветку."""
    if channel == "releases":
        try:
            r = json.loads(get(f"https://api.github.com/repos/{repo}/releases/latest"))
            tag = r["tag_name"]
            return dict(id=tag, version=tag.lstrip("vV"), title=r.get("name") or tag,
                        notes=(r.get("body") or "").strip(), channel="releases",
                        zip=f"https://codeload.github.com/{repo}/zip/refs/tags/{tag}")
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
    c = json.loads(get(f"https://api.github.com/repos/{repo}/commits/{branch}"))
    sha = c["sha"]
    try:
        ver = get(f"https://raw.githubusercontent.com/{repo}/{sha}/VERSION").decode("utf-8").strip()
    except Exception:
        ver = "?"
    return dict(id=sha, version=f"{ver} ({sha[:7]})", title=c["commit"]["message"].split("\n")[0],
                notes="\n".join(l for l in c["commit"]["message"].strip().split("\n")
                                if not l.lower().startswith("co-authored-by")).strip(), channel="main",
                zip=f"https://codeload.github.com/{repo}/zip/{sha}")


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
    """Копирует файл; занятый (загруженная .rtw, запущенный .exe) — в .old и кладёт новый."""
    if same_file(src, dst):
        return False
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    try:
        shutil.copy2(src, dst)
    except PermissionError:
        old, n = dst + ".old", 1
        while os.path.exists(old):
            try:
                os.remove(old)
            except OSError:
                old, n = f"{dst}.old{n}", n + 1
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
    for f in glob.glob(os.path.join(root, "*.old*")) + glob.glob(os.path.join(root, "*", "*.old*")):
        try:
            os.remove(f)
        except OSError:
            pass   # ещё занят — удалим в следующий раз


def install(root, info):
    tmp = tempfile.mkdtemp(prefix="kp_update_")
    try:
        zipfile.ZipFile(io.BytesIO(get(info["zip"], 180))).extractall(tmp)
        src = os.path.join(tmp, os.listdir(tmp)[0])          # <repo>-<версия>\
        pending_file = os.path.join(root, "register_pending.txt")
        pending = os.path.isfile(pending_file)

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
        for f in ROOT_FILES:
            if os.path.isfile(os.path.join(src, f)):
                put_file(os.path.join(src, f), os.path.join(root, f))

        sys.path.insert(0, os.path.join(root, "_setup"))
        import install as inst
        try:
            inst.install_updater(root)   # ярлыки менеджера и автообновления (могли появиться в новой версии)
        except Exception:
            traceback.print_exc()

        reg_note = ""
        if new_items or pending:
            if inst.kompas_running():
                with open(pending_file, "w", encoding="utf-8") as fh:
                    fh.write("\n".join(i[0] for i in new_items))
                reg_note = ("\n\nНовые плагины появятся в меню КОМПАС после того, как вы закроете "
                            "КОМПАС и ещё раз нажмёте «Обновить».")
            else:
                reg_note = "\n\n" + "\n".join(inst.register_in_kompas(all_items))
                if os.path.exists(pending_file):
                    os.remove(pending_file)

        with open(os.path.join(root, "version.txt"), "w", encoding="utf-8") as f:
            f.write(f"{info['id']}\n{info['version']}\n{info['title']}\n")
        return updated, reg_note
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def main():
    global AUTO, QUIET
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=os.path.dirname(HERE))
    ap.add_argument("--auto", action="store_true")
    ap.add_argument("--check", action="store_true")
    ap.add_argument("--force", action="store_true")
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()
    AUTO, QUIET = a.auto, a.quiet or a.check
    root = os.path.abspath(a.root)
    if a.auto and user_setting(root, "auto", "true").lower() not in ("1", "true", "yes", "да", "on"):
        return
    cleanup_old(root)

    repo, branch = repo_config()
    channel = user_setting(root, "channel", "releases")
    cur_id, cur_ver = installed(root)
    try:
        info = latest(repo, branch, channel)
    except Exception as e:
        if a.check:
            print(json.dumps(dict(error=str(e), installed_version=cur_ver), ensure_ascii=False))
        elif not AUTO:
            msg(f"Не удалось связаться с GitHub ({repo}):\n{e}", 0x10)
        return
    pending = os.path.isfile(os.path.join(root, "register_pending.txt"))
    available = info["id"] != cur_id

    if a.check:
        print(json.dumps(dict(installed_version=cur_ver, installed_id=cur_id,
                              latest_version=info["version"], latest_id=info["id"],
                              title=info["title"], notes=info["notes"], channel=info["channel"],
                              update_available=available, register_pending=pending,
                              repo=repo), ensure_ascii=False))
        return
    if not available and not a.force and not pending:
        msg(f"Установлена последняя версия: {cur_ver}.")
        return

    updated, reg_note = install(root, info)
    changes = ", ".join(updated) if updated else "служебные файлы"
    notes = f"\n\n{info['notes'][:1500]}" if info["notes"] and info["notes"] != info["title"] else ""
    msg(f"Плагины обновлены: {cur_ver} → {info['version']}\n«{info['title']}»{notes}"
        f"\n\nОбновлено: {changes}{reg_note}\n\nЕсли КОМПАС открыт — перезапустите его.")


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log("ошибка:\n" + traceback.format_exc())
        msg("Ошибка обновления:\n\n" + traceback.format_exc(), 0x10)
        raise
