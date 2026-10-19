# -*- coding: utf-8 -*-
"""Журнал работы плагинов КОМПАС и отчёт об ошибке разработчику в Telegram.

Перенесено из плагинов Archicad (kd/report.py), бот тот же — @lider_house_debug_bot.
Журнал: %LOCALAPPDATA%\\KompasPlugins\\logs\\<плагин>.log — туда же идёт всё, что плагин печатает.
Отчёт: бот из «_setup\\update.ini», раздел [report] (bot_token, chat_id) — одно сообщение:
файл (заголовок + журнал) с подписью. Каждый отчёт дописывается в закреплённый в чате
«Реестр отчётов» — разработчик забирает отчёты сам, без пересылки (--inbox).

Окно «Сообщить о проблеме» — в «Менеджер плагинов КОМПАС.exe» (--report): в переносном
Python нет tkinter. Скрипт плагина при ошибке вызывает error_dialog().

    python report.py --test                          проверить отправку
    python report.py --send --plugin X [--error-file F] [--comment-file C]   (зовёт менеджер)
    python report.py --inbox <папка>                 забрать отчёты КОМПАС (для разработчика)
"""
import configparser
import getpass
import io
import json
import os
import platform
import subprocess
import sys
import tempfile
import time
import traceback
import urllib.request
import uuid

SETUP = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(SETUP)
LOG_DIR = os.path.join(os.environ.get("LOCALAPPDATA") or os.path.expanduser("~"), "KompasPlugins", "logs")
MAX_LOG = 2_000_000              # байт; больше — старый журнал уходит в .1
SYSTEM = "КОМПАС-3D"             # метка в отчёте: бот общий с плагинами Archicad
_state = {"plugin": "plugins", "log": None, "document": ""}


def version():
    try:
        with open(os.path.join(ROOT, "VERSION"), encoding="utf-8") as f:
            return f.read().strip()
    except OSError:
        return "?"


# ---------- журнал ----------

class _Tee(io.TextIOBase):
    """Пишет и в исходный поток (у pythonw его нет), и в журнал."""

    def __init__(self, orig, f):
        self.orig, self.f = orig, f

    def write(self, s):
        try:
            self.f.write(s)
            self.f.flush()
        except Exception:
            pass
        if self.orig is not None:
            try:
                self.orig.write(s)
            except Exception:
                pass
        return len(s)

    def flush(self):
        for x in (self.orig, self.f):
            try:
                x and x.flush()
            except Exception:
                pass


def log_path(plugin=None):
    return os.path.join(LOG_DIR, f"{plugin or _state['plugin']}.log")


def start_log(plugin):
    """Включает журнал плагина; вызывать первым делом в main()."""
    _state["plugin"] = plugin
    path = log_path(plugin)
    try:
        os.makedirs(LOG_DIR, exist_ok=True)
        if os.path.isfile(path) and os.path.getsize(path) > MAX_LOG:
            os.replace(path, path + ".1")
        f = open(path, "a", encoding="utf-8")
    except OSError:
        return path
    _state["log"] = path
    f.write(f"\n===== {time.strftime('%Y-%m-%d %H:%M:%S')}  {plugin}  v{version()}  "
            f"{platform.node()}\\{getpass.getuser()}  argv={sys.argv[1:]}\n")
    f.flush()
    sys.stdout, sys.stderr = _Tee(sys.stdout, f), _Tee(sys.stderr, f)
    return path


def set_document(name):
    """Текущий документ КОМПАС — попадает в журнал и в отчёт."""
    _state["document"] = name or ""
    print(f"Документ: {name}")


# ---------- окно ошибки (через менеджер) ----------

def error_dialog(title, text):
    """Окно «Произошла ошибка» с кнопкой «Отправить разработчику» (окно — в менеджере).
    Нет менеджера — обычное окно с текстом ошибки."""
    print(text, file=sys.stderr)
    exe = os.path.join(ROOT, "Менеджер плагинов КОМПАС.exe")
    try:
        if os.path.isfile(exe):
            fd, path = tempfile.mkstemp(prefix="kp_error_", suffix=".txt")
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                f.write(text)
            subprocess.Popen([exe, "--report", f"--plugin={_state['plugin']}", f"--error={path}",
                              f"--title={title}"], cwd=ROOT)
            return
    except Exception:
        traceback.print_exc()
    import ctypes
    ctypes.windll.user32.MessageBoxW(None, text[-3000:], title, 0x10)


# ---------- отправка ----------

def _settings():
    cp = configparser.ConfigParser()
    try:
        cp.read(os.path.join(SETUP, "update.ini"), encoding="utf-8-sig")
    except Exception:
        pass
    return (cp.get("report", "bot_token", fallback="").strip(),
            cp.get("report", "chat_id", fallback="").strip())


def configured():
    return all(_settings())


def _api(token, method, fields, file=None, timeout=30, field="document"):
    url = f"https://api.telegram.org/bot{token}/{method}"
    if file is None:
        req = urllib.request.Request(url, json.dumps(fields).encode("utf-8"), {"Content-Type": "application/json"})
    else:
        b = uuid.uuid4().hex
        body = io.BytesIO()
        for k, v in fields.items():
            body.write(f'--{b}\r\nContent-Disposition: form-data; name="{k}"\r\n\r\n{v}\r\n'.encode("utf-8"))
        name, blob = file
        body.write(f'--{b}\r\nContent-Disposition: form-data; name="{field}"; filename="{name}"\r\n'
                   f"Content-Type: text/plain\r\n\r\n".encode("utf-8"))
        body.write(blob)
        body.write(f"\r\n--{b}--\r\n".encode("utf-8"))
        req = urllib.request.Request(url, body.getvalue(), {"Content-Type": f"multipart/form-data; boundary={b}"})
    res = json.loads(urllib.request.urlopen(req, timeout=timeout).read().decode("utf-8"))
    if not res.get("ok"):
        raise RuntimeError(res.get("description") or "Telegram: ошибка")
    return res


def _tail(path, limit=400_000):
    try:
        with open(path, "rb") as f:
            f.seek(0, 2)
            n = f.tell()
            f.seek(max(0, n - limit))
            return f.read()
    except OSError:
        return b""


def header(title):
    return "\n".join([
        f"[{SYSTEM}] {title}",
        f"Плагин: {_state['plugin']}  v{version()}",
        f"Документ: {_state['document'] or '—'}",
        f"Компьютер: {platform.node()}  пользователь: {getpass.getuser()}",
        f"Windows {platform.release()} ({platform.version()}), Python {platform.python_version()}",
        f"Время: {time.strftime('%Y-%m-%d %H:%M:%S')}",
    ])


def send(title, error="", comment="", log=None):
    """Отправляет отчёт; → текст для пользователя. Исключение — не удалось."""
    token, chat = _settings()
    if not (token and chat):
        raise RuntimeError("Отправка не настроена: в «_setup\\update.ini» нет раздела [report].")
    text = header(title)
    if comment.strip():
        text += f"\n\nКомментарий: {comment.strip()}"
    if error.strip():
        err = error.strip()
        text += "\n\nОшибка:\n" + (err if len(err) < 2500 else "…" + err[-2500:])
    log = log or _state["log"] or log_path()
    full = text.encode("utf-8") + b"\n\n" + ("=" * 60).encode() + b"\n" + _tail(log)
    caption = text if len(text) <= 1000 else text[:990] + "\n…(полностью — в файле)"
    fname = f"KOMPAS_{_state['plugin']}_{time.strftime('%Y%m%d_%H%M%S')}.log".replace(" ", "_")
    msg = _api(token, "sendDocument", {"chat_id": chat, "caption": caption}, (fname, full), timeout=60)["result"]
    try:
        _register(token, chat, msg, text)
    except Exception as e:                      # реестр — для разбора; сам отчёт уже ушёл
        print(f"Реестр отчётов: {e}")
    return "Отчёт отправлен разработчику."


INDEX_NAME = "reports_index.jsonl"


def _download(token, file_id):
    path = _api(token, "getFile", {"file_id": file_id})["result"]["file_path"]
    return urllib.request.urlopen(f"https://api.telegram.org/file/bot{token}/{path}", timeout=60).read()


def _pinned_index(token, chat):
    """→ (message_id, байты реестра) закреплённого реестра или (None, b"")."""
    pin = _api(token, "getChat", {"chat_id": chat})["result"].get("pinned_message") or {}
    doc = pin.get("document") or {}
    if doc.get("file_name") == INDEX_NAME:
        return pin["message_id"], _download(token, doc["file_id"])
    return None, b""


def _register(token, chat, msg, text):
    """Дописывает отчёт в закреплённый файл-реестр (общий с плагинами Archicad):
    свои сообщения бот не читает, а закреплённое — может (getChat)."""
    entry = {"message_id": msg["message_id"], "date": msg["date"], "file_id": msg["document"]["file_id"],
             "file_name": msg["document"].get("file_name"), "text": text, "system": SYSTEM}
    mid, blob = _pinned_index(token, chat)
    blob += (json.dumps(entry, ensure_ascii=False) + "\n").encode("utf-8")
    if mid:
        _api(token, "editMessageMedia",
             {"chat_id": chat, "message_id": mid,
              "media": json.dumps({"type": "document", "media": "attach://index",
                                   "caption": "Реестр отчётов (служебное, не откреплять)"})},
             (INDEX_NAME, blob), field="index")
    else:
        new = _api(token, "sendDocument", {"chat_id": chat, "disable_notification": "true",
                                           "caption": "Реестр отчётов (служебное, не откреплять)"},
                   (INDEX_NAME, blob))["result"]
        _api(token, "pinChatMessage", {"chat_id": chat, "message_id": new["message_id"],
                                       "disable_notification": True})


def _is_kompas(entry):
    return entry.get("system") == SYSTEM or f"[{SYSTEM}]" in entry.get("text", "") \
        or (entry.get("file_name") or "").startswith("KOMPAS_")


def inbox(out_dir):
    """Забирает новые отчёты КОМПАС из закреплённого реестра (отчёты Archicad пропускает).
    → список сохранённых папок."""
    token, chat = _settings()
    saved = []
    _, blob = _pinned_index(token, chat)
    for line in blob.decode("utf-8", "replace").splitlines():
        try:
            e = json.loads(line)
        except ValueError:
            continue
        if not _is_kompas(e):
            continue
        d = os.path.join(out_dir, time.strftime("%Y%m%d_%H%M%S", time.localtime(e["date"])) + f"_m{e['message_id']}")
        if os.path.isdir(d):
            continue                            # уже забран
        os.makedirs(d, exist_ok=True)
        with open(os.path.join(d, "message.txt"), "w", encoding="utf-8") as f:
            f.write(e.get("text", ""))
        with open(os.path.join(d, e.get("file_name") or "log.txt"), "wb") as f:
            f.write(_download(token, e["file_id"]))
        saved.append(d)
    return saved


def latest_log():
    """Самый свежий журнал (для «Сообщить о проблеме» без ошибки)."""
    if not os.path.isdir(LOG_DIR):
        return None
    logs = [os.path.join(LOG_DIR, f) for f in os.listdir(LOG_DIR) if f.endswith(".log")]
    return max(logs, key=os.path.getmtime) if logs else None


def main(argv=None):
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--test", action="store_true", help="отправить тестовое сообщение")
    ap.add_argument("--send", action="store_true", help="отправить отчёт (вызывает менеджер)")
    ap.add_argument("--plugin", default="")
    ap.add_argument("--title", default="")
    ap.add_argument("--error-file")
    ap.add_argument("--comment-file")
    ap.add_argument("--inbox", metavar="ПАПКА", help="забрать отчёты КОМПАС в папку")
    ap.add_argument("--log-path", action="store_true", help="напечатать путь к журналу плагина")
    a = ap.parse_args(argv)

    if a.plugin:
        _state["plugin"] = a.plugin
    else:
        last = latest_log()
        if last:
            _state["plugin"] = os.path.splitext(os.path.basename(last))[0]
    _state["log"] = log_path()

    if a.log_path:
        print(_state["log"])
    elif a.inbox:
        got = inbox(a.inbox)
        print("\n".join(got) if got else "Новых отчётов нет.")
    elif a.test:
        print(send("Проверка отправки отчётов", comment="тест"))
    elif a.send:
        read = lambda p: open(p, encoding="utf-8").read() if p and os.path.isfile(p) else ""  # noqa: E731
        error = read(a.error_file)
        title = a.title or ("ОШИБКА" if error else "Журнал по запросу пользователя")
        print(send(title, error, read(a.comment_file)))
    else:
        ap.print_help()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Не отправилось: {e}")
        sys.exit(1)
