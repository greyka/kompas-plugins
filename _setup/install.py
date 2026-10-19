# -*- coding: utf-8 -*-
"""Установка плагинов КОМПАС-3D на компьютер (запускается из «Установить.bat»).

Python к этому моменту уже стоит в <корень>\\Python — его ставит «Установить.bat»
(из интернета: Python\\Онлайн, без интернета: Python\\Офлайн).

1. Каждая папка с plugin.ini копируется в <корень>\\<имя папки>.
2. Библиотека плагина (.rtw) подключается в Менеджер библиотек КОМПАС (меню «Приложения»).
   Старые подключения с тем же именем или файлом (из другой папки) удаляются.

Корень: C:\\KompasPlugins, если туда нельзя писать — %LOCALAPPDATA%\\KompasPlugins (передаёт --root).
"""
import argparse
import configparser
import ctypes
import os
import shutil
import subprocess
import traceback

PLUGINS_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))  # папка «Плагины»
TITLE = "Установка плагинов КОМПАС"
QUIET = False   # --quiet: без окон (для проверки/автоматизации)


def msg(text, flags=0x40):
    if QUIET:
        print(text)
        return 2
    return ctypes.windll.user32.MessageBoxW(None, text, TITLE, flags)


def kompas_running():
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq KOMPAS.exe", "/NH"],
                         capture_output=True, text=True, creationflags=0x08000000).stdout
    return "kompas.exe" in out.lower()


def same_path(a, b):
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def copy_tree(src, dst):
    if same_path(src, dst):
        return
    shutil.copytree(src, dst, dirs_exist_ok=True,
                    ignore=shutil.ignore_patterns("__pycache__", "*.pyc", "Чертежи"))


def find_plugins():
    plugins = []
    for name in sorted(os.listdir(PLUGINS_DIR)):
        ini = os.path.join(PLUGINS_DIR, name, "plugin.ini")
        if os.path.isfile(ini):
            cp = configparser.ConfigParser(inline_comment_prefixes=(";",))
            cp.read(ini, encoding="utf-8")
            sec = cp["plugin"]
            plugins.append(dict(dir=name, name=sec.get("name", name),
                                library=sec["library"], folder=sec.get("folder", "Приложения")))
    return plugins


def register_in_kompas(items):
    """items: [(имя, путь к .rtw, раздел)]. КОМПАС закрыт: запускаем скрытый, подключаем и закрываем —
    при закрытии он сохраняет список библиотек."""
    from win32com.client import CastTo, gencache
    api7 = gencache.EnsureDispatch("Kompas.Application.7")
    try:
        libs = CastTo(CastTo(api7.LibraryManager, "ILibraryManager").ProceduresLibraries,
                      "IProceduresLibraries")
        report = []
        for name, rtw, folder in items:
            for i in reversed(range(libs.Count)):          # старые подключения этого плагина
                lib = CastTo(libs.Item(i), "IProceduresLibrary")
                same_file = os.path.basename(lib.PathName).lower() == os.path.basename(rtw).lower()
                if (lib.Name == name or same_file) and not same_path(lib.PathName, rtw):
                    lib.Delete()
            found = None
            for i in range(libs.Count):
                lib = CastTo(libs.Item(i), "IProceduresLibrary")
                if same_path(lib.PathName, rtw):
                    found = lib
            if found is None:
                added = libs.Add(rtw, folder)
                if added is None:
                    report.append(f"✗ {name}: КОМПАС не принял библиотеку {rtw}")
                    continue
                found = CastTo(added, "IProceduresLibrary")
            found.Attach = True
            report.append(f"✓ {name} — меню «{folder}» → «{found.Name}»")
        return report
    finally:
        api7.Quit()


def main():
    global QUIET
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", required=True)
    ap.add_argument("--quiet", action="store_true")
    a = ap.parse_args()
    QUIET = a.quiet

    while kompas_running():
        if msg("Закройте КОМПАС-3D (все окна) и нажмите «Повтор».", 0x35) != 4:  # Повтор/Отмена
            return
    root = a.root
    ver_file = os.path.join(root, "Python", "VERSION.txt")
    py_ver = open(ver_file, encoding="utf-8").read().strip() if os.path.isfile(ver_file) else "есть"

    items = []
    for p in find_plugins():
        dst = os.path.join(root, p["dir"])
        copy_tree(os.path.join(PLUGINS_DIR, p["dir"]), dst)
        items.append((p["name"], os.path.join(dst, p["library"]), p["folder"]))
    report = register_in_kompas(items) if items else ["Плагинов не найдено"]
    report.append(install_updater(root))

    msg(f"Папка установки: {root}\nPython: {py_ver}\n\n" + "\n".join(report) +
        "\n\nЗапустите КОМПАС: меню «Приложения».")


ROOT_FILES = ("Обновить.bat", "ПРОЧТИ.txt", "VERSION", "CHANGELOG.md", "Менеджер плагинов КОМПАС.exe")


def install_updater(root):
    """Обновлятор и менеджер: _setup и файлы корня в <корень>, ярлыки — Автозагрузка (обновление),
    меню «Пуск» и рабочий стол (менеджер). Файл настроек settings.ini не трогаем."""
    copy_tree(os.path.join(PLUGINS_DIR, "_setup"), os.path.join(root, "_setup"))
    for f in ROOT_FILES:
        src = os.path.join(PLUGINS_DIR, f)
        if os.path.isfile(src) and not same_path(src, os.path.join(root, f)):
            shutil.copy2(src, os.path.join(root, f))
    # что установлено: версия из VERSION (= тег релиза vX.Y.Z), чтобы обновлятор не качал её повторно
    ver_src = os.path.join(PLUGINS_DIR, "VERSION")
    if os.path.isfile(ver_src):
        ver = open(ver_src, encoding="utf-8").read().strip()
        with open(os.path.join(root, "version.txt"), "w", encoding="utf-8") as f:
            f.write(f"v{ver}\n{ver}\nустановлено из папки\n")
    pending = os.path.join(root, "register_pending.txt")
    if os.path.exists(pending):
        os.remove(pending)
    report = []
    try:
        from win32com.client import Dispatch
        sh = Dispatch("WScript.Shell")

        def shortcut(folder, name, target, args, desc, icon=None):
            lnk = sh.CreateShortcut(os.path.join(folder, name + ".lnk"))
            lnk.TargetPath, lnk.Arguments, lnk.WorkingDirectory, lnk.Description = target, args, root, desc
            if icon:
                lnk.IconLocation = icon
            lnk.Save()

        programs = os.path.join(os.environ["APPDATA"], r"Microsoft\Windows\Start Menu\Programs")
        shortcut(os.path.join(programs, "Startup"), "Обновление плагинов КОМПАС",
                 os.path.join(root, "Python", "pythonw.exe"),
                 f'"{os.path.join(root, "_setup", "update.py")}" --root "{root}" --auto',
                 "Проверка обновлений плагинов КОМПАС на GitHub")
        report.append("✓ Автообновление с GitHub — при входе в Windows (вкл./выкл. — в менеджере)")
        exe = os.path.join(root, "Менеджер плагинов КОМПАС.exe")
        if os.path.isfile(exe):
            desktop = sh.SpecialFolders("Desktop")
            for folder in (programs, desktop):
                shortcut(folder, "Менеджер плагинов КОМПАС", exe, "", "Плагины КОМПАС: настройки, версия, обновление", exe)
            report.append("✓ Менеджер плагинов — ярлык на рабочем столе и в меню «Пуск»")
    except Exception as e:
        report.append(f"✗ Ярлыки не созданы: {e} (обновлять вручную: «Обновить.bat»)")
    return "\n".join(report)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        msg("Ошибка установки:\n\n" + traceback.format_exc(), 0x10)
        raise
