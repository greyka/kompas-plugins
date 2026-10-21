# -*- coding: utf-8 -*-
"""Плагин «PDF из чертежей»: выбранные чертежи КОМПАС (.cdw) → один общий PDF.
Чертежи не изменяются: каждый конвертируется из файла на диске во временный PDF, затем всё склеивается.

Использование:
    python cdw2pdf.py <чертёж.cdw | папка> [...] -o <итог.pdf>
Меню КОМПАС (--menu):
    1 — выбрать чертежи в окне (несколько файлов) → один PDF
    2 — все чертежи, открытые сейчас в КОМПАС → один PDF
    3 — все чертежи из папки (с подпапками) → один PDF
В PDF — закладка на каждый чертёж; многолистовые чертежи входят всеми листами.
"""
import argparse
import os
import re
import shutil
import sys
import tempfile
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "lib"))                           # pypdf
sys.path.insert(0, HERE)                    # cdwmerge (переносной Python сам папку скрипта не добавляет)
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "Чертежи рам"))  # общие функции КОМПАС

import frame2cdw as fd  # noqa: E402
from pypdf import PdfReader, PdfWriter  # noqa: E402
from win32com.client import CastTo  # noqa: E402

CFG, kpsettings = fd.CFG, fd.kpsettings
TITLE = f"PDF из чертежей — версия {kpsettings.version()}"
ksDocumentDrawing = 1


def msgbox(text, error=False, yesno=False):
    import ctypes
    flags = (0x10 if error else 0x40) | (0x04 if yesno else 0)
    return ctypes.windll.user32.MessageBoxW(None, text, TITLE, flags)


def natural_key(path):
    """Ст2 раньше Ст10."""
    name = os.path.basename(path).lower()
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", name)]


# --- диалоги Windows ---------------------------------------------------------------
def ask_drawings(initial_dir=None):
    import win32con
    import win32gui
    try:
        names, _, _ = win32gui.GetOpenFileNameW(
            InitialDir=initial_dir or "",
            Filter="Чертежи КОМПАС (*.cdw)\0*.cdw\0Все файлы (*.*)\0*.*\0",
            Title="Выберите чертежи (несколько — с Ctrl или Shift)",
            Flags=win32con.OFN_ALLOWMULTISELECT | win32con.OFN_EXPLORER | win32con.OFN_FILEMUSTEXIST,
            MaxFile=65535)
    except Exception:          # нажали «Отмена»
        return []
    parts = names.split("\0")
    if len(parts) == 1:
        return parts
    return [os.path.join(parts[0], p) for p in parts[1:] if p]


def ask_target(initial_dir, default_name, ext="pdf"):
    import win32con
    import win32gui
    filt, title = (("Чертёж КОМПАС (*.cdw)\0*.cdw\0", "Сохранить общий чертёж как") if ext == "cdw"
                   else ("PDF (*.pdf)\0*.pdf\0", "Сохранить общий PDF как"))
    try:
        name, _, _ = win32gui.GetSaveFileNameW(
            InitialDir=initial_dir or "", File=default_name, DefExt=ext,
            Filter=filt, Title=title,
            Flags=win32con.OFN_OVERWRITEPROMPT | win32con.OFN_EXPLORER)
        return name
    except Exception:
        return None


# --- сборка PDF ----------------------------------------------------------------------
def collect(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for root, _, names in os.walk(p):
                files += [os.path.join(root, n) for n in names if n.lower().endswith(".cdw")]
        elif p.lower().endswith(".cdw"):
            files.append(p)
    seen, out = set(), []
    for f in files:
        k = os.path.normcase(os.path.abspath(f))
        if k not in seen:
            seen.add(k)
            out.append(os.path.abspath(f))
    return out


def make_pdf(api7, files, target):
    """Конвертирует каждый чертёж (из файла, без открытия на редактирование) и склеивает в target."""
    tmp = tempfile.mkdtemp(prefix="cdw2pdf-")
    w, errors, done = PdfWriter(), [], 0
    try:
        for n, f in enumerate(files, 1):
            name = os.path.splitext(os.path.basename(f))[0]
            print(f"[{n}/{len(files)}] {name}")
            pdf = os.path.join(tmp, f"{n:04d}.pdf")
            try:
                fd.export_pdf(api7, f, pdf)
                if not os.path.isfile(pdf):
                    raise RuntimeError("КОМПАС не создал PDF")
                start = len(w.pages)
                for page in PdfReader(pdf).pages:
                    w.add_page(page)
                if CFG.get("pdf", "bookmarks"):
                    w.add_outline_item(name, start)
                done += 1
            except Exception as e:
                errors.append(f"{name}: {e}")
                traceback.print_exc()
        if done:
            os.makedirs(os.path.dirname(os.path.abspath(target)), exist_ok=True)
            # метка версии — по присланному PDF видно, какой версией плагинов он сделан
            w.add_metadata({"/Creator": f"Плагины КОМПАС {fd.kpsettings.version()}", "/Producer": "pypdf"})
            with open(target, "wb") as fh:
                w.write(fh)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return done, len(w.pages), errors


def run(files, target, rot=None, gui=False, open_pdf=False):
    if not files:
        raise RuntimeError("Не выбрано ни одного чертежа .cdw")
    api7, started = fd.connect_kompas(rot)
    try:
        done, pages, errors = make_pdf(api7, files, target)
    finally:
        if started:
            api7.Quit()
            del api7
            for _ in range(30):
                if not fd.kompas_running():
                    break
                time.sleep(1)
    text = f"Чертежей: {done} из {len(files)}, листов в PDF: {pages}\nPDF: {target}"
    if errors:
        text += "\n\nНе вошли:\n" + "\n".join(errors[:15])
    print(text)
    if gui and (errors or not done):
        fd.kpreport.error_dialog(TITLE, text)
    if open_pdf and done and CFG.get("pdf", "open_pdf"):
        os.startfile(target)
    elif gui and done and not errors:
        msgbox(text)
    return target


def sort_files(files):
    """«Ст2» раньше «Ст10» (настройка «Сортировка»), иначе — по алфавиту."""
    return sorted(files, key=natural_key if CFG.get("pdf", "natural_sort") else str.lower)


def open_drawings(api7):
    """Открытые в КОМПАС чертежи (в порядке вкладок); несохранённые изменения — спросить."""
    docs = api7.Documents
    out, unsaved, noname = [], [], []
    for i in range(docs.Count):
        d = CastTo(docs.Item(i), "IKompasDocument")
        if d.DocumentType != ksDocumentDrawing:
            continue
        if not d.PathName:
            noname.append(d.Name)
            continue
        if d.Changed:
            unsaved.append(d)
        out.append(d.PathName)
    if unsaved:
        names = "\n".join(d.Name for d in unsaved)
        if msgbox("Эти чертежи изменены и не сохранены — в PDF попадёт версия с диска:\n\n"
                  f"{names}\n\nСохранить их сейчас?", yesno=True) == 6:
            for d in unsaved:
                d.Save()
    if noname:
        msgbox("Пропущены несохранённые новые чертежи (нет файла):\n" + "\n".join(noname))
    return out


def menu(a):
    api7, _ = fd.connect_kompas(a.rot)
    doc = api7.ActiveDocument
    cur = CastTo(doc, "IKompasDocument").PathName if doc is not None else ""
    start_dir = os.path.dirname(cur) if cur else ""
    if a.menu == 9:
        kpsettings.open_manager("pdf")
        return
    # 1–3 → один PDF, 4–6 → один общий чертёж (.cdw); источник: выбрать / открытые / папка
    if a.menu not in (1, 2, 3, 4, 5, 6):
        return
    to_cdw, how = a.menu >= 4, (a.menu - 1) % 3
    what = "один общий чертёж" if to_cdw else "один PDF"
    if how == 0:
        files = sort_files(collect(ask_drawings(start_dir)))
    elif how == 1:
        files = open_drawings(api7)
        if not files:
            msgbox("В КОМПАС нет открытых сохранённых чертежей.")
            return
    else:
        folder = fd.pick_folder(f"Папка с чертежами (.cdw, включая подпапки) → {what}")
        files = sort_files(collect([folder])) if folder else []
        if folder and not files:
            msgbox(f"В папке нет чертежей .cdw:\n{folder}")
            return
    if not files:
        return
    first_dir = os.path.dirname(files[0])
    base = os.path.basename(first_dir) or "Чертежи"
    if to_cdw:
        target = ask_target(first_dir, f"{base} — общий чертёж.cdw", "cdw")
        if target:
            files = [f for f in files if os.path.normcase(f) != os.path.normcase(target)]
            run_merge(files, target, api7)
        return
    target = ask_target(first_dir, base + ".pdf")
    if target:
        run(files, target, rot=a.rot, gui=True, open_pdf=True)


def run_merge(files, target, api7=None, gui=True):
    """Общий многолистовой чертёж из files; PDF рядом (настройка), открыть в КОМПАС (настройка)."""
    import cdwmerge
    started = False
    if api7 is None:
        api7, started = fd.connect_kompas()
    hide = api7.HideMessage
    api7.HideMessage = 1
    try:
        sheets, errors = cdwmerge.merge(api7, files, target)
        pdf = None
        if CFG.get("pdf", "merged_pdf"):
            pdf = os.path.splitext(target)[0] + ".pdf"
            if not fd.export_pdf(api7, target, pdf) or not os.path.isfile(pdf):
                errors.append((target, "не удалось сделать PDF из общего чертежа"))
                pdf = None
    finally:
        api7.HideMessage = hide
    text = (f"Общий чертёж: листов {sheets} из {len(files)} чертежей\n{target}"
            + (f"\nPDF: {pdf}" if pdf else ""))
    text += fd.problems_text("Не вошли", [fd.problem(f, r) for f, r in errors])
    print(text)
    if gui:
        if errors:
            fd.kpreport.error_dialog(TITLE, text)
        elif CFG.get("pdf", "open_merged"):
            api7.Documents.Open(target, True, False)
        else:
            msgbox(text)
    if started:
        api7.Quit()
    return target


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="*", help="чертежи .cdw и/или папки")
    ap.add_argument("-o", "--out", help="итоговый PDF (или .cdw с --merge)")
    ap.add_argument("--merge", action="store_true", help="собрать общий чертёж .cdw вместо PDF")
    ap.add_argument("--menu", type=int)
    ap.add_argument("--rot")
    a = ap.parse_args()
    fd.kpreport.start_log("PDF из чертежей")
    try:
        if a.menu:
            menu(a)
        elif a.src:
            files = collect(a.src)
            if len(a.src) == 1 and os.path.isdir(a.src[0]):
                files = sort_files(files)
            if a.merge:
                target = a.out or os.path.join(os.path.dirname(files[0]) if files else ".", "Общий чертёж.cdw")
                run_merge([f for f in files if os.path.normcase(f) != os.path.normcase(target)], target, gui=False)
            else:
                target = a.out or os.path.join(os.path.dirname(files[0]) if files else ".", "Чертежи.pdf")
                run(files, target)
        else:
            ap.error("укажите чертежи или папку")
    except Exception:
        if a.menu:
            fd.kpreport.error_dialog(TITLE, traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
