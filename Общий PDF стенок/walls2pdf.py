# -*- coding: utf-8 -*-
"""Плагин «Общий PDF стенок»: каждая стенка (.a3d) — отдельные чертежи .cdw,
а PDF — один общий файл на всю папку.

Использование:
    python walls2pdf.py <папка со сборками> [--sides left,right] [--out <папка>]

Результат в <папка>\\Чертежи\\:
    Вид слева\\<стенка>.cdw, Вид справа\\<стенка>.cdw   — чертежи каждой стенки отдельно
    <имя папки> — все стенки.pdf                       — один PDF: стенка за стенкой
                                                         (вид слева, вид справа), с закладками
Построение листа берётся из соседнего плагина «Чертежи рам» (frame2cdw.py).
"""
import argparse
import os
import shutil
import sys
import tempfile
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "lib"))                       # pypdf
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "Чертежи рам"))  # frame2cdw

import frame2cdw as fd  # noqa: E402
from pypdf import PdfReader, PdfWriter  # noqa: E402

CFG, kpsettings = fd.CFG, fd.kpsettings
TITLE = f"Общий PDF стенок — версия {kpsettings.version()}"


def msgbox(text, error=False):
    import ctypes
    ctypes.windll.user32.MessageBoxW(None, text, TITLE, 0x10 if error else 0x40)


def list_assemblies(src):
    if os.path.isfile(src):
        return [src]
    return [os.path.join(src, f) for f in sorted(os.listdir(src), key=str.lower)
            if f.lower().endswith(".a3d")]


def build_all(api7, files, out, sides):
    """Чертежи каждой стенки (.cdw) + временные PDF → [(стенка, вид, pdf)], ошибки, пропуски, tmp.
    Ошибки/пропуски — строки fd.problem(): имя файла, полный путь, причина."""
    tmp = tempfile.mkdtemp(prefix="walls2pdf-")
    pages, errors, skips = [], [], []
    for n, f in enumerate(files, 1):
        base = os.path.splitext(os.path.basename(f))[0]
        print(f"[{n}/{len(files)}] {base}")
        try:
            info, bodies = fd.read_frame(api7, f)
        except fd.NotAFrame as e:
            skips.append(fd.problem(f, e))
            print(f"--  {base}: {e}\n    {f}")
            continue
        except Exception as e:
            errors.append(fd.problem(f, e))
            print(f"ERR {base}: {f}")
            traceback.print_exc()
            continue
        for side in sides:
            title = fd.SIDES[side][0]
            sub = os.path.join(out, fd.side_folder(side))
            os.makedirs(sub, exist_ok=True)
            cdw = os.path.join(sub, base + ".cdw")
            try:
                doc2d, stat = fd.build_drawing(api7, f, info, bodies, cdw, side)   # уже проверен/исправлен
                doc2d.Close(0)
                if stat.get("remarks"):
                    rep = fd.write_remarks(cdw, stat["remarks"])
                    skips.append(fd.problem(cdw, "после автоисправления остались замечания — справка: " + rep, side))
                bak = cdw + ".bak"
                if os.path.exists(bak):
                    os.remove(bak)
                pdf = os.path.join(tmp, f"{n:03d}-{side}.pdf")
                fd.export_pdf(api7, cdw, pdf)
                if not os.path.isfile(pdf):
                    raise RuntimeError(f"КОМПАС не создал PDF из чертежа {cdw}")
                pages.append((base, title, pdf))
            except Exception as e:
                errors.append(fd.problem(f, e, side))
                print(f"ERR {base} [{title}]: {f}")
                traceback.print_exc()
    return pages, errors, skips, tmp


def merge_pdf(pages, target):
    """Один PDF: страницы по порядку, закладки «стенка → вид»."""
    w = PdfWriter()
    parents = {}
    for wall, view, pdf in pages:
        start = len(w.pages)
        for p in PdfReader(pdf).pages:
            w.add_page(p)
        if wall not in parents:
            parents[wall] = w.add_outline_item(wall, start)
        w.add_outline_item(view, start, parent=parents[wall])
    # метка версии — по присланному PDF видно, какой версией плагинов он сделан
    w.add_metadata({"/Creator": f"Плагины КОМПАС {fd.kpsettings.version()}", "/Producer": "pypdf"})
    with open(target, "wb") as fh:
        w.write(fh)
    return len(w.pages)


def run(src, out=None, sides=None, rot=None, gui=False, open_pdf=False):
    sides = sides or CFG.sides("walls")
    open_pdf = open_pdf and CFG.get("walls", "open_pdf")
    src = os.path.abspath(src)
    files = list_assemblies(src)
    if not files:
        raise RuntimeError(f"Нет сборок .a3d: {src}")
    folder = src if os.path.isdir(src) else os.path.dirname(src)
    out = os.path.abspath(out or os.path.join(folder, fd.out_folder_name()))
    os.makedirs(out, exist_ok=True)
    name = (CFG.get("walls", "pdf_name") or "{папка} — все стенки").replace("{папка}", os.path.basename(folder))
    target = os.path.join(out, name + ".pdf")
    side_list = [s.strip() for s in sides.split(",") if s.strip()]

    api7, started = fd.connect_kompas(rot)
    hide_prev = api7.HideMessage
    api7.HideMessage = 1
    tmp = None
    try:
        pages, errors, skips, tmp = build_all(api7, files, out, side_list)
    finally:
        api7.HideMessage = hide_prev
        if started:
            api7.Quit()
            del api7
            for _ in range(30):
                if not fd.kompas_running():
                    break
                time.sleep(1)
    try:
        n_pages = merge_pdf(pages, target) if pages else 0
    finally:
        if tmp:
            shutil.rmtree(tmp, ignore_errors=True)

    walls = len({p[0] for p in pages})
    text = (f"Стенок: {walls} из {len(files)}, листов в PDF: {n_pages}\n"
            f"PDF: {target}\nЧертежи .cdw: {out}")
    text += fd.problems_text("Ошибки — не вошли в PDF", errors)
    text += fd.problems_text("Пропущено / замечания", skips)
    print(text)
    if gui:
        if errors:
            fd.kpreport.error_dialog(TITLE, text)
        else:
            msgbox(text)
    if open_pdf and n_pages:
        os.startfile(target)
    return target


def menu(a):
    """1 — папка текущей сборки, 2 — выбрать папку, 3 — открыть папку «Чертежи»."""
    api7, _ = fd.connect_kompas(a.rot)
    doc = api7.ActiveDocument
    path = fd.CastTo(doc, "IKompasDocument").PathName if doc is not None else ""
    if a.menu == 9:
        kpsettings.open_manager("walls")
        return
    if a.menu == 1:
        if not path:
            msgbox("Откройте любую сохранённую сборку из папки со стенками и повторите команду.")
            return
        sides = a.sides or fd.choose_sides("walls", TITLE)
        if sides:
            run(os.path.dirname(path), sides=sides, rot=a.rot, gui=True, open_pdf=True)
    elif a.menu == 2:
        folder = fd.pick_folder(f"Папка со стенками (.a3d). Результат — в подпапке «{fd.out_folder_name()}».")
        sides = (a.sides or fd.choose_sides("walls", TITLE)) if folder else None
        if folder and sides:
            run(folder, sides=sides, rot=a.rot, gui=True, open_pdf=True)
    elif a.menu == 3:
        if not path:
            msgbox("Нет активного сохранённого документа.")
            return
        d = os.path.join(os.path.dirname(path), fd.out_folder_name())
        os.startfile(d if os.path.isdir(d) else os.path.dirname(path))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?")
    ap.add_argument("--out")
    ap.add_argument("--sides", help="left,right (по умолчанию — из настроек)")
    ap.add_argument("--menu", type=int)
    ap.add_argument("--rot")
    a = ap.parse_args()
    fd.kpreport.start_log("Общий PDF стенок")
    try:
        if a.menu:
            menu(a)
        elif a.src:
            run(a.src, a.out, a.sides)
        else:
            ap.error("укажите папку со сборками")
    except Exception:
        if a.menu:
            fd.kpreport.error_dialog(TITLE, traceback.format_exc())
        raise


if __name__ == "__main__":
    main()
