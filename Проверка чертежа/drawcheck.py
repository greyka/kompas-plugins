# -*- coding: utf-8 -*-
"""Плагин «Проверка чертежа»: смотрит готовый чертёж .cdw и выдаёт справку с рекомендациями.
НИЧЕГО НЕ ИСПРАВЛЯЕТ: чертёж открывается только для чтения и закрывается без сохранения.

Использование:
    python drawcheck.py <чертёж.cdw | папка>        — справка «<имя> — проверка.txt» рядом
Что проверяется:
    1. надписи налезают друг на друга (отметки, выноски позиций, размеры);
    2. подписи стоят вплотную (зазор меньше 1 мм);
    3. надписи выходят за рамку листа или заходят на основную надпись / таблицу;
    4. таблица: сечения вида «60,00x40,00x5,00», пустые длины у профилей;
    5. по модели (если вид связан со сборкой): от чего отсчитаны нулевые отметки —
       если край задаёт не основная рама (например, боковая пластина), об этом будет сказано.
"""
import argparse
import os
import re
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "Чертежи рам"))  # общие функции КОМПАС

import frame2cdw as fd  # noqa: E402
from win32com.client import CastTo, Dispatch, gencache  # noqa: E402

CFG, kpsettings = fd.CFG, fd.kpsettings
TITLE = f"Проверка чертежа — версия {kpsettings.version()}"
TOUCH_GAP = CFG.get("check", "touch_gap")            # мм листа: ближе — «вплотную»
OVERLAP_MIN = CFG.get("check", "overlap_min") / 100  # доля площади меньшей надписи — наложение
FORMATS_MM = {1: (841, 594), 2: (594, 420), 3: (420, 297), 4: (297, 210)}


def msgbox(text, error=False):
    import ctypes
    ctypes.windll.user32.MessageBoxW(None, text, TITLE, 0x10 if error else 0x40)


# --- чтение чертежа -------------------------------------------------------------
class Sheet:
    """Всё, что нужно для проверки, собранное из открытого чертежа (координаты — мм листа)."""

    def __init__(self):
        self.labels = []      # dict(kind, text, box)
        self.tables = []      # dict(box, rows)
        self.fmt = None       # (ширина, высота)
        self.sources = set()  # файлы моделей ассоциативных видов


def open_api5(api7, doc2d):
    K5 = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
    api5 = K5.KompasObject(Dispatch("Kompas.Application.5")._oleobj_)
    d5 = K5.ksDocument2D(api5.TransferInterface(doc2d._oleobj_, 1, 0))
    return K5, api5, d5


def read_sheet(api7, doc2d):
    K5, api5, d5 = open_api5(api7, doc2d)

    def gab(obj):
        r = K5.ksRectParam(api5.GetParamStruct(15))
        if not d5.ksGetObjGabaritRect(obj.Reference, r):
            return None
        a, b = K5.ksMathPointParam(r.GetpBot()), K5.ksMathPointParam(r.GetpTop())
        box = (min(a.x, b.x), min(a.y, b.y), max(a.x, b.x), max(a.y, b.y))
        return box if box[2] - box[0] > 0.01 or box[3] - box[1] > 0.01 else None

    sh = Sheet()
    try:
        sf = CastTo(CastTo(doc2d.LayoutSheets.Item(0), "ILayoutSheet").Format, "ISheetFormat")
        w, h = FORMATS_MM.get(sf.Format, (sf.FormatWidth, sf.FormatHeight))
        sh.fmt = (h, w) if sf.VerticalOrientation and w > h else (w, h)
    except Exception:
        pass

    views = CastTo(doc2d.ViewsAndLayersManager, "IViewsAndLayersManager").Views
    for i in range(views.Count):
        v = CastTo(views.View(i), "IView")
        k, vx, vy = v.Scale, v.X, v.Y
        try:
            av = CastTo(v, "IAssociationView")
            if av.SourceFileName:
                sh.sources.add(av.SourceFileName)
        except Exception:
            pass
        dc = CastTo(v, "IDrawingContainer")
        for j in range(dc.DrawingTexts.Count):
            t = CastTo(dc.DrawingTexts.DrawingText(j), "IDrawingText")
            s = CastTo(t, "IText").Str
            box = gab(t)
            if box and s.strip() and not s.startswith("^("):   # ^( — служебная ссылка (обозначение вида)
                sh.labels.append(dict(kind="надпись", text=s.strip(), box=box))
        sc = CastTo(v, "ISymbols2DContainer")
        for j in range(sc.Leaders.Count):
            ld = sc.Leaders.Leader(j)
            try:
                s = CastTo(CastTo(ld, "IPositionLeader").Positions, "IText").Str
                kind = "позиция"
            except Exception:
                s, kind = "", "выноска"
            g = gab(CastTo(ld, "IBaseLeader"))
            if not g:
                continue
            br = CastTo(ld, "IBranchs")
            x0, y0 = vx + br.X0 * k, vy + br.Y0 * k
            # текст над полкой: полка идёт от X0 в сторону дальнего края габарита
            if g[2] - x0 > x0 - g[0]:
                box = (x0, y0, min(g[2], x0 + 4 + 3 * len(s)), y0 + 6)
            else:
                box = (max(g[0], x0 - 4 - 3 * len(s)), y0, x0, y0 + 6)
            sh.labels.append(dict(kind=kind, text=s.strip() or "?", box=box))
        for j in range(sc.LineDimensions.Count):
            d = CastTo(sc.LineDimensions.LineDimension(j), "ILineDimension")
            try:
                val = CastTo(d, "IDimensionText").NominalValue
            except Exception:
                continue
            text = fd.fmt_len(val)
            g = gab(d)
            if not g:
                continue
            x1, y1 = vx + d.X1 * k, vy + d.Y1 * k
            x2, y2 = vx + d.X2 * k, vy + d.Y2 * k
            x3, y3 = vx + d.X3 * k, vy + d.Y3 * k
            tw, th = 2.6 * len(text) + 1, 5.0
            if d.Orientation == 1:      # горизонтальный: текст над размерной линией по центру
                cx = (x1 + x2) / 2
                box = (cx - tw / 2, y3 + 0.5, cx + tw / 2, y3 + 0.5 + th)
            elif d.Orientation == 2:    # вертикальный: текст слева от линии, повёрнут
                cy = (y1 + y2) / 2
                box = (x3 - 0.5 - th, cy - tw / 2, x3 - 0.5, cy + tw / 2)
            else:                       # наклонный (диагональ) — пропускаем, положение сложно оценить
                continue
            sh.labels.append(dict(kind="размер", text=text, box=box))
        for j in range(sc.DrawingTables.Count):
            tb = sc.DrawingTables.DrawingTable(j)
            g = gab(CastTo(tb, "IDrawingTable"))
            t = CastTo(tb, "ITable")
            rows = [[CastTo(CastTo(t.Cell(r, c), "ITableCell").Text, "IText").Str
                     for c in range(t.ColumnsCount)] for r in range(t.RowsCount)]
            sh.tables.append(dict(box=g, rows=rows))
    return sh


# --- правила --------------------------------------------------------------------
def inter(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return w * h if w > 0 and h > 0 else 0.0


def gap(a, b):
    dx = max(a[0] - b[2], b[0] - a[2], 0)
    dy = max(a[1] - b[3], b[1] - a[3], 0)
    return (dx * dx + dy * dy) ** 0.5


def area(a):
    return max(a[2] - a[0], 0.01) * max(a[3] - a[1], 0.01)


def where(box, sh):
    """Человеческое описание места на листе."""
    if not sh.fmt:
        return f"x={box[0]:.0f}, y={box[1]:.0f} мм"
    w, h = sh.fmt
    cx, cy = (box[0] + box[2]) / 2, (box[1] + box[3]) / 2
    gx = "слева" if cx < w / 3 else ("справа" if cx > 2 * w / 3 else "по центру")
    gy = "внизу" if cy < h / 3 else ("вверху" if cy > 2 * h / 3 else "посередине")
    return f"{gy} {gx} листа (x≈{cx:.0f}, y≈{cy:.0f} мм)"


def q(lb):
    return f"{lb['kind']} «{lb['text']}»"


def check_overlaps(sh):
    out, near = [], []
    L = sh.labels
    for i in range(len(L)):
        for j in range(i + 1, len(L)):
            a, b = L[i], L[j]
            if a["kind"] == "размер" and b["kind"] == "размер":
                continue  # габариты размеров оцениваются грубо — не сравниваем между собой
            s = inter(a["box"], b["box"])
            if s > OVERLAP_MIN * min(area(a["box"]), area(b["box"])):
                out.append(f"• {q(a)} налезает на {q(b)} — {where(a['box'], sh)}.")
            elif a["kind"] == b["kind"] == "надпись" and gap(a["box"], b["box"]) < TOUCH_GAP and s == 0:
                near.append(f"• {q(a)} и {q(b)} стоят вплотную — {where(a['box'], sh)}.")
    rec = []
    if out:
        rec.append(("Надписи налезают друг на друга", out,
                    "Сдвиньте полку выноски или подпись так, чтобы текст читался; "
                    "выноски лучше уводить внутрь рамы, подальше от столбика отметок."))
    if near:
        rec.append(("Подписи стоят вплотную", near,
                    "Близкие отметки лучше развести изломом выносной линии или дать одну общую."))
    return rec


def check_bounds(sh):
    if not sh.fmt:
        return []
    w, h = sh.fmt
    inner = (20, 5, w - 5, h - 5)                 # внутренняя рамка (ГОСТ 2.104)
    stamp = (w - 5 - 185, 5, w - 5, 5 + 55)       # основная надпись (форма 1)
    out = []
    for lb in sh.labels:
        b = lb["box"]
        if b[0] < inner[0] - 0.5 or b[1] < inner[1] - 0.5 or b[2] > inner[2] + 0.5 or b[3] > inner[3] + 0.5:
            out.append(f"• {q(lb)} выходит за рамку листа — {where(b, sh)}.")
        elif inter(b, stamp) > 0:
            out.append(f"• {q(lb)} заходит на основную надпись — {where(b, sh)}.")
        else:
            for t in sh.tables:
                if t["box"] and inter(b, t["box"]) > 0:
                    out.append(f"• {q(lb)} заходит на таблицу — {where(b, sh)}.")
                    break
    return [("Надписи за рамкой / на штампе / на таблице", out,
             "Перенесите надпись внутрь поля чертежа.")] if out else []


SECTION_ZEROS = re.compile(r"\d+[.,]0+\s*[xх×]", re.I)


def check_tables(sh):
    zeros, empty = [], []
    for t in sh.tables:
        for r, row in enumerate(t["rows"][1:], start=2):
            cells = [c.strip() for c in row]
            sec = cells[1] if len(cells) > 1 else ""
            if SECTION_ZEROS.search(sec):
                zeros.append(f"• строка {r}, поз. {cells[0] or '?'}: «{sec}»")
            if sec[:1] in "□—" and len(cells) > 3 and not cells[3]:
                empty.append(f"• строка {r}, поз. {cells[0] or '?'}: «{sec}» — нет длины")
    rec = []
    if zeros:
        rec.append(("Сечения с лишними нулями", zeros,
                    "Запишите сечение без «,00»: например «□60x40x5» вместо «□60,00x40,00x5,00»."))
    if empty:
        rec.append(("Пустая длина профиля в таблице", empty,
                    "Проверьте в модели свойство «Длина профиля» у этих тел."))
    return rec


def main_section(bodies, u, v):
    """Сечение основной рамы: элементы этого сечения (не меньше 3 шт.) охватывают наибольшую площадь.
    Пары боковых пластин (2 шт.) и мелкие вставки так не пройдут."""
    by = {}
    for b in bodies:
        by.setdefault(b["sec"], []).append(b)

    def span(group):
        return ((max(b["box"][u + 3] for b in group) - min(b["box"][u] for b in group)) *
                (max(b["box"][v + 3] for b in group) - min(b["box"][v] for b in group)))
    cands = [s for s, g in by.items() if len(g) >= 3] or list(by)
    return max(cands, key=lambda s: (round(span(by[s]), -2), len(by[s])))


def check_base(api7, sources):
    """По модели: задаёт ли нулевые отметки основная рама или «чужой» элемент (пластина и т.п.)."""
    rec = []
    for src in sources:
        if not os.path.isfile(src):
            rec.append(("Модель вида не найдена", [f"• {src}"],
                        "Проверка отсчёта отметок по модели пропущена."))
            continue
        try:
            _, bodies = fd.read_frame(api7, src)
        except Exception as e:
            rec.append(("Не удалось прочитать модель вида", [f"• {os.path.basename(src)}: {e}"], ""))
            continue
        u, v, thin, lo, hi = fd.frame_axes(bodies)
        main = main_section(bodies, u, v)
        frame = [b for b in bodies if b["sec"] == main]
        flo = [min(b["box"][i] for b in frame) for i in range(3)]
        fhi = [max(b["box"][i + 3] for b in frame) for i in range(3)]
        lines = []
        for axis, name in ((u, "по горизонтали"), (v, "по вертикали")):
            for edge, (all_e, fr_e) in (("начало", (lo[axis], flo[axis])), ("конец", (hi[axis], fhi[axis]))):
                d = abs(all_e - fr_e)
                if d > 1:
                    culprit = [b for b in bodies if b["sec"] != main and
                               abs((b["box"][axis] if edge == "начало" else b["box"][axis + 3]) - all_e) < 0.5]
                    who = ", ".join(sorted({f"{b['sec']} (поз. {b['pos']})" for b in culprit})) or "другой элемент"
                    shift = "начинается" if edge == "начало" else "заканчивается"
                    lines.append(f"• {name}, {edge}: крайний элемент — {who}; рама «{main}» "
                                 f"{shift} на {d:.0f} мм {'дальше' if edge == 'начало' else 'раньше'}.")
        if lines:
            rec.append((f"Отсчёт отметок ({os.path.basename(src)})", lines,
                        "Нулевые отметки и габариты сейчас считаются от крайнего элемента, а не от рамы. "
                        "Решите, от чего давать размеры, — при необходимости укажите базу явно."))
    return rec


# --- справка ---------------------------------------------------------------------
def check_file(api7, cdw):
    opened = fd.find_open_doc(api7, cdw)
    doc = opened or api7.Documents.Open(cdw, False, True)   # только чтение
    try:
        sh = read_sheet(api7, CastTo(doc, "IKompasDocument2D"))
    finally:
        if opened is None:
            doc.Close(0)                                     # без сохранения
    rec = check_overlaps(sh) + check_bounds(sh) + check_tables(sh) + check_base(api7, sh.sources)
    return sh, rec


def format_report(cdw, sh, rec):
    lines = [f"ПРОВЕРКА ЧЕРТЕЖА: {os.path.basename(cdw)}",
             f"Папка: {os.path.dirname(cdw)}",
             f"Дата: {time.strftime('%d.%m.%Y %H:%M')}",
             f"Лист: {sh.fmt[0]}×{sh.fmt[1]} мм" if sh.fmt else "Лист: формат не определён",
             f"Надписей проверено: {len(sh.labels)}, таблиц: {len(sh.tables)}",
             "Чертёж НЕ изменялся — ниже только рекомендации.", ""]
    if not rec:
        lines.append("Замечаний нет.")
    for n, (title, items, advice) in enumerate(rec, 1):
        lines.append(f"{n}. {title} — {len(items)}")
        lines += ["   " + s for s in items[:30]]
        if len(items) > 30:
            lines.append(f"   … и ещё {len(items) - 30}")
        if advice:
            lines.append(f"   Рекомендация: {advice}")
        lines.append("")
    return "\n".join(lines)


def list_drawings(src):
    if os.path.isfile(src):
        return [src]
    out = []
    for root, _, files in os.walk(src):
        out += [os.path.join(root, f) for f in sorted(files) if f.lower().endswith(".cdw")]
    return out


def run(src, rot=None, gui=False):
    src = os.path.abspath(src)
    files = list_drawings(src)
    if not files:
        raise RuntimeError(f"Нет чертежей .cdw: {src}")
    api7, started = fd.connect_kompas(rot)
    hide_prev = api7.HideMessage
    api7.HideMessage = 1
    reports, summary = [], []
    try:
        for f in files:
            try:
                sh, rec = check_file(api7, f)
                text = format_report(f, sh, rec)
                cnt = sum(len(i) for _, i, _ in rec)
            except Exception as e:
                text, cnt = f"ПРОВЕРКА ЧЕРТЕЖА: {os.path.basename(f)}\nОшибка: {e}\n", -1
                traceback.print_exc()
            reports.append(text)
            summary.append(f"{'ошибка' if cnt < 0 else (str(cnt) + ' замеч.' if cnt else 'ок')} — "
                           f"{os.path.relpath(f, src if os.path.isdir(src) else os.path.dirname(src))}")
            print(summary[-1])
    finally:
        api7.HideMessage = hide_prev
        if started:
            api7.Quit()
            del api7
            for _ in range(30):
                if not fd.kompas_running():
                    break
                time.sleep(1)
    if os.path.isfile(src):
        target = os.path.splitext(src)[0] + " — проверка.txt"
        body = reports[0]
    else:
        target = os.path.join(src, "Проверка чертежей.txt")
        body = "СВОДКА\n" + "\n".join(summary) + "\n\n" + ("\n" + "=" * 70 + "\n\n").join(reports)
    with open(target, "w", encoding="utf-8-sig") as fh:
        fh.write(body)
    print(f"Справка: {target}")
    if gui and CFG.get("check", "open_report"):
        os.startfile(target)
    elif gui:
        msgbox(f"Проверено чертежей: {len(files)}\nСправка: {target}")
    return target


def menu(a):
    """1 — текущий чертёж, 2 — все чертежи в папке, 3 — открыть последнюю справку."""
    api7, _ = fd.connect_kompas(a.rot)
    doc = api7.ActiveDocument
    path = CastTo(doc, "IKompasDocument").PathName if doc is not None else ""
    if a.menu == 9:
        kpsettings.open_manager("check")
        return
    if a.menu == 1:
        if not path.lower().endswith(".cdw"):
            msgbox("Откройте сохранённый чертёж (.cdw) и повторите команду.")
            return
        run(path, rot=a.rot, gui=True)
    elif a.menu == 2:
        folder = fd.pick_folder("Папка с чертежами (.cdw, включая подпапки). Справка — в этой папке.")
        if folder:
            run(folder, rot=a.rot, gui=True)
    elif a.menu == 3:
        cand = []
        if path:
            d = os.path.dirname(path)
            cand = [os.path.join(d, f) for f in os.listdir(d) if f.endswith("проверка.txt") or
                    f == "Проверка чертежей.txt"]
        if not cand:
            msgbox("Справок рядом с текущим документом нет.")
            return
        os.startfile(max(cand, key=os.path.getmtime))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?")
    ap.add_argument("--menu", type=int)
    ap.add_argument("--rot")
    a = ap.parse_args()
    try:
        if a.menu:
            menu(a)
        elif a.src:
            run(a.src)
        else:
            ap.error("укажите чертёж или папку")
    except Exception:
        if a.menu:
            msgbox(traceback.format_exc(), error=True)
        raise


if __name__ == "__main__":
    main()
