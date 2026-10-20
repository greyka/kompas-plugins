# -*- coding: utf-8 -*-
"""Исправитель чертежа: берёт замечания «Проверки чертежа» (drawcheck) и правит то, что можно
поправить без человека. Отдельный скрипт — проверка сама ничего не меняет.

Что исправляет:
  • позиционные выноски, налезающие на надписи / размеры / другие выноски / таблицу / штамп или
    выходящие за рамку, — переносит полку на свободное место вокруг той же точки на элементе;
  • надписи за рамкой листа — сдвигает внутрь;
  • сечения с лишними нулями в таблице («□60,00x40,00x5,00» → «□60x40x5»).
Что осталось — возвращает списком (как справка проверки), это надо смотреть руками.

Вызывается:
  • «Чертежи рам» / «Общий PDF стенок» — сразу после построения листа, до сохранения и PDF
    (настройка «Проверка чертежа» → «Исправлять чертёж сразу при построении»);
  • меню «Проверка чертежа» → «Исправить текущий чертёж» / «Исправить чертежи в папке…»;
  • python drawfix.py <чертёж.cdw | папка>
"""
import math
import os
import sys
import time
import traceback

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import drawcheck as dc  # noqa: E402

fd, CastTo = dc.fd, dc.CastTo

RADII = (12, 18, 25, 33, 42)                 # мм листа: куда пробуем переносить полку
DIRS = [(1, 1), (-1, 1), (1, -1), (-1, -1), (1, 0.35), (-1, 0.35), (0.35, 1), (-0.35, 1),
        (1, -0.35), (-1, -0.35), (0.35, -1), (-0.35, -1)]
CLEAR = 2.5                                  # мм — минимальный зазор вокруг текста выноски


def _grow(b, d):
    return (b[0] - d, b[1] - d, b[2] + d, b[3] + d)


W_TEXT, W_LINE, W_MEMBER = 10.0, 3.0, 1.5   # «цена» мм² перекрытия: текст, линия, элемент рамы


def _blocked(box, others, sh, zones, lines=(), members=()):
    """Сколько «мешает» положение box: перекрытия (с зазором) с надписями, штампом/таблицами,
    линиями (размеры, засечки, диагональ) и элементами рамы; выход за рамку листа."""
    w, h = sh.fmt or (420, 297)
    pen = 0.0
    if box[0] < 21 or box[1] < 6 or box[2] > w - 6 or box[3] > h - 6:
        pen += 1000
    g = _grow(box, CLEAR)
    for o in others:
        pen += dc.inter(g, o) * W_TEXT
    for z in zones:
        pen += dc.inter(g, z) * W_TEXT
    for ln in lines:
        pen += dc.inter(box, ln) * W_LINE
    for m in members:
        pen += dc.inter(box, m) * W_MEMBER
    return pen


def _seg_boxes(x1, y1, x2, y2, w=0.6, step=2.0):
    """Отрезок → цепочка маленьких квадратов (для оценки перекрытия)."""
    n = max(1, int(math.hypot(x2 - x1, y2 - y1) / step))
    return [(x1 + (x2 - x1) * i / n - w, y1 + (y2 - y1) * i / n - w,
             x1 + (x2 - x1) * i / n + w, y1 + (y2 - y1) * i / n + w) for i in range(n + 1)]


def read_geometry(api7, doc2d, bodies=None):
    """Линии и элементы рамы на листе (мм листа): размерные линии, засечки отметок, диагональ —
    из чертежа; элементы рамы — из модели ассоциативного вида (bodies можно передать готовые)."""
    lines, members = [], []
    views = CastTo(doc2d.ViewsAndLayersManager, "IViewsAndLayersManager").Views
    for i in range(views.Count):
        v = CastTo(views.View(i), "IView")
        k, vx, vy = v.Scale, v.X, v.Y
        S = lambda x, y: (vx + x * k, vy + y * k)  # noqa: E731
        sc = CastTo(v, "ISymbols2DContainer")
        for j in range(sc.LineDimensions.Count):
            d = CastTo(sc.LineDimensions.LineDimension(j), "ILineDimension")
            p1, p2, p3 = S(d.X1, d.Y1), S(d.X2, d.Y2), S(d.X3, d.Y3)
            if d.Orientation == 1:
                lines += _seg_boxes(p1[0], p3[1], p2[0], p3[1])
            elif d.Orientation == 2:
                lines += _seg_boxes(p3[0], p1[1], p3[0], p2[1])
            else:
                lines += _seg_boxes(p1[0], p1[1], p2[0], p2[1])
        dcont = CastTo(v, "IDrawingContainer")
        for j in range(dcont.LineSegments.Count):
            s = CastTo(dcont.LineSegments.LineSegment(j), "ILineSegment")
            a, b = S(s.X1, s.Y1), S(s.X2, s.Y2)
            lines += _seg_boxes(a[0], a[1], b[0], b[1])
        try:
            av = CastTo(v, "IAssociationView")
            src = av.SourceFileName
        except Exception:
            src = ""
        if not src:
            continue
        try:
            bs = bodies if bodies is not None else fd.read_frame(api7, src)[1]
            u, vv, thin, lo, hi = fd.frame_axes(bs)
            for b in bs:
                pts = [av.GetProjectionPoint(*fd._abs(lo, u, vv, thin, x, y))[-2:]
                       for x in (b["box"][u], b["box"][u + 3]) for y in (b["box"][vv], b["box"][vv + 3])]
                xs = [vx + p[0] * k for p in pts]
                ys = [vy + p[1] * k for p in pts]
                members.append((min(xs), min(ys), max(xs), max(ys)))
        except Exception as e:
            print(f"  элементы рамы для исправителя не прочитаны: {e}")
    return lines, members


def _zones(sh):
    w, h = sh.fmt or (420, 297)
    z = [(w - 5 - 185, 5, w - 5, 5 + 55)]                       # основная надпись
    z += [t["box"] for t in sh.tables if t["box"]]
    return z


def leader_lines(lb):
    """Линии выноски на листе: от точки на элементе до полки + сама полка (цепочка квадратиков)."""
    ld = lb["obj"]
    k, vx, vy = lb["view"]
    br = CastTo(ld, "IBranchs")
    if br.BranchCount < 1:
        return []
    ax, ay = vx + br.BranchX(0) * k, vy + br.BranchY(0) * k
    x0, y0 = vx + br.X0 * k, vy + br.Y0 * k
    b = lb["box"]
    x_end = b[2] if abs(b[0] - x0) < abs(b[2] - x0) else b[0]
    return _seg_boxes(ax, ay, x0, y0) + _seg_boxes(x0, y0, x_end, y0)


def _own_lines_pen(lb, text_boxes):
    """Линия перемещаемой выноски не должна резать чужие надписи."""
    pen = 0.0
    for s in leader_lines(lb)[:-1]:
        for t in text_boxes:
            pen += dc.inter(s, t) * W_TEXT
    return pen


def _move_leader(lb, sh, others, zones, lines, members):
    """Перенести полку выноски туда, где меньше всего помех. → (было, стало) — «цена» помех."""
    ld = lb["obj"]
    k, vx, vy = lb["view"]
    br = CastTo(ld, "IBranchs")
    base = CastTo(ld, "IBaseLeader")
    if br.BranchCount < 1:
        return None
    ax, ay = br.BranchX(0), br.BranchY(0)                       # точка на элементе (вид)
    start = _blocked(lb["box"], others, sh, zones, lines, members) + _own_lines_pen(lb, others)
    best = (start, (br.X0, br.Y0), lb["box"])
    for r in RADII:
        for dx, dy in DIRS:
            n = math.hypot(dx, dy)
            br.X0, br.Y0 = ax + r * dx / n / k, ay + r * dy / n / k
            base.Update()
            box = dc.leader_box(sh.gab, ld, lb["text"], lb["view"])
            if not box:
                continue
            lb["box"] = box
            pen = (_blocked(box, others, sh, zones, lines, members) + _own_lines_pen(lb, others)
                   + r * 0.05)                                   # ближе — лучше
            if pen < best[0] - 0.5:
                best = (pen, (br.X0, br.Y0), box)
    br.X0, br.Y0 = best[1]
    base.Update()
    lb["box"] = best[2]
    return start, best[0]


def fix_sheet(sh, lines=(), members=(), max_rounds=2):
    """Исправить прочитанный лист (drawcheck.read_sheet). → список сделанных исправлений."""
    done = []
    zones = _zones(sh)
    moved_names = set()
    for _ in range(max_rounds):
        moved = 0
        for lb in sh.labels:
            if lb["kind"] != "позиция":
                continue
            others = [o["box"] for o in sh.labels if o is not lb]
            # линии других выносок — тоже препятствие (их положение могло уже измениться)
            all_lines = list(lines) + [s for o in sh.labels if o is not lb and o["kind"] == "позиция"
                                       for s in leader_lines(o)]
            if (_blocked(lb["box"], others, sh, zones, all_lines, members)
                    + _own_lines_pen(lb, others)) < 1.0:
                continue                                         # и так чисто
            res = _move_leader(lb, sh, others, zones, all_lines, members)
            if res and res[1] < res[0] - 0.5:
                moved += 1
                moved_names.add(lb["text"])
        if not moved:
            break
    done += [f"выноска «{t}» перенесена на свободное место" for t in sorted(moved_names, key=str)]
    # надписи за рамкой листа (например подпись вида у высокой рамы)
    w, h = sh.fmt or (420, 297)
    for lb in sh.labels:
        if lb["kind"] != "надпись":
            continue
        b = lb["box"]
        dy = (h - 6) - b[3] if b[3] > h - 6 else (6 - b[1] if b[1] < 6 else 0)
        dx = (w - 6) - b[2] if b[2] > w - 6 else (21 - b[0] if b[0] < 21 else 0)
        if dx or dy:
            t = lb["obj"]
            k = lb["view"][0]
            t.X, t.Y = t.X + dx / k, t.Y + dy / k
            t.Update()
            done.append(f"надпись «{lb['text']}» сдвинута внутрь рамки")
    # сечения с лишними нулями
    for tb in sh.tables:
        t = CastTo(tb["obj"], "ITable")
        for r, row in enumerate(tb["rows"]):
            for c, val in enumerate(row):
                if dc.SECTION_ZEROS.search(val or ""):
                    new = dc.re.sub(r"(\d+)[.,]0+(?=\D|$)", r"\1", val)
                    CastTo(CastTo(t.Cell(r, c), "ITableCell").Text, "IText").Str = new
                    done.append(f"сечение «{val}» → «{new}»")
        CastTo(tb["obj"], "IDrawingTable").Update()
    return done


def remaining(sh):
    """Замечания, оставшиеся после исправления (без проверки по модели — она отдельно)."""
    return dc.check_overlaps(sh) + dc.check_bounds(sh) + dc.check_tables(sh)


def fix_document(api7, doc2d, bodies=None):
    """Проверить и исправить открытый чертёж (не сохраняет). → (исправления, оставшиеся замечания).
    bodies — тела рамы, если уже прочитаны (при построении), иначе модель читается заново."""
    sh = dc.read_sheet(api7, doc2d)
    lines, members = read_geometry(api7, doc2d, bodies)
    done = fix_sheet(sh, lines, members)
    if done:
        sh = dc.read_sheet(api7, doc2d)            # перечитать: точные габариты после правок
    return done, remaining(sh)


def remarks_text(rec):
    """Оставшиеся замечания одной строкой-перечнем (для итогового окна)."""
    return "; ".join(f"{title.lower()} — {len(items)}" for title, items, _ in rec)


def fix_file(api7, cdw):
    """Исправить чертёж на диске: открыть, исправить, сохранить (+ PDF рядом, если он был)."""
    opened = fd.find_open_doc(api7, cdw)
    doc = opened or fd.need(api7.Documents.Open(cdw, False, False), "открыть чертёж")
    try:
        done, rec = fix_document(api7, CastTo(doc, "IKompasDocument2D"))
        if done:
            CastTo(doc, "IKompasDocument").Save()
    finally:
        if opened is None:
            doc.Close(0)
    pdf = os.path.splitext(cdw)[0] + ".pdf"
    if done and os.path.isfile(pdf):
        fd.export_pdf(api7, cdw, pdf)              # PDF — из исправленного чертежа
    return done, rec


def run(src, rot=None, gui=False):
    files = dc.list_drawings(os.path.abspath(src))
    if not files:
        raise RuntimeError(f"Нет чертежей .cdw: {src}")
    api7, started = fd.connect_kompas(rot)
    hide = api7.HideMessage
    api7.HideMessage = 1
    lines, fixed, left = [], 0, 0
    try:
        for f in files:
            try:
                done, rec = fix_file(api7, f)
                fixed += bool(done)
                left += bool(rec)
                lines.append(f"• {os.path.basename(f)}: исправлено {len(done)}"
                             + (f"; осталось: {remarks_text(rec)}" if rec else "; замечаний нет")
                             + f"\n   Путь: {f}")
            except Exception as e:
                lines.append(f"• {os.path.basename(f)}\n   Путь: {f}\n   Причина: {e}")
                traceback.print_exc()
    finally:
        api7.HideMessage = hide
        if started:
            api7.Quit()
    text = f"Чертежей: {len(files)}, исправлено: {fixed}, с оставшимися замечаниями: {left}\n\n" + "\n".join(lines[:25])
    print(text)
    if gui:
        dc.msgbox(text)
    return text


if __name__ == "__main__":
    fd.kpreport.start_log("Исправление чертежа")
    try:
        run(sys.argv[1])
    except Exception:
        traceback.print_exc()
        raise
