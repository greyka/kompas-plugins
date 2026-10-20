# -*- coding: utf-8 -*-
"""Пакетная генерация чертежей рам из сборок КОМПАС-3D (.a3d → .cdw + .pdf).

Использование:
    python frame2cdw.py <файл.a3d | папка> [--out <папка>] [--no-pdf]
    (по умолчанию результат — в подпапку «Чертежи» рядом с исходниками)

Для каждой сборки: ассоциативный вид рамы, ординатные отметки (оси стоек по
горизонтали, низ ригелей по вертикали), габариты, диагональ, позиционные выноски,
таблица профилей (позиция / сечение / кол-во / длина), основная надпись.
"""
import argparse
import math
import os
import re
import sys
import traceback
from collections import OrderedDict

from win32com.client import CastTo, gencache

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(os.path.dirname(HERE), "_setup"))
import kpsettings  # noqa: E402  настройки (settings.ini) и версия плагинов
import report as kpreport  # noqa: E402  журнал и отчёт об ошибке в Telegram

CFG = kpsettings.load()
TITLE = f"Чертежи рам — версия {kpsettings.version()}"

# --- константы API7 -------------------------------------------------------
ksDocumentDrawing = 1
vt_Arbitrary = 2
ksLinDParallel, ksLinDHorizontal, ksLinDVertical = 0, 1, 2

SCALES = [1, 2, 2.5, 4, 5, 10, 15, 20, 25, 40, 50, 75, 100]  # знаменатели 1:N
# формат: (код ksFormatA*, ширина, высота) — альбомная ориентация
FORMATS = [(3, 420, 297), (2, 594, 420), (1, 841, 594)]  # A3, A2, A1
TABLE_W = 110  # ширина таблицы профилей, мм листа
STAMP_H = 55   # высота основной надписи
# сторона → (подпись вида, зеркальная ли проекция относительно оси X модели)
SIDES = {"left": ("Вид слева", False), "right": ("Вид справа", True)}


def side_folder(side):
    """Имя подпапки для вида — из настроек (Общие → Подпапка «вид слева/справа»)."""
    return CFG.get("general", "left_folder" if side == "left" else "right_folder") or SIDES[side][0]


def problem(path, reason, side=None):
    """Строка для итогового окна и журнала: имя файла, полный путь, причина."""
    if isinstance(reason, Exception):
        text = str(reason).strip() or type(reason).__name__
        if not isinstance(reason, (KompasBusy, NotAFrame, RuntimeError)):
            text = f"{type(reason).__name__}: {text}"
    else:
        text = reason
    view = f" [{SIDES[side][0]}]" if side else ""
    return f"• {os.path.basename(path)}{view}\n   Путь: {path}\n   Причина: {text}"


def problems_text(title, items, limit=12):
    if not items:
        return ""
    more = f"\n…и ещё {len(items) - limit} (полностью — в журнале)" if len(items) > limit else ""
    return f"\n\n{title} ({len(items)}):\n" + "\n".join(items[:limit]) + more


def choose_sides(section, title):
    """Какие виды строить: окно выбора (если в настройках «Спрашивать при каждом построении»),
    иначе — из настроек. → 'left' | 'right' | 'left,right' или None (нажата «Отмена»)."""
    import subprocess
    if not CFG.get(section, "ask_sides") or not os.path.isfile(kpsettings.MANAGER_EXE):
        return CFG.sides(section)
    r = subprocess.run([kpsettings.MANAGER_EXE, "--ask-sides", f"--section={section}", f"--title={title}"],
                       cwd=kpsettings.ROOT).returncode
    return {1: "left", 2: "right", 3: "left,right"}.get(r)


def out_folder_name():
    return CFG.get("general", "out_folder") or "Чертежи"
PROJECTIONS = ["#Сверху", "#Снизу", "#Спереди", "#Сзади", "#Слева", "#Справа"]


# --- чтение модели ----------------------------------------------------------
def prop_values(api7, doc):
    pm = CastTo(api7, "IPropertyMng")
    return {CastTo(p, "IProperty").Name: CastTo(p, "IProperty") for p in pm.GetProperties(doc)}


def get_prop(obj, prop):
    try:
        return CastTo(obj, "IPropertyKeeper").GetPropertyValue(prop, None, True, True)[1]
    except Exception:
        return None


# коды условных знаков сечения в свойстве «Сечение» (@131… — замкнутый, @016… — полоса)
SECTION_SIGNS = {"131": "□", "016": "—"}


def parse_section(name):
    """Запасной разбор имени тела: '…@/Профиль $d 60x40x4 ГОСТ…@/L = 3200 мм'."""
    m = re.search(r"@/(\S+)\s+(?:\$d\s+)?(?:Д1 Н ПР\s+)?(\S+).*?L\s*=\s*([\d.,]+)", name or "")
    if not m:
        return None, None
    return m.group(2), float(m.group(3).replace(",", "."))


def section_text(raw, name):
    """'@13160х40х4' → '□60x40x4'."""
    if raw:
        m = re.match(r"@(\d{3})(.*)", raw)
        sign, sec = (SECTION_SIGNS.get(m.group(1), ""), m.group(2)) if m else ("", raw)
    else:
        sign, sec = "", parse_section(name)[0] or ""
    return sign + sec.replace("х", "x")


def find_open_doc(api7, path):
    docs = api7.Documents
    for i in range(docs.Count):
        d = CastTo(docs.Item(i), "IKompasDocument")
        if os.path.normcase(d.PathName or "") == os.path.normcase(path):
            return d
    return None


class KompasBusy(RuntimeError):
    """КОМПАС вернул пустой объект (занят, показал своё окно) — имеет смысл подождать и повторить."""


class NotAFrame(RuntimeError):
    """В сборке нет профилей металлоконструкций (сборка блока, пол, потолок…) — пропускаем, не ошибка."""


RETRY_WAITS = (2, 5, 10)   # секунды между повторами, если КОМПАС не ответил


def need(obj, what, hint="КОМПАС был занят или показал своё окно"):
    if obj is None:
        raise KompasBusy(f"КОМПАС не выполнил «{what}» ({hint})")
    return obj


def with_retry(fn, *args, **kw):
    """Вызов с повторами при KompasBusy: 19.10 у Димы КОМПАС ~минуту отвечал «пусто» на открытие
    и создание документов, потом сам пришёл в норму — без повтора выпало 9 сборок подряд."""
    import time
    for wait in RETRY_WAITS + (None,):
        try:
            return fn(*args, **kw)
        except KompasBusy as e:
            if wait is None:
                raise KompasBusy(f"{e}; повтор {len(RETRY_WAITS)} раза не помог") from None
            print(f"  {e} — повтор через {wait} с")
            time.sleep(wait)


def read_frame(api7, path):
    return with_retry(_read_frame, api7, path)


def _read_frame(api7, path):
    """Читает уже открытый в КОМПАС документ (и не закрывает его) или открывает файл скрыто."""
    opened = find_open_doc(api7, path)
    doc = opened or need(api7.Documents.Open(path, False, True), "открыть сборку",
                         "файл повреждён, сохранён в более новой версии КОМПАС, открыт у другого "
                         "пользователя — или КОМПАС был занят")
    try:
        top = CastTo(doc, "IKompasDocument3D").TopPart
        P = prop_values(api7, doc)
        info = {
            "name": get_prop(top, P["Наименование"]) or "",
            "marking": get_prop(top, P["Обозначение"]) if "Обозначение" in P else "",
            "mass": get_prop(top, P["Масса"]) if "Масса" in P else None,
        }
        prop = lambda o, k: get_prop(o, P[k]) if k in P else None
        bodies = []
        for b in CastTo(top, "IFeature7").ResultBodies or []:
            b = CastTo(b, "IBody7")
            ok, x1, y1, z1, x2, y2, z2 = b.GetGabarit()
            Lm = prop(b, "Длина профиля")
            if not Lm:          # не профиль металлоконструкций (пластина, вспомогательное тело)
                continue
            pos = prop(b, "Позиция")
            bodies.append(dict(name=b.Name, sec=section_text(prop(b, "Сечение"), b.Name),
                               L=round(float(Lm) * 1000, 1),
                               pos=int(pos) if pos not in (None, "") else None,
                               box=(x1, y1, z1, x2, y2, z2)))
        if not bodies:
            raise NotAFrame("в сборке нет профилей металлоконструкций — не рама, пропущено")
        # стандартные изделия / детали первого уровня (крепёж и т.п.) — в таблицу
        hardware = OrderedDict()
        for i in range(top.Parts.Count):
            pt = CastTo(top.Parts.Item(i), "IPart7")
            pos = prop(pt, "Позиция")
            key = (int(pos) if pos not in (None, "") else None, prop(pt, "Наименование") or pt.Name)
            hardware[key] = hardware.get(key, 0) + 1
        info["hardware"] = hardware
        return info, bodies
    finally:
        if opened is None:
            doc.Close(0)


# --- геометрия рамы ----------------------------------------------------------
def frame_axes(bodies):
    """Определить плоскость рамы: две оси с наибольшим размахом (u — горизонталь, v — вертикаль)."""
    lo = [min(b["box"][i] for b in bodies) for i in range(3)]
    hi = [max(b["box"][i + 3] for b in bodies) for i in range(3)]
    span = [hi[i] - lo[i] for i in range(3)]
    thin = span.index(min(span))
    a, c = [i for i in range(3) if i != thin]
    # вертикаль — Z, если рама стоит (Z в плоскости), иначе меньшая из двух
    v = 2 if 2 in (a, c) else (a if span[a] < span[c] else c)
    u = c if v == a else a
    return u, v, thin, lo, hi


def classify(bodies, u, v):
    posts, rails, braces = [], [], []
    for b in bodies:
        x1, x2 = b["box"][u], b["box"][u + 3]
        y1, y2 = b["box"][v], b["box"][v + 3]
        du, dv = x2 - x1, y2 - y1
        if dv > 4 * du:
            posts.append(b)
        elif du > 4 * dv:
            rails.append(b)
        else:
            braces.append(b)
    return posts, rails, braces


# --- чертёж ------------------------------------------------------------------
def pick_layout(W, H):
    """Наименьший лист (не меньше «Наименьший формат») и крупнейший масштаб,
    но не мельче «Самый мелкий масштаб» (иначе — лист больше)."""
    max_scale = int(CFG.get("frame", "max_scale"))
    first = {"A3": 0, "A2": 1, "A1": 2}[CFG.get("frame", "min_format")]
    formats = FORMATS[first:]
    for limit in (max_scale, SCALES[-1]):
        for fmt, fw, fh in formats:
            aw = fw - 20 - 5 - TABLE_W - 50   # рамка + таблица + поля под размеры
            ah = fh - 10 - STAMP_H - 50
            for n in SCALES:
                if n <= limit and W / n <= aw and H / n <= ah:
                    return fmt, fw, fh, n
    fmt, fw, fh = FORMATS[-1]
    return fmt, fw, fh, SCALES[-1]


def merge_close(values, tol=1.0):
    """Отметки, отличающиеся меньше tol мм (разные стороны рамы), считаем одной."""
    out = []
    for x in sorted(values):
        if not out or x - out[-1] >= tol:
            out.append(x)
    return out


def add_line_dim(cont, p1, p2, p3, orient, text=None, shift=0.0):
    """shift — смещение текста вдоль размерной линии, мм листа (0 — по центру)."""
    d = CastTo(cont.LineDimensions.Add(), "ILineDimension")
    d.X1, d.Y1 = p1
    d.X2, d.Y2 = p2
    d.X3, d.Y3 = p3
    d.Orientation = orient
    if shift:
        dp = CastTo(d, "IDimensionParams")
        dp.ArrowPos = 1  # ksDimArrowOutside
        dp.TextPos = shift
    if text is not None:
        t = CastTo(d, "IDimensionText")
        t.AutoNominalValue = False
        t.NominalValue = text
    d.Update()
    return d


def ordinate_marks(view, items, side, edge, scale, gap=2.0, tick=12.0, text_h=3.5):
    """Ординатные отметки как на образце: засечка + значение, разведение близких подписей.

    items — [((xv, yv), значение)], side — 'bottom' | 'right', edge — координата края рамы в виде.
    Все длины gap/tick/text_h — в мм листа.
    """
    k = 1.0 / scale
    dc = CastTo(view, "IDrawingContainer")
    items = sorted(items, key=lambda it: it[0][0] if side == "bottom" else it[0][1])
    slot_prev = None
    for (xv, yv), val in items:
        c = xv if side == "bottom" else yv               # координата вдоль края
        slot = c if slot_prev is None else max(c, slot_prev + (text_h + 3) * k)
        slot_prev = slot
        e0, e1, e2 = edge, edge, edge
        sgn = -1 if side == "bottom" else 1
        e0 += sgn * gap * k            # начало засечки
        e1 += sgn * (gap + 3) * k      # излом (если подпись сдвинута)
        e2 += sgn * (gap + 3 + tick) * k
        pts = [(c, e0), (c, e1), (slot, e1 + sgn * 3 * k), (slot, e2)] if abs(slot - c) > 1e-6             else [(c, e0), (c, e2)]
        for (a1, b1), (a2, b2) in zip(pts, pts[1:]):
            ls = CastTo(dc.LineSegments.Add(), "ILineSegment")
            if side == "bottom":
                ls.X1, ls.Y1, ls.X2, ls.Y2 = a1, b1, a2, b2
            else:
                ls.X1, ls.Y1, ls.X2, ls.Y2 = b1, a1, b2, a2
            ls.Style = 2  # тонкая
            ls.Update()
        t = CastTo(dc.DrawingTexts.Add(), "IDrawingText")
        if side == "bottom":
            t.X, t.Y, t.Angle = slot - 1.8 * k, e2, 90
            t.Allocation = 0
        else:
            t.X, t.Y, t.Angle = edge + (gap + 0.5) * k + (3 * k if abs(slot - c) > 1e-6 else 0), slot + 1.8 * k, 0
            t.Allocation = 0
        CastTo(t, "IText").Str = fmt_len(val)
        t.Update()


LEADER_OFFSETS = [(12, 10), (-22, 10), (12, -12), (-22, -12), (12, 20), (-22, 20), (12, -22), (-22, -22)]


def add_pos_leader(cont, pt, pos, scale, placed):
    """Выноска с полкой; placed — занятые прямоугольники полок (мм листа), чтобы не накладывались."""
    k = 1.0 / scale
    sx, sy = pt[0] * scale, pt[1] * scale
    dx, dy = LEADER_OFFSETS[0]
    for ox, oy in LEADER_OFFSETS:
        box = (sx + ox - 2, sy + oy - 1, sx + ox + 14, sy + oy + 8)
        if not any(box[0] < q[2] and q[0] < box[2] and box[1] < q[3] and q[1] < box[3] for q in placed):
            dx, dy = ox, oy
            break
    placed.append((sx + dx - 2, sy + dy - 1, sx + dx + 14, sy + dy + 8))
    ld = cont.Leaders.Add(21)  # ksDrPosLeader
    br = CastTo(ld, "IBranchs")
    br.X0, br.Y0 = pt[0] + dx * k, pt[1] + dy * k
    br.AddBranchByPoint(-1, pt[0], pt[1])
    txt = CastTo(CastTo(ld, "IPositionLeader").Positions, "IText")
    txt.Str = str(pos)
    CastTo(ld, "IBaseLeader").Update()


def fmt_len(v):
    v = round(v, 1)
    return f"{v:.1f}".rstrip("0").rstrip(".").replace(".", ",")


def build_drawing(api7, a3d, info, bodies, out_cdw, side="left"):
    return with_retry(_build_drawing_safe, api7, a3d, info, bodies, out_cdw, side)


def _build_drawing_safe(api7, a3d, info, bodies, out_cdw, side):
    """Недостроенный чертёж при ошибке закрываем — иначе в КОМПАС копятся скрытые документы."""
    created = []
    try:
        return _build_drawing(api7, a3d, info, bodies, out_cdw, side, created)
    except Exception:
        for d in created:
            try:
                d.Close(0)
            except Exception:
                pass
        raise


def _build_drawing(api7, a3d, info, bodies, out_cdw, side, created):
    u, v, thin, lo, hi = frame_axes(bodies)
    W, H = hi[u] - lo[u], hi[v] - lo[v]
    fmt, fw, fh, n = pick_layout(W, H)
    scale = 1.0 / n

    doc = need(api7.Documents.Add(ksDocumentDrawing, False), "создать чертёж")
    created.append(CastTo(doc, "IKompasDocument"))
    doc2d = CastTo(doc, "IKompasDocument2D")
    sheet = CastTo(doc2d.LayoutSheets.Item(0), "ILayoutSheet")
    sf = CastTo(sheet.Format, "ISheetFormat")
    sf.Format = fmt
    sf.VerticalOrientation = False
    sheet.Update()

    views = CastTo(doc2d.ViewsAndLayersManager, "IViewsAndLayersManager").Views
    title, mirrored = SIDES[side]
    pname = pick_projection(views, a3d, lo, hi, u, v, thin, mirrored)

    view = CastTo(need(views.Add(vt_Arbitrary), "создать вид"), "IAssociationView")
    view.SourceFileName = a3d
    view.ProjectionName = pname
    view.Scale = scale
    view.HiddenLinesVisible = False
    view.CenterLinesVisible = False
    vd = CastTo(view, "IViewDesignation")
    vd.ShowScale = False
    vd.ShowName = False
    view.Update()
    best = (pname,)

    def proj(pu, pv):
        """Координаты модели (u, v) → координаты вида (мм модели)."""
        return tuple(view.GetProjectionPoint(*_abs(lo, u, v, thin, pu, pv))[-2:])

    # разместить вид: левый нижний угол рамы
    x0, y0 = proj(lo[u], lo[v])
    x1, y1 = proj(hi[u], hi[v])
    left, right = min(x0, x1), max(x0, x1)
    bottom, top_ = min(y0, y1), max(y0, y1)
    area_l, area_b = 20 + 30, 5 + STAMP_H + 25
    area_r = fw - 5 - TABLE_W - 20
    view.X = area_l + ((area_r - area_l) - (right - left) * scale) / 2 - left * scale
    view.Y = area_b + 10 - bottom * scale
    view.Update()

    sysdc = CastTo(views.View(0), "IDrawingContainer")
    if CFG.get("frame", "show_title"):
        t = CastTo(sysdc.DrawingTexts.Add(), "IDrawingText")
        t.X = view.X + (left + right) / 2 * scale
        t.Y = min(view.Y + top_ * scale + 32, fh - 14)  # выше полок выносок, но внутри рамки
        t.Allocation = 1  # по центру
        CastTo(t, "IText").Str = title
        t.Update()

    cont = CastTo(view, "ISymbols2DContainer")
    off = 10 / scale  # 10 мм листа в мм модели

    posts, rails, braces = classify(bodies, u, v)

    # ординатные отметки снизу (оси стоек) и справа (низ ригелей)
    ys = sorted({round(lo[v], 1), round(hi[v], 1)} |
                {round(b["box"][v], 1) for b in rails
                 if lo[v] + 1 < b["box"][v] and b["box"][v + 3] < hi[v] - 1})
    left_u = lo[u] if proj(lo[u], 0)[0] < proj(hi[u], 0)[0] else hi[u]
    right_u = hi[u] if left_u == lo[u] else lo[u]
    # стойки — от начала профиля: грань, ближняя к нулевой (левой на виде) кромке рамы
    start = (lambda b: b["box"][u]) if left_u == lo[u] else (lambda b: b["box"][u + 3])
    if CFG.get("frame", "ordinate_from") == "center":
        start = lambda b: (b["box"][u] + b["box"][u + 3]) / 2  # noqa: E731
    xs = sorted({round(lo[u], 1), round(hi[u], 1)} |
                {round(start(b), 1) for b in posts
                 if lo[u] + 1 < b["box"][u] and b["box"][u + 3] < hi[u] - 1})
    xs, ys = merge_close(xs), merge_close(ys)
    ordinate_marks(view, [(proj(x, lo[v]), abs(x - left_u)) for x in xs], "bottom", bottom, scale)
    ordinate_marks(view, [(proj(right_u, y), y - lo[v]) for y in ys], "right", right, scale)

    # габариты
    mark_len = 22 / scale
    if CFG.get("frame", "overall_dims"):
        pa, pb = proj(lo[u], lo[v]), proj(hi[u], lo[v])
        add_line_dim(cont, pa, pb, (pa[0], bottom - mark_len - off), ksLinDHorizontal)
        pa, pb = proj(right_u, lo[v]), proj(right_u, hi[v])
        add_line_dim(cont, pa, pb, (right + mark_len + off, pa[1]), ksLinDVertical)

    # позиционные выноски — по одной на позицию, на первом теле
    seen, placed = set(), []
    cu0, cv0 = (lo[u] + hi[u]) / 2, (lo[v] + hi[v]) / 2
    def centrality(b):
        bx = b["box"]
        return abs((bx[u] + bx[u + 3]) / 2 - cu0) + abs((bx[v] + bx[v + 3]) / 2 - cv0)
    for bd in sorted(bodies, key=lambda b: (b["pos"] or 0, centrality(b))):
        if not CFG.get("frame", "leaders"):
            break
        if bd["pos"] is None or bd["pos"] in seen:
            continue
        seen.add(bd["pos"])
        bx = bd["box"]
        cu, cv = (bx[u] + bx[u + 3]) / 2, (bx[v] + bx[v + 3]) / 2
        f = (0.3, 0.5, 0.7, 0.4, 0.6)[len(seen) % 5]   # точка вдоль самого элемента
        if bd in posts:
            cv = bx[v] + (bx[v + 3] - bx[v]) * f
        elif bd in rails:
            cu = bx[u] + (bx[u + 3] - bx[u]) * f
        try:
            add_pos_leader(cont, proj(cu, cv), bd["pos"], scale, placed)
        except Exception as e:
            print("  выноска", bd["pos"], e)

    # диагональ
    if CFG.get("frame", "diagonal"):
        pa, pb = proj(lo[u], lo[v]), proj(hi[u], hi[v])
        add_line_dim(cont, pa, pb, ((pa[0] + pb[0]) / 2, (pa[1] + pb[1]) / 2), ksLinDParallel)

    # таблица профилей
    groups = OrderedDict()
    for b in sorted(bodies, key=lambda b: (b["pos"] or 999, b["sec"], -(b["L"] or 0))):
        key = (b["pos"], b["sec"], b["L"])
        groups[key] = groups.get(key, 0) + 1
    if CFG.get("frame", "hardware"):
        for (pos, name), cnt in sorted(info.get("hardware", {}).items(),
                                       key=lambda kv: (kv[0][0] or 999, kv[0][1])):
            groups[(pos, name, None)] = cnt
    if CFG.get("frame", "table"):
        add_table(doc2d, groups, fw)

    # основная надпись
    try:
        stamp = CastTo(sheet.Stamp, "IStamp")
        CastTo(stamp.Text(1), "IText").Str = info["name"]
        if info.get("marking"):
            CastTo(stamp.Text(2), "IText").Str = info["marking"]
        CastTo(stamp.Text(6), "IText").Str = f"1:{n:g}"
        # ячейки основной надписи: 110 — Разраб., 111 — Пров., 9 — организация
        for cell, key in ((110, "designer"), (111, "checker"), (9, "organization")):
            val = CFG.get("frame", key)
            if val:
                CastTo(stamp.Text(cell), "IText").Str = val
        stamp.Update()
    except Exception:
        pass

    CastTo(doc2d, "IKompasDocument").SaveAs(out_cdw)
    return doc2d, dict(W=W, H=H, scale=n, fmt=fmt, projection=best[0],
                       posts=len(posts), rails=len(rails), braces=len(braces))


def pick_projection(views, a3d, lo, hi, u, v, thin, mirrored=False):
    """Временный вид: проекция, где рама видна плоско и вертикаль вверх.
    mirrored=False — ось X модели направлена вправо (вид слева), True — влево (вид справа)."""
    probe = CastTo(need(views.Add(vt_Arbitrary), "создать вид"), "IAssociationView")
    probe.SourceFileName = a3d
    found = None
    try:
        for pname in PROJECTIONS:
            probe.ProjectionName = pname
            probe.Update()
            p00, p10, p01 = (tuple(probe.GetProjectionPoint(*_pt(lo, hi, u, v, thin, a, b))[-2:])
                             for a, b in ((0, 0), (1, 0), (0, 1)))
            flat = abs(p10[0] - p00[0]) > 1 and abs(p01[1] - p00[1]) > 1
            if flat and abs(p10[1] - p00[1]) < 1e-3 and abs(p01[0] - p00[0]) < 1e-3 and p01[1] > p00[1]:
                if (p10[0] < p00[0]) == mirrored:
                    return pname
                found = found or pname
    finally:
        probe.Delete()
    if not found:
        raise RuntimeError("не удалось подобрать проекцию для плоскости рамы")
    return found


def _abs(lo, u, v, thin, pu, pv):
    p = [0.0, 0.0, 0.0]
    p[u], p[v], p[thin] = pu, pv, lo[thin]
    return p


def _pt(lo, hi, u, v, thin, fu, fv):
    return _abs(lo, u, v, thin, hi[u] if fu else lo[u], hi[v] if fv else lo[v])


def add_table(doc2d, groups, fw):
    views = CastTo(doc2d.ViewsAndLayersManager, "IViewsAndLayersManager").Views
    sysview = views.View(0)
    cont = CastTo(sysview, "ISymbols2DContainer")
    rows = len(groups) + 1
    widths = [15, 35, 25, 35]
    tbl = CastTo(cont.DrawingTables.Add(rows, len(widths), 8, sum(widths) / len(widths), 0), "IDrawingTable")
    t = CastTo(tbl, "ITable")
    head = ["Поз.", "Сечение / наименование", "Кол-во", "Длина, мм"]
    for c, h in enumerate(head):
        CastTo(CastTo(t.Cell(0, c), "ITableCell").Text, "IText").Str = h
    for r, ((pos, sec, L), cnt) in enumerate(groups.items(), start=1):
        vals = [pos if pos is not None else "", sec, cnt, fmt_len(L) if L else ""]
        for c, val in enumerate(vals):
            CastTo(CastTo(t.Cell(r, c), "ITableCell").Text, "IText").Str = str(val)
    tbl.X = fw - 5 - sum(widths)
    tbl.Y = 5 + STAMP_H + 5 + 8 * rows
    tbl.Update()


def export_pdf(api7, cdw, pdf):
    conv = api7.Converter(os.path.join(kompas_bin(), "Pdf2d.dll"))
    if conv is None:
        raise RuntimeError("конвертер Pdf2d.dll не загрузился")
    return CastTo(conv, "IConverter").Convert(cdw, pdf, 0, False)


def kompas_bin():
    import winreg
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r"KOMPAS.Application.7\CLSID") as k:
        clsid = winreg.QueryValue(k, None)
    with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, rf"CLSID\{clsid}\LocalServer32") as k:
        exe = winreg.QueryValue(k, None)
    return os.path.dirname(exe.split('"')[1] if exe.startswith('"') else exe.split(" /")[0])


# --- main --------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("src", nargs="?")
    ap.add_argument("--out")
    ap.add_argument("--no-pdf", action="store_true")
    ap.add_argument("--keep-open", action="store_true")
    ap.add_argument("--sides", help="какие виды делать: left,right (по умолчанию — из настроек)")
    ap.add_argument("--open", action="store_true", help="открыть готовый чертёж в КОМПАС")
    ap.add_argument("--gui", action="store_true", help="итог и ошибки — окном (для запуска из меню)")
    ap.add_argument("--menu", type=int, help="команда из меню библиотеки КОМПАС (1..3)")
    ap.add_argument("--rot", help="имя КОМПАС в ROT (передаёт плагин)")
    a = ap.parse_args()
    kpreport.start_log("Чертежи рам")
    a.sides_explicit = a.sides is not None     # задано в командной строке — не спрашиваем
    a.sides = a.sides or CFG.sides("frame")
    a.no_pdf = a.no_pdf or not CFG.get("frame", "pdf")
    if a.menu:
        a.gui = True
    try:
        if a.menu:
            menu(a)
        else:
            run(a)
    except Exception:
        if a.gui:
            kpreport.error_dialog(TITLE, traceback.format_exc())
        raise


def ask_sides(a):
    """Окно «Какие виды строить?» перед построением. False — нажата «Отмена»."""
    if a.sides_explicit:
        return True
    s = choose_sides("frame", TITLE)
    if not s:
        return False
    a.sides = s
    return True


def menu(a):
    """Команды библиотеки «Чертежи рам»: 1 — текущая сборка, 2 — папка, 3 — открыть папку чертежей."""
    import ctypes
    api7, _ = connect_kompas(a.rot)
    a.api7 = api7
    doc = api7.ActiveDocument
    doc = CastTo(doc, "IKompasDocument") if doc is not None else None
    path = doc.PathName if doc is not None else ""
    kpreport.set_document(path)

    if a.menu == 9:
        kpsettings.open_manager("frame")
        return
    if a.menu == 1:
        if not path.lower().endswith(".a3d"):
            msgbox("Откройте и сохраните сборку рамы (.a3d), затем повторите команду.")
            return
        if doc.Changed:
            r = ctypes.windll.user32.MessageBoxW(
                None, "Сборка изменена. Сохранить перед построением чертежа?\n"
                      "(вид чертежа строится по файлу на диске)", TITLE, 0x23)
            if r == 2:      # Отмена
                return
            if r == 6:      # Да
                doc.Save()
        if not ask_sides(a):
            return
        a.src, a.open = path, CFG.get("frame", "open_after")
        run(a)
    elif a.menu == 2:
        folder = pick_folder(f"Папка со сборками рам (.a3d). Чертежи появятся в подпапке «{out_folder_name()}».")
        if not folder or not ask_sides(a):
            return
        a.src = folder
        run(a)
    elif a.menu == 3:
        if not path:
            msgbox("Нет активного сохранённого документа.")
            return
        d = os.path.join(os.path.dirname(path), out_folder_name())
        os.startfile(d if os.path.isdir(d) else os.path.dirname(path))


def pick_folder(title):
    """Стандартный диалог выбора папки Windows (без tkinter — его нет в переносном Python)."""
    from win32com.shell import shell, shellcon
    pidl, _, _ = shell.SHBrowseForFolder(
        0, None, title, shellcon.BIF_RETURNONLYFSDIRS | 0x40)  # 0x40 = BIF_NEWDIALOGSTYLE
    if not pidl:
        return None
    path = shell.SHGetPathFromIDList(pidl)
    return path.decode("mbcs") if isinstance(path, bytes) else path


def connect_kompas(rot_name=None):
    """Подключиться к КОМПАС.

    rot_name — имя, под которым плагин опубликовал «свой» КОМПАС в ROT (запуск из меню):
    так скрипт попадает именно в то окно, откуда вызвана команда, даже если КОМПАСов несколько.
    Иначе — к уже запущенному, иначе запускается новый (скрытый) экземпляр.
    Возвращает (api7, started) — started=True, если КОМПАС запущен скриптом и его надо закрыть.
    """
    import pythoncom, win32com.client
    if rot_name:
        rot, ctx = pythoncom.GetRunningObjectTable(), pythoncom.CreateBindCtx(0)
        for mk in rot:
            if mk.GetDisplayName(ctx, None).lstrip("!") == rot_name:
                obj = rot.GetObject(mk).QueryInterface(pythoncom.IID_IDispatch)
                return gencache.EnsureDispatch(obj), False
        raise RuntimeError("Не найден КОМПАС, из которого вызвана команда. Перезапустите КОМПАС.")
    try:
        return gencache.EnsureDispatch(win32com.client.GetActiveObject("Kompas.Application.7")), False
    except Exception:
        pass
    # Внимание: CreateObject может подключиться к уже открытому КОМПАС пользователя,
    # даже если его нет в ROT. «Свой» экземпляр — только если до этого КОМПАС не был запущен.
    running = kompas_running()
    api7 = gencache.EnsureDispatch("Kompas.Application.7")
    return api7, not running and not api7.Visible


def kompas_running():
    import subprocess
    out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq KOMPAS.exe", "/NH"],
                         capture_output=True, text=True, creationflags=0x08000000).stdout
    return "kompas.exe" in out.lower()


def msgbox(text, error=False):
    import ctypes
    ctypes.windll.user32.MessageBoxW(None, text, TITLE, 0x10 if error else 0x40)


def run(a):
    src = os.path.abspath(a.src)
    files = [src] if os.path.isfile(src) else [
        os.path.join(src, f) for f in sorted(os.listdir(src)) if f.lower().endswith(".a3d")]
    out = os.path.abspath(a.out or os.path.join(src if os.path.isdir(src) else os.path.dirname(src),
                                                out_folder_name()))
    os.makedirs(out, exist_ok=True)

    if not files:
        raise RuntimeError(f"Нет файлов .a3d: {src}")
    api7, started = getattr(a, "api7", None), False
    if api7 is None:
        api7, started = connect_kompas(getattr(a, "rot", None))
    hide_prev = api7.HideMessage
    api7.HideMessage = 1
    ok = fail = 0
    errors, last_cdw = [], None
    try:
        ok, fail, errors, last_cdw = process(api7, a, files, out)
    finally:
        api7.HideMessage = hide_prev
        if started and not a.open:
            api7.Quit()          # КОМПАС запускали сами — не оставлять висеть скрытым
            del api7
            import time
            for _ in range(30):  # дождаться выхода, иначе следующий запуск примет его за чужой
                if not kompas_running():
                    break
                time.sleep(1)
    if a.open and last_cdw:
        for c in last_cdw:
            api7.Documents.Open(c, True, False)
    skips = getattr(a, "skips", [])
    if a.gui and (fail or skips or len(files) > 1):  # одиночная сборка без проблем — без окна, чертежи откроются
        msg = f"Готово чертежей: {ok}, сборок: {len(files)}\nПапка: {out}"
        msg += problems_text("Ошибки — чертёж не построен", errors)
        msg += problems_text("Пропущено", skips)
        print(msg)
        if fail:
            kpreport.error_dialog(TITLE, msg)   # окно с кнопкой «Отправить разработчику»
        else:
            msgbox(msg)


def process(api7, a, files, out):
    """Для каждой сборки — чертёж на каждую сторону, в подпапки «Вид слева» / «Вид справа»."""
    sides = [x.strip() for x in a.sides.split(",") if x.strip()]
    ok = fail = 0
    errors, last_cdw = [], []
    a.skips = []
    for f in files:
        base = os.path.splitext(os.path.basename(f))[0]
        try:
            info, bodies = read_frame(api7, f)
        except NotAFrame as e:
            a.skips.append(problem(f, e))
            print(f"--  {base}: {e}\n    {f}")
            continue
        except Exception as e:
            fail += 1
            errors.append(problem(f, e))
            print(f"ERR {base}: {f}")
            traceback.print_exc()
            continue
        last_cdw = []
        for side in sides:
            sub = os.path.join(out, side_folder(side))
            os.makedirs(sub, exist_ok=True)
            cdw = os.path.join(sub, base + ".cdw")
            if os.path.exists(cdw) and not CFG.get("general", "overwrite"):
                a.skips.append(problem(f, f"чертёж уже есть ({cdw}), перезапись выключена в настройках", side))
                print(f"--  {base} [{SIDES[side][0]}]: уже есть, пропуск")
                continue
            try:
                doc2d, stat = build_drawing(api7, f, info, bodies, cdw, side)
                if not a.no_pdf:
                    pdf = os.path.join(sub, base + ".pdf")
                    try:
                        export_pdf(api7, cdw, pdf)
                        if not os.path.isfile(pdf):
                            raise RuntimeError("КОМПАС не создал файл")
                    except Exception as e:
                        a.skips.append(problem(f, f"PDF не создан ({pdf}): {e} — чертёж .cdw есть", side))
                        print(f"  pdf: {e}")
                if not a.keep_open:
                    doc2d.Close(0)
                bak = cdw + ".bak"
                if os.path.exists(bak):
                    os.remove(bak)
                print(f"OK  {base} [{SIDES[side][0]}]: {stat}")
                ok += 1
                last_cdw.append(cdw)
            except Exception as e:
                fail += 1
                errors.append(problem(f, e, side))
                print(f"ERR {base} [{SIDES[side][0]}]: {f}")
                traceback.print_exc()
    print(f"Готово: {ok} ок, {fail} ошибок")
    return ok, fail, errors, last_cdw


if __name__ == "__main__":
    main()
