# -*- coding: utf-8 -*-
"""Общий чертёж: несколько чертежей КОМПАС (.cdw) → один многолистовой .cdw.

Каждый лист исходного чертежа становится листом общего (тот же формат и ориентация).
Виды копируются как есть, со всеми ручными правками (DrawingGroup.CopyToDocument):
ассоциативный вид переносится как обычный вид (геометрия, размеры, выноски, тексты) —
связь с моделью в общем чертеже не нужна. Основная надпись копируется по ячейкам.

Почему не «вставка вида по ссылке»: конвертер КОМПАС не делает PDF из таких файлов,
а после «Разрушить» геометрия теряет масштаб вида (проверено 20.10).

Листы многолистового чертежа стоят по горизонтали подряд: начало листа i = сумма ширин
предыдущих листов (проверено GetSheetNumberByPoint).
"""
import os
import traceback

from win32com.client import CastTo

FORMATS_MM = {0: (1189, 841), 1: (841, 594), 2: (594, 420), 3: (420, 297), 4: (297, 210), 5: (210, 148)}
vt_Normal = 1


def sheet_size(sheet):
    sf = CastTo(CastTo(sheet, "ILayoutSheet").Format, "ISheetFormat")
    w, h = FORMATS_MM.get(sf.Format, (sf.FormatWidth, sf.FormatHeight))
    if sf.FormatMultiplicity and sf.FormatMultiplicity > 1:            # «A3×3» и т. п.
        w = w * sf.FormatMultiplicity if w < h else w
    w, h = max(w, h), min(w, h)
    return (h, w) if sf.VerticalOrientation else (w, h)


def copy_format(src_sheet, dst_sheet):
    s_sheet = CastTo(src_sheet, "ILayoutSheet")
    s = CastTo(s_sheet.Format, "ISheetFormat")
    d_sheet = CastTo(dst_sheet, "ILayoutSheet")
    # оформление (рамка + штамп) как у источника: новые листы по умолчанию получают
    # «последующий лист» (стиль 2) — у его штампа нет ячеек наименования, масштаба и т. п.
    try:
        d_sheet.LayoutLibraryFileName = s_sheet.LayoutLibraryFileName
        d_sheet.LayoutStyleNumber = s_sheet.LayoutStyleNumber
    except Exception:
        pass
    d = CastTo(d_sheet.Format, "ISheetFormat")
    d.Format, d.VerticalOrientation = s.Format, s.VerticalOrientation
    try:
        d.FormatMultiplicity = s.FormatMultiplicity
        if s.Format not in FORMATS_MM:                                  # пользовательский формат
            d.FormatWidth, d.FormatHeight = s.FormatWidth, s.FormatHeight
    except Exception:
        pass
    d_sheet.Update()


def copy_stamp(src_sheet, dst_sheet):
    s = CastTo(CastTo(src_sheet, "ILayoutSheet").Stamp, "IStamp")
    d = CastTo(CastTo(dst_sheet, "ILayoutSheet").Stamp, "IStamp")
    n = 0
    for cell in range(1, 400):
        if cell in (1, 5, 7, 8):           # наименование и масса — общие для документа (масса листа —
            continue                       # в подписи листа); «Лист»/«Листов» — свои
        try:
            if not s.IsCellPresent(cell) or not d.IsCellPresent(cell):
                continue
            txt = CastTo(s.Text(cell), "IText").Str
            if txt:
                CastTo(d.Text(cell), "IText").Str = txt
                n += 1
        except Exception:
            pass
    d.Update()
    return n


DESIGNATION_FLAGS = ("ShowName", "ShowScale", "ShowTurn", "ShowAngle", "ShowPage", "ShowZone", "ShowUnfold")


def copy_designation(src_view, dst_view):
    """Обозначение вида как у источника; у нового вида КОМПАС по умолчанию рисует
    знак поворота и т. п. (кружки со стрелками над рамой)."""
    d = CastTo(dst_view, "IViewDesignation")
    try:
        s = CastTo(src_view, "IViewDesignation")
    except Exception:
        s = None
    for flag in DESIGNATION_FLAGS:
        try:
            setattr(d, flag, getattr(s, flag) if s is not None else False)
        except Exception:
            pass


def label_sheet(dst2d, dviews, x0, sheet_h, text):
    """Подпись листа общего чертежа (имя исходного файла и наименование) — в верхнем левом углу
    поля чертежа: наименование в штампе многолистового документа общее для всех листов."""
    v = CastTo(dviews.View(0), "IView")
    v.Current = True
    t = CastTo(CastTo(v, "IDrawingContainer").DrawingTexts.Add(), "IDrawingText")
    t.X, t.Y = x0 + 25, 9           # внизу слева, у рамки — там обычно пусто (штамп справа)
    t.Allocation = 0
    CastTo(t, "IText").Str = text
    t.Update()


def copy_view_objects(src2d, src_view, dst_doc, dst2d, dst_view, dx=0.0):
    """Все объекты вида-источника → вид-приёмник (координаты вида сохраняются; dx — сдвиг,
    нужен только для системного вида, у которого нет собственного начала)."""
    objs = CastTo(src_view, "IDrawingContainer").Objects(0)        # ksAllObj
    objs = list(objs) if objs else []
    if not objs:
        return 0
    CastTo(dst_view, "IView").Current = True
    CastTo(src_view, "IView").Current = True
    g = CastTo(CastTo(src2d, "IKompasDocument2D1").DrawingGroups.Add(True, ""), "IDrawingGroup")
    g.AddObjects(objs)
    copied = CastTo(g.CopyToDocument(dst_doc), "IDrawingGroup")
    copied.Store()
    if dx:
        moved = copied.Objects(0)
        if moved:
            CastTo(dst2d, "IKompasDocument2D1").MoveObjects(list(moved), dx, 0)
    g.Close()
    drop_designation_texts(dst_view)
    return len(objs)


def drop_designation_texts(view):
    """Текст обозначения вида («^(#0~)…type=view…»): в источнике пустой, а в копии КОМПАС
    подставляет знак поворота «@64» — он рисуется кружками со стрелками над рамой и в её углу.
    Подписи видов у нас — обычные надписи, поэтому служебные тексты обозначения удаляем."""
    texts = CastTo(view, "IDrawingContainer").DrawingTexts
    for i in reversed(range(texts.Count)):
        t = CastTo(texts.DrawingText(i), "IDrawingText")
        try:
            if CastTo(t, "IText").Str.startswith("^(") and "type=view" in CastTo(t, "IText").Str:
                t.Delete()
        except Exception:
            pass


def sheet_names(items):
    """[(лист, файл, № листа источника|None, масса)] → [(лист, наименование, масса)].
    Наименование — имя файла; одинаковые имена из разных папок различаем папкой:
    «Ст4 — Вид слева», «Ст4 — Вид справа» (24.10: оба листа были «Ст4»)."""
    stem = lambda f: os.path.splitext(os.path.basename(f))[0]  # noqa: E731
    dirs = {}
    for _, f, _, _ in items:
        dirs.setdefault(stem(f).lower(), set()).add(os.path.normcase(os.path.dirname(os.path.abspath(f))))
    out = []
    for k, f, j, mass in items:
        name = stem(f)
        if len(dirs[name.lower()]) > 1:
            name += " — " + os.path.basename(os.path.dirname(os.path.abspath(f)))
        if j is not None:
            name += f" (лист {j + 1})"
        out.append((k, name, mass))
    return out


def merge(api7, files, target, log=print):
    """→ (листов, ошибки [(файл, причина)]). Общий чертёж сохраняется в target."""
    import frame2cdw as fd
    dst = fd.need(api7.Documents.Add(1, False), "создать чертёж")
    dst2d = CastTo(dst, "IKompasDocument2D")
    dviews = CastTo(dst2d.ViewsAndLayersManager, "IViewsAndLayersManager").Views
    sysview = dviews.View(0)
    sheets_used, origin, errors, names = 0, 0.0, [], []
    try:
        for f in files:
            opened = fd.find_open_doc(api7, f)
            src = opened or api7.Documents.Open(f, False, True)
            if src is None:
                errors.append((f, "КОМПАС не открыл файл (повреждён, другой версии или занят)"))
                continue
            try:
                src2d = CastTo(src, "IKompasDocument2D")
                n_src = src2d.LayoutSheets.Count
                first = sheets_used                        # лист общего чертежа для 1-го листа источника
                for j in range(n_src):
                    if sheets_used + j >= dst2d.LayoutSheets.Count:
                        dst2d.LayoutSheets.Add()
                    copy_format(src2d.LayoutSheets.Item(j), dst2d.LayoutSheets.Item(sheets_used + j))
                # сдвиг: начало первого листа источника в общем чертеже
                dx = sum(sheet_size(dst2d.LayoutSheets.Item(k))[0] for k in range(first))
                sviews = CastTo(src2d.ViewsAndLayersManager, "IViewsAndLayersManager").Views
                for i in range(sviews.Count):
                    v = CastTo(sviews.View(i), "IView")
                    if v.Number == 0:
                        if first == 0:
                            copy_view_objects(src2d, v, dst, dst2d, sysview)
                        else:
                            # системный вид источника → отдельный вид 1:1 в начале его листа
                            nv = CastTo(dviews.Add(vt_Normal), "IView")
                            nv.Scale, nv.X, nv.Y = 1.0, dx, 0.0
                            nv.Name = f"Лист {first + 1}"
                            copy_designation(None, nv)
                            nv.Update()
                            copy_view_objects(src2d, v, dst, dst2d, nv)
                        continue
                    nv = CastTo(dviews.Add(vt_Normal), "IView")
                    nv.Scale, nv.X, nv.Y, nv.Angle = v.Scale, v.X + dx, v.Y, v.Angle
                    try:
                        nv.Name = v.Name
                    except Exception:
                        pass
                    copy_designation(v, nv)
                    nv.Update()
                    copy_view_objects(src2d, v, dst, dst2d, nv)
                for j in range(n_src):
                    n = copy_stamp(src2d.LayoutSheets.Item(j), dst2d.LayoutSheets.Item(sheets_used + j))
                    sst = CastTo(CastTo(src2d.LayoutSheets.Item(j), "ILayoutSheet").Stamp, "IStamp")
                    cell = lambda c: (CastTo(sst.Text(c), "IText").Str if sst.IsCellPresent(c) else "")  # noqa: E731
                    try:
                        mass = cell(5)
                    except Exception:
                        mass = ""
                    # наименование листа — из имени файла (Ст4, Стенка под металл двери…)
                    names.append((sheets_used + j, f, j if n_src > 1 else None, mass))
                    log(f"    штамп листа {sheets_used + j + 1}: ячеек {n}")
                sheets_used += n_src
                log(f"[{sheets_used}] {os.path.basename(f)}")
            except Exception as e:
                traceback.print_exc()
                errors.append((f, str(e)))
            finally:
                if opened is None:
                    src.Close(0)
        if sheets_used == 0:
            raise RuntimeError("ни один чертёж не удалось перенести")
        # лишние пустые листы (если последние источники не открылись)
        while dst2d.LayoutSheets.Count > sheets_used:
            CastTo(dst2d.LayoutSheets.Item(dst2d.LayoutSheets.Count - 1), "ILayoutSheet").Delete()
        names = sheet_names(names)
        # наименование (из имени файла) и масса — в штамп каждого листа. Пишем в самом конце, когда
        # все листы уже созданы и оформлены: иначе у последующих листов запись не держится.
        for k, name, mass in names:
            try:
                st = CastTo(CastTo(dst2d.LayoutSheets.Item(k), "ILayoutSheet").Stamp, "IStamp")
                CastTo(st.Text(1), "IText").Str = name
                if st.IsCellPresent(5):
                    CastTo(st.Text(5), "IText").Str = mass
                st.Update()
            except Exception as e:
                log(f"    штамп листа {k + 1}: {e}")
        # проверка: если КОМПАС всё же сделал наименование общим — подпись на самом листе
        for k, name, mass in names:
            st = CastTo(CastTo(dst2d.LayoutSheets.Item(k), "ILayoutSheet").Stamp, "IStamp")
            if CastTo(st.Text(1), "IText").Str != name:
                x0 = sum(sheet_size(dst2d.LayoutSheets.Item(i))[0] for i in range(k))
                label_sheet(dst2d, dviews, x0, sheet_size(dst2d.LayoutSheets.Item(k))[1],
                            name + (f", масса {mass} кг" if mass else ""))
        os.makedirs(os.path.dirname(os.path.abspath(target)), exist_ok=True)
        CastTo(dst2d, "IKompasDocument").SaveAs(target)
    finally:
        dst.Close(0)
    return sheets_used, errors
