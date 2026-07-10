"""Пост-обработка книги после BookBuilder (tools/pivot) — правки XML прямо в .xlsx.

Требование пользователя (жёсткое): сводная «Сводная_пески» должна СТРОИТЬСЯ ИЗ умной
таблицы «Таблица1» на листе «Сводный_Реестр» (ПОЛНЫЙ реестр, ~43000 строк, ВСЕ
переделы), и при этом ПОКАЗЫВАТЬ ТОЛЬКО «Транспортировка песков» — по подразделениям
(колонки) × Дата.Факт/Время (строки), Σ «Обьем работ, м3».

Почему нельзя просто выставить фильтр + refreshOnLoad="1": при открытии Excel
ПЕРЕСТРАИВАЕТ кэш из источника, ПЕРЕУПОРЯДОЧИВАЕТ sharedItems, и индексный выбор
фильтра «съезжает» → «несколько элементов», неверные итоги (проверено на Excel-for-Mac).

Решение: МАТЕРИАЛИЗУЕМ кэш сводной прямо из данных «Таблица1» и ставим
refreshOnLoad="0" — Excel ДОВЕРЯЕТ готовому кэшу и НЕ перестраивает его, поэтому
индексный фильтр «Передел»=«Транспортировка песков» остаётся валидным.
EPPlus 4.5.3.3 пересчитывать pivotCacheRecords не умеет (нет Refresh/Calculate) —
поэтому весь кэш + записи + сводную формирует этот Python-пост-шаг из листа-реестра.

Делаем:
  1) де-дубликация ТОЛЬКО «Таблица1» на листе-реестре — оставляем ОДНУ, чиним её ref
     под фактический объём, выкидываем лишние table-файлы и их обвязку. Лист
     «Данные_пески» не трогаем;
  2) читаем данные «Таблица1» (openpyxl) и МАТЕРИАЛИЗУЕМ:
       • pivotCacheDefinition1.xml — worksheetSource name="Таблица1",
         refreshOnLoad="0", recordCount, sharedItems для полей осей/фильтра
         (Подразделение, Дата.Факт, Время, Передел);
       • pivotCacheRecords1.xml — по одной <r> на строку реестра (23 значения);
       • pivotTable1.xml — раскладка (Подразделение=колонки, Дата.Факт+Время=строки,
         Σ Обьем=данные), фильтр «Передел» показывает ТОЛЬКО «Транспортировка песков»
         (остальные h="1"), «Дата выдачи» из области страниц убрана (все даты).

Запуск: python3 pivot_postprocess.py <книга.xlsx>
"""
from __future__ import annotations

import datetime as _dt
import re
import shutil
import sys
import zipfile
from pathlib import Path

import openpyxl


def _xa(s) -> str:
    """XML-экранирование значения для атрибута в двойных кавычках (& < > " )."""
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))

REESTR_SHEET = "Сводный_Реестр"       # лист с полным реестром и «Таблица1»
REESTR_TABLE = "Таблица1"             # умная таблица полного реестра (её и дедупим)
PIVOT_SHEET = "Сводная_пески"         # лист со сводной
SAND = "Транспортировка песков"       # единственный передел, который показываем
PERDEL_HDR = "Передел"
POD_HDR = "Подразделение"
DATEFACT_HDR = "Дата. Факт"
DATEORDER_HDR = "Дата выдачи наряд-задания"
TIME_HDR = "Время"
VOL_HDR = "Обьем работ, м3"

_SHEET_RE = re.compile(r"^xl/worksheets/sheet\d+\.xml$")
_PT_RE = re.compile(r"^xl/pivotTables/pivotTable\d+\.xml$")
_CDEF_RE = re.compile(r"pivotCacheDefinition\d+\.xml$")
_CREC_RE = re.compile(r"pivotCacheRecords\d+\.xml$")

NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


# ─────────────────────────── утилиты ────────────────────────────
def _norm(s) -> str:
    """Схлопнуть пробелы (в т.ч. NBSP) и подрезать края — как в BookBuilder."""
    if s is None:
        return ""
    return " ".join(str(s).replace(" ", " ").split())


def _abspart(target: str) -> str:
    """'../tables/table1.xml' → 'xl/tables/table1.xml'."""
    return "xl/" + target.replace("../", "").lstrip("/")


def _table_name(part_txt: str) -> str:
    m = re.search(r'<table\b[^>]*\bname="([^"]+)"', part_txt)
    return m.group(1) if m else ""


# ──────────────── 1) де-дубликация «Таблица1» ────────────────
def _reestr_sheet(files) -> str:
    """Лист-реестр = worksheet, чей <tableParts> ссылается на «Таблица1»."""
    for n in sorted(files):
        if not (_SHEET_RE.match(n) and b"<tableParts" in files[n]):
            continue
        rels = files.get(f"xl/worksheets/_rels/{Path(n).name}.rels", b"").decode("utf-8")
        for _, tgt in re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels):
            if "tables/table" not in tgt:
                continue
            if _table_name(files.get(_abspart(tgt), b"").decode("utf-8")) == REESTR_TABLE:
                return n
    raise RuntimeError(f"не найден лист с умной таблицей «{REESTR_TABLE}»")


def _dedup_reestr_table(files) -> str:
    """Оставить ОДНУ «Таблица1» на листе-реестре, ref = фактический dimension.
    Возвращает новый ref («A1:W…»)."""
    sheet = _reestr_sheet(files)
    sheet_txt = files[sheet].decode("utf-8")

    m = re.search(r'<dimension ref="([A-Z]+)\d+:([A-Z]+)(\d+)"', sheet_txt)
    if not m:
        raise RuntimeError("нет <dimension> у листа-реестра")
    new_ref = f"A1:{m.group(2)}{m.group(3)}"

    rels_name = f"xl/worksheets/_rels/{Path(sheet).name}.rels"
    rels_txt = files.get(rels_name, b"").decode("utf-8")
    rels = re.findall(r'<Relationship\b[^>]*?Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?/>',
                      rels_txt)
    tbl_rels = [(rid, tgt) for rid, tgt in rels if "tables/table" in tgt
                and _table_name(files.get(_abspart(tgt), b"").decode("utf-8")) == REESTR_TABLE]
    if not tbl_rels:
        raise RuntimeError(f"у листа-реестра нет отношений к «{REESTR_TABLE}»")

    keep_rid, keep_tgt = tbl_rels[0]
    keep_part = _abspart(keep_tgt)
    drop = [(rid, _abspart(tgt)) for rid, tgt in tbl_rels[1:]]

    for rid, _ in drop:                       # 1) rels листа: выкинуть дубли
        rels_txt = re.sub(rf'<Relationship\b[^>]*?Id="{re.escape(rid)}"[^>]*?/>', "", rels_txt)
    files[rels_name] = rels_txt.encode("utf-8")

    drop_rids = {rid for rid, _ in drop}      # 2) <tableParts>: убрать дубли, поправить count
    m_tp = re.search(r"<tableParts\b[^>]*>(.*?)</tableParts>", sheet_txt, re.S)
    if m_tp:
        kept = [p for p in re.findall(r'<tablePart\b[^>]*\br:id="([^"]+)"[^>]*/>', m_tp.group(1))
                if p not in drop_rids]
        new_tp = (f'<tableParts count="{len(kept)}">'
                  + "".join(f'<tablePart r:id="{rid}"/>' for rid in kept) + "</tableParts>")
        sheet_txt = sheet_txt[:m_tp.start()] + new_tp + sheet_txt[m_tp.end():]
    files[sheet] = sheet_txt.encode("utf-8")

    keep_txt = files[keep_part].decode("utf-8")   # 3) ref оставленной таблицы
    keep_txt = re.sub(r'(<table\b[^>]*\bref=")[^"]*(")', rf"\g<1>{new_ref}\g<2>", keep_txt, count=1)
    keep_txt = re.sub(r'(<autoFilter\b[^>]*\bref=")[^"]*(")', rf"\g<1>{new_ref}\g<2>", keep_txt, count=1)
    files[keep_part] = keep_txt.encode("utf-8")
    for _, part in drop:
        files.pop(part, None)

    ct = files["[Content_Types].xml"].decode("utf-8")   # 4) Content_Types
    for _, part in drop:
        ct = re.sub(rf'<Override\b[^>]*?PartName="/{re.escape(part)}"[^>]*?/>', "", ct)
    files["[Content_Types].xml"] = ct.encode("utf-8")

    return new_ref


# ──────────────── 2) чтение данных «Таблица1» ────────────────
def _read_registry(xlsx: Path):
    """Читает лист «Сводный_Реестр» → (headers, rows). rows — список кортежей длиной
    len(headers). Значения: datetime | int | float | str | None."""
    wb = openpyxl.load_workbook(xlsx, data_only=True, read_only=True)
    try:
        ws = wb[REESTR_SHEET]
        it = ws.iter_rows(values_only=True)
        headers = [_norm(h) for h in next(it)]
        n = len(headers)
        rows = []
        for r in it:
            r = tuple(r)
            r = r[:n] if len(r) >= n else r + (None,) * (n - len(r))  # ровно n колонок
            # полностью пустые строки пропускаем (openpyxl иногда тянет хвост)
            if all(v is None or (isinstance(v, str) and v == "") for v in r):
                continue
            rows.append(r)
        return headers, rows
    finally:
        wb.close()


# ──────────────── OOXML-сериализация значений ────────────────
def _iso(v) -> str:
    """datetime/date → ISO без микросекунд (формат sharedItems <d v=.../> Excel)."""
    if isinstance(v, _dt.datetime):
        d = v
    else:  # date
        d = _dt.datetime(v.year, v.month, v.day)
    return d.strftime("%Y-%m-%dT%H:%M:%S")


def _is_date(v) -> bool:
    return isinstance(v, (_dt.datetime, _dt.date)) and not isinstance(v, bool)


def _num_str(v) -> str:
    """Число → строка для <n v=..>. int без .0, float минимально."""
    if isinstance(v, bool):
        return "1" if v else "0"
    if isinstance(v, int):
        return str(v)
    f = float(v)
    if f.is_integer():
        return str(int(f))
    return repr(f)


# ──────────────── 3) материализация сводной ────────────────
def _build_pivot(files, headers, rows) -> dict:
    """Собрать pivotCacheDefinition / pivotCacheRecords / pivotTable из данных реестра.
    Пишет XML в files. Возвращает статистику для лога."""
    idx = {h: i for i, h in enumerate(headers)}
    for req in (POD_HDR, DATEFACT_HDR, TIME_HDR, VOL_HDR, PERDEL_HDR):
        if req not in idx:
            raise RuntimeError(f"в реестре нет колонки «{req}»")
    ncols = len(headers)
    i_pod, i_df, i_tm, i_per = idx[POD_HDR], idx[DATEFACT_HDR], idx[TIME_HDR], idx[PERDEL_HDR]

    # поля, для которых нужны sharedItems (оси/фильтр). Порядок значений = порядок
    # первого появления в данных (Excel так и делает при построении кэша).
    SHARED = {i_pod, i_df, i_tm, i_per}
    shared_vals: dict[int, list] = {c: [] for c in SHARED}       # список значений (в исходном типе)
    shared_index: dict[int, dict] = {c: {} for c in SHARED}      # ключ→индекс
    shared_has_blank = {c: False for c in SHARED}

    def _key(c, v):
        # ключ для дедупа sharedItems: даты по ISO, текст как есть (без norm — важно
        # для «Передел», где нам нужно сохранить исходные варианты пробелов).
        if _is_date(v):
            return ("d", _iso(v))
        return ("s", "" if v is None else str(v))

    # ── проход по строкам: наполняем sharedItems, строим записи И параллельно
    #    накапливаем ПЕСКОВУЮ раскладку (какие пары дата×время и какие подразделения
    #    реально имеют «Транспортировка песков»), чтобы тело сводной было
    #    корректным-по-построению при refreshOnLoad=0. ── ── ── ── ── ── ── ── ──
    rec_parts: list[str] = []
    # индексы кэша для сандовой раскладки (заполняются по ходу — sharedItems растут)
    sand_dt: dict[int, set] = {}   # dateIdx -> {timeIdx,...} (только пески)
    sand_pod: set = set()          # {podIdx,...} (только пески)
    for r in rows:
        cells = []
        row_idx = [None] * ncols   # индексы sharedItems этой строки (для SHARED-полей)
        for c in range(ncols):
            v = r[c] if c < len(r) else None
            blank = v is None or (isinstance(v, str) and v == "")
            if c in SHARED:
                if blank:
                    shared_has_blank[c] = True
                    cells.append("<m/>")
                    continue
                k = _key(c, v)
                si = shared_index[c].get(k)
                if si is None:
                    si = len(shared_vals[c])
                    shared_index[c][k] = si
                    shared_vals[c].append(v)
                row_idx[c] = si
                cells.append(f'<x v="{si}"/>')
            else:
                if blank:
                    cells.append("<m/>")
                elif _is_date(v):
                    cells.append(f'<d v="{_iso(v)}"/>')
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    cells.append(f'<n v="{_num_str(v)}"/>')
                else:
                    cells.append(f'<s v="{_xa(str(v))}"/>')
        rec_parts.append("<r>" + "".join(cells) + "</r>")
        # песковая строка? (передел нормализуется в SAND) — копим раскладку
        pv = r[i_per] if i_per < len(r) else None
        if pv is not None and _norm(pv) == SAND:
            di, ti, pi = row_idx[i_df], row_idx[i_tm], row_idx[i_pod]
            if di is not None and ti is not None:
                sand_dt.setdefault(di, set()).add(ti)
            if pi is not None:
                sand_pod.add(pi)

    n_rec = len(rec_parts)

    # ── sharedItems XML для полей осей/фильтра ──
    def _shared_items_xml(c) -> str:
        vals = shared_vals[c]
        is_date_field = any(_is_date(v) for v in vals)
        items = []
        for v in vals:
            if _is_date(v):
                items.append(f'<d v="{_iso(v)}"/>')
            else:
                items.append(f'<s v="{_xa(str(v))}"/>')
        count = len(vals) + (1 if shared_has_blank[c] else 0)
        blank_attr = ' containsBlank="1"' if shared_has_blank[c] else ""
        if is_date_field:
            iso_dates = [_iso(v) for v in vals if _is_date(v)]
            mn = min(iso_dates)
            mx = max(iso_dates)
            attrs = (f'{blank_attr} containsNonDate="0" containsDate="1" '
                     f'containsString="0" minDate="{mn}" maxDate="{mx}" count="{count}"')
        else:
            attrs = f'{blank_attr} count="{count}"'
        body = "".join(items) + ("<m/>" if shared_has_blank[c] else "")
        return f"<sharedItems{attrs}>{body}</sharedItems>"

    # ── cacheFields (все 23, порядок = колонки) ──
    cf_parts = []
    for c in range(ncols):
        name = _xa(headers[c])
        if c in SHARED:
            cf_parts.append(f'<cacheField name="{name}" numFmtId="0">{_shared_items_xml(c)}</cacheField>')
        else:
            # для не-осевых полей sharedItems не обязательны; ставим пустой контейнер
            # с признаком blank (Excel это допускает, записи хранят значения inline).
            cf_parts.append(f'<cacheField name="{name}" numFmtId="0"><sharedItems containsBlank="1"/></cacheField>')

    cdef = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<pivotCacheDefinition xmlns="{NS_MAIN}" xmlns:r="{NS_R}" '
        'r:id="rId1" refreshOnLoad="0" refreshedBy="hronos" '
        'createdVersion="4" refreshedVersion="4" minRefreshableVersion="3" '
        f'recordCount="{n_rec}">'
        '<cacheSource type="worksheet"><worksheetSource name="Таблица1"/></cacheSource>'
        f'<cacheFields count="{ncols}">' + "".join(cf_parts) + "</cacheFields>"
        "</pivotCacheDefinition>"
    )

    crec = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<pivotCacheRecords xmlns="{NS_MAIN}" xmlns:r="{NS_R}" count="{n_rec}">'
        + "".join(rec_parts) + "</pivotCacheRecords>"
    )

    # ── имена частей (как в книге) ──
    cdef_part = next((n for n in files if _CDEF_RE.search(n)), "xl/pivotCache/pivotCacheDefinition1.xml")
    crec_part = next((n for n in files if _CREC_RE.search(n)), "xl/pivotCache/pivotCacheRecords1.xml")
    files[cdef_part] = cdef.encode("utf-8")
    files[crec_part] = crec.encode("utf-8")

    # ── pivotTable1.xml ──
    pod_items = shared_vals[i_pod]         # индексы колонок
    df_items = shared_vals[i_df]
    tm_items = shared_vals[i_tm]
    per_items = shared_vals[i_per]

    # индекс(ы) «Транспортировка песков» (по нормализации) среди Передел sharedItems
    sand_idx = {si for si, v in enumerate(per_items) if _norm(v) == SAND}
    if not sand_idx:
        raise RuntimeError(f"в данных нет передела «{SAND}» — фильтр не построить")

    pt_part = next((n for n in files if _PT_RE.match(n)), "xl/pivotTables/pivotTable1.xml")
    files[pt_part] = _build_pivottable(
        ncols, i_pod, i_df, i_tm, i_per, idx[VOL_HDR],
        pod_items, df_items, tm_items, per_items,
        shared_has_blank, sand_idx, sand_dt, sand_pod,
    ).encode("utf-8")

    return {
        "records": n_rec,
        "podr": len(pod_items),
        "dates": len(df_items),
        "times": len(tm_items),
        "peredel": len(per_items),
        "sand_items": sorted(sand_idx),
        "cdef": cdef_part,
        "crec": crec_part,
        "pt": pt_part,
    }


def _build_pivottable(ncols, i_pod, i_df, i_tm, i_per, i_vol,
                      pod_items, df_items, tm_items, per_items,
                      has_blank, sand_idx, sand_dt, sand_pod) -> str:
    """Собрать pivotTable1.xml: Подразделение=колонки, Дата.Факт+Время=строки,
    Σ Обьем=данные, фильтр «Передел»=только пески.

    Тело (rowItems/colItems/location) строим КОРРЕКТНО-ПО-ПОСТРОЕНИЮ из уже
    посчитанной песковой раскладки (sand_dt: dateIdx→{timeIdx}, sand_pod: {podIdx}),
    в точности как это делает Excel: перечисляем только НЕПУСТЫЕ комбинации
    (даты по возрастанию, под каждой — её времена по возрастанию), затем grand-строка.
    Значения ячеек в pivotTable НЕ хранятся — их Excel считает из кэша в ячейки листа;
    от нас требуется лишь верная структура. refreshOnLoad=0 → Excel доверяет и не
    перестраивает кэш, поэтому индексный фильтр «Передел» не «съезжает»."""
    n_pod = len(pod_items)
    n_df = len(df_items)
    n_tm = len(tm_items)
    n_per = len(per_items)

    def _items(count, has_bl, sel_hidden=None):
        """Список <item x=../> в порядке индексов кэша + (blank) + t=default.
        sel_hidden: множество индексов, которые h='1' (скрыть). None → все видимы."""
        parts = []
        total = count + (1 if has_bl else 0)
        for i in range(total):
            if sel_hidden is not None and i in sel_hidden:
                parts.append(f'<item h="1" x="{i}"/>')
            else:
                parts.append(f'<item x="{i}"/>')
        parts.append('<item t="default"/>')
        return f'<items count="{total + 1}">' + "".join(parts) + "</items>"

    # «Передел»: показать ТОЛЬКО индексы sand_idx, остальные h="1"
    per_total = n_per + (1 if has_blank[i_per] else 0)
    per_hidden = {i for i in range(per_total) if i not in sand_idx}

    # pivotFields (23). Оси/фильтр — с <items>, прочие — заглушки.
    pf = []
    for c in range(ncols):
        if c == i_pod:
            pf.append(f'<pivotField axis="axisCol" showAll="0">{_items(n_pod, has_blank[i_pod])}</pivotField>')
        elif c == i_df:
            pf.append(f'<pivotField axis="axisRow" showAll="0">{_items(n_df, has_blank[i_df])}</pivotField>')
        elif c == i_tm:
            pf.append(f'<pivotField axis="axisRow" showAll="0">{_items(n_tm, has_blank[i_tm])}</pivotField>')
        elif c == i_per:
            pf.append(f'<pivotField axis="axisPage" multipleItemSelectionAllowed="1" showAll="0">'
                      f'{_items(n_per, has_blank[i_per], per_hidden)}</pivotField>')
        elif c == i_vol:
            pf.append('<pivotField dataField="1" showAll="0"/>')
        else:
            pf.append('<pivotField showAll="0"/>')
    pivot_fields = f'<pivotFields count="{ncols}">' + "".join(pf) + "</pivotFields>"

    # ── rowItems: только НЕПУСТЫЕ (дата → её времена), как строит Excel ──
    dates_sorted = sorted(sand_dt)                # по возрастанию индекса кэша
    ri = []
    for di in dates_sorted:
        # строка-заголовок даты (уровень 0). Первый <x> без v, если индекс совпал бы с 0
        ri.append(f"<i>{_xitem(di)}</i>")
        for ti in sorted(sand_dt[di]):            # времена под датой (уровень 1)
            ri.append(f'<i r="1">{_xitem(ti)}</i>')
    ri.append('<i t="grand"><x/></i>')
    row_items = f'<rowItems count="{len(ri)}">' + "".join(ri) + "</rowItems>"

    # ── colItems: подразделения с песками + grand ──
    pods_sorted = sorted(sand_pod)
    ci = [f"<i>{_xitem(pi)}</i>" for pi in pods_sorted]
    ci.append('<i t="grand"><x/></i>')
    col_items = f'<colItems count="{len(ci)}">' + "".join(ci) + "</colItems>"

    row_fields = f'<rowFields count="2"><field x="{i_df}"/><field x="{i_tm}"/></rowFields>'
    col_fields = f'<colFields count="1"><field x="{i_pod}"/></colFields>'
    page_fields = f'<pageFields count="1"><pageField fld="{i_per}" hier="-1"/></pageFields>'
    data_fields = (f'<dataFields count="1"><dataField name="Сумма по полю {_xa(VOL_HDR)}" '
                   f'fld="{i_vol}" baseField="0" baseItem="0"/></dataFields>')

    # ── location: A4 = страница-фильтр, A6 = заголовки колонок, A7 = первые данные.
    #    Ширина: 1 (подписи строк) + число колонок-подразделений + grand.
    #    Высота: firstHeaderRow(=строка колонок) + строки rowItems.
    n_body_cols = len(ci)                          # подразделения + grand
    n_body_rows = len(ri)                          # строки данных + grand
    last_col = _col_letter(1 + n_body_cols)
    # A4 старт как в шаблоне; firstDataRow смещаем на 2 (шапка колонок + строка полей)
    loc = (f'<location ref="A4:{last_col}{5 + n_body_rows}" '
           'firstHeaderRow="1" firstDataRow="2" firstDataCol="1" '
           'rowPageCount="1" colPageCount="1"/>')

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
        f'<pivotTableDefinition xmlns="{NS_MAIN}" name="Сводная таблица1" cacheId="78" '
        'applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" '
        'applyPatternFormats="0" applyAlignmentFormats="0" applyWidthHeightFormats="1" '
        'dataCaption="Значения" updatedVersion="4" minRefreshableVersion="3" '
        'itemPrintTitles="1" createdVersion="4" indent="0" outline="1" outlineData="1" '
        'multipleFieldFilters="0">'
        + loc + pivot_fields + row_fields + row_items + col_fields + col_items
        + page_fields + data_fields +
        '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" '
        'showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/>'
        "</pivotTableDefinition>"
    )


def _xitem(idx: int) -> str:
    """<x/> для индекса 0 (Excel опускает v="0"), иначе <x v="idx"/>."""
    return "<x/>" if idx == 0 else f'<x v="{idx}"/>'


def _col_letter(n: int) -> str:
    """1→A, 26→Z, 27→AA…"""
    s = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def _clear_pivot_sheet(files) -> None:
    """Очистить лист со сводной от УСТАРЕВШИХ ячеек тела (значения авторского шаблона
    A1:G41 и leftover-ячейка страничного поля «Дата выдачи»). Иначе при refreshOnLoad=0
    потребитель мог бы показать старые числа. После очистки любой движок (Excel,
    LibreOffice) перерисует сводную из МАТЕРИАЛИЗОВАННОГО кэша.
    Находим лист по rel-ссылке на pivotTable; чистим <sheetData> и <dimension>."""
    pt_part = next((n for n in files if _PT_RE.match(n)), None)
    if not pt_part:
        return
    pt_file = Path(pt_part).name                      # pivotTable1.xml
    for n in list(files):
        if not _SHEET_RE.match(n):
            continue
        rels = files.get(f"xl/worksheets/_rels/{Path(n).name}.rels", b"").decode("utf-8")
        if pt_file not in rels:
            continue
        txt = files[n].decode("utf-8")
        txt = re.sub(r"<sheetData>.*?</sheetData>", "<sheetData/>", txt, count=1, flags=re.S)
        txt = re.sub(r'<dimension ref="[^"]*"\s*/>', '<dimension ref="A1"/>', txt, count=1)
        files[n] = txt.encode("utf-8")
        return


# ──────────────── оркестрация ────────────────
def postprocess(xlsx: Path) -> str:
    xlsx = Path(xlsx)
    with zipfile.ZipFile(xlsx) as z:
        files = {i.filename: z.read(i.filename) for i in z.infolist()}

    new_ref = _dedup_reestr_table(files)                 # 1) одна «Таблица1», верный ref
    headers, rows = _read_registry(xlsx)                 # 2) данные реестра
    stats = _build_pivot(files, headers, rows)           # 3) материализовать сводную
    _clear_pivot_sheet(files)                            # 4) снять устаревшее тело сводной
    print(f"[pivot] кэш материализован: {stats['records']} записей, "
          f"подразделений {stats['podr']}, дат {stats['dates']}, времён {stats['times']}, "
          f"переделов {stats['peredel']}; фильтр «Передел»→пески idx {stats['sand_items']}")

    tmp = xlsx.with_name(xlsx.stem + ".pp.tmp.xlsx")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in files.items():
            zout.writestr(name, data)
    shutil.move(str(tmp), str(xlsx))
    return new_ref


def main() -> int:
    if len(sys.argv) < 2:
        print("usage: python3 pivot_postprocess.py <книга.xlsx>")
        return 2
    ref = postprocess(Path(sys.argv[1]))
    print(f"[OK] «{REESTR_TABLE}» → {ref}, сводная материализована из «Таблица1», "
          f"refreshOnLoad=0, дубль-таблицы вычищены")
    return 0


if __name__ == "__main__":
    sys.exit(main())
