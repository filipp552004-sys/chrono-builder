"""Генератор клиентского Excel «Баланс рабочего времени» (ФРД / самофотография). v3.

Методика 2.2: одиннадцать индексов затрат рабочего времени, причины потерь,
фактический и проектируемый баланс, коэффициенты, справочные надбавки.

Вход (payload из process-frd v2):
  position, department, industry, supervisor, period, schedule_type,
  work_start, work_end, lunch_minutes, days, source ('self'|'delegated'),
  records: [{day_number, obs_date, start, end, duration, activity,
             wt_index, wt_label, loss_reason, needs_review}],
  balance: {byIndex, projected, total, perDay, operative, servicing, normirovannoe,
            losses, lossesByReason, kisp, kpnt, kpnd, knr, checksum, ppt, pptFormula,
            dKisp, reservePerDay, aob, aotl, tpzPerDay, warnings}

Факт считается живыми формулами от наблюдательного листа (SUMIF по скрытой
колонке индекса) — клиент может править индекс, баланс пересчитается.
Проект берётся из payload как значения: продолжительность дня не меняется,
потери переходят в оперативное основное время (ОО). Считать проект как
«всего минус потери» НЕЛЬЗЯ — это ошибка v2.
"""
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.workbook.properties import CalcProperties

GREEN = "FF2E6E4E"; RESULT = "FFDDF0E4"; SUB = "FFEAF2ED"; WHITE = "FFFFFFFF"; YELLOW = "FFFFF2CC"
RED = "FFE24B4A"; AMBER = "FFEF9F27"; BLUE = "FF378ADD"; GREY = "FF7A7A7A"
thin = Side(style="thin", color="FFBFBFBF"); BORD = Border(left=thin, right=thin, top=thin, bottom=thin)

INDEXES = ["PZ", "OO", "OV", "OBT", "OBO", "OTL", "PT", "NR", "SR", "PNT", "PND"]
IDX_RU = {"PZ": "ПЗ", "OO": "ОО", "OV": "ОВ", "OBT": "ОБТ", "OBO": "ОБО", "OTL": "ОТЛ", "PT": "ПТ",
          "NR": "НР", "SR": "СР", "PNT": "ПНТ", "PND": "ПНД"}
IDX_LABEL = {
    "PZ": "Подготовительно-заключительное", "OO": "Оперативное основное", "OV": "Оперативное вспомогательное",
    "OBT": "Обслуживание техническое", "OBO": "Обслуживание организационное",
    "OTL": "Отдых и личные надобности", "PT": "Перерывы по технологии",
    "NR": "Непроизводительная работа", "SR": "Случайная работа",
    "PNT": "Потери организационно-технические", "PND": "Потери по вине работника",
}
IDX_GROUP = {"PZ": "Работа", "OO": "Работа", "OV": "Работа", "OBT": "Работа", "OBO": "Работа",
             "OTL": "Регламентированные перерывы", "PT": "Регламентированные перерывы",
             "NR": "Потери", "SR": "Потери", "PNT": "Потери", "PND": "Потери"}
IDX_COLOR = {"PZ": GREY, "OO": GREEN, "OV": "FF4B9B6F", "OBT": "FF5A7090", "OBO": "FF5A7090",
             "OTL": BLUE, "PT": BLUE, "NR": AMBER, "SR": AMBER, "PNT": RED, "PND": RED}
LOSS = ["NR", "SR", "PNT", "PND"]
NORM = ["PZ", "OO", "OV", "OBT", "OBO", "OTL", "PT"]
REASON_RU = {"ORG": "Организационные", "TECH": "Технические", "INFO": "Информационные",
             "DISC": "Дисциплинарные", "EXT": "Внешние"}
LST = "Наблюдательный лист"
BAL = "Баланс"


def bar(ws, row, text, fill=GREEN, size=11, span=6):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    c = ws.cell(row, 1, text); c.font = Font(bold=True, color=WHITE, size=size)
    c.fill = PatternFill("solid", fgColor=fill); c.alignment = Alignment(vertical="center"); ws.row_dimensions[row].height = 19


def kv(ws, row, label, value, fill=None, bold=False, note="", fmt=None):
    b = ws.cell(row, 2, label); b.font = Font(size=10, bold=bold)
    v = ws.cell(row, 3, value); v.font = Font(size=10, bold=bold); v.alignment = Alignment(horizontal="center")
    if fmt: v.number_format = fmt
    if fill:
        f = PatternFill("solid", fgColor=fill); b.fill = f; v.fill = f
    if note:
        n = ws.cell(row, 4, note); n.font = Font(size=9, italic=True, color=GREY)
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=6)
    return v


def hdr(ws, row, cols, texts):
    for c, t in zip(cols, texts):
        cc = ws.cell(row, c, t); cc.font = Font(bold=True, size=10); cc.fill = PatternFill("solid", fgColor=SUB)
        cc.border = BORD; cc.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _f(x):
    try: return float(x or 0)
    except (TypeError, ValueError): return 0.0


def build_frd(d):
    recs = d.get("records", []) or []
    bal = d.get("balance", {}) or {}
    by = bal.get("byIndex", {}) or {}
    proj = bal.get("projected", {}) or {}
    days = max(1, int(_f(d.get("days", 1)) or 1))
    delegated = d.get("source") == "delegated"

    wb = openpyxl.Workbook()
    ws = wb.active; ws.title = BAL
    lst = wb.create_sheet(LST)
    ind = wb.create_sheet("Показатели")
    rsn = wb.create_sheet("Потери по причинам")
    for s in (ws, lst, ind, rsn):
        s.calculation = CalcProperties(fullCalcOnLoad=True); s.sheet_view.showGridLines = False

    # ── НАБЛЮДАТЕЛЬНЫЙ ЛИСТ ──────────────────────────────────────────
    lst.column_dimensions["A"].width = 5
    for col, w in zip("BCDEFGHIJ", [7, 12, 9, 9, 10, 46, 30, 22, 11]): lst.column_dimensions[col].width = w
    lst.cell(1, 1, "НАБЛЮДАТЕЛЬНЫЙ ЛИСТ · самофотография рабочего дня"); lst.merge_cells("A1:J1")
    lst.cell(1, 1).font = Font(bold=True, color=WHITE, size=12); lst.cell(1, 1).fill = PatternFill("solid", fgColor=GREEN); lst.row_dimensions[1].height = 22
    lst.cell(2, 1, f"Дней наблюдения: {days} · Записей: {len(recs)} · Индекс можно исправить в колонке K (латиницей: PZ, OO, OV, OBT, OBO, OTL, PT, NR, SR, PNT, PND) — баланс пересчитается").font = Font(size=9, italic=True, color=GREY)
    lst.merge_cells("A2:J2")
    hdr(lst, 3, range(1, 11), ["№", "День", "Дата", "Начало", "Оконч.", "Мин", "Действие", "Индекс затрат", "Причина потери", "Проверить"])
    lst.cell(3, 11, "Код"); lst.cell(3, 11).font = Font(bold=True, size=9, color=GREY)
    r = 4
    for i, rec in enumerate(recs, 1):
        idx = str(rec.get("wt_index", "")).strip().upper()
        reason = str(rec.get("loss_reason") or "").strip().upper()
        lst.cell(r, 1, i); lst.cell(r, 2, int(_f(rec.get("day_number", 1)) or 1)); lst.cell(r, 3, rec.get("obs_date", "") or "")
        lst.cell(r, 4, rec.get("start", "")); lst.cell(r, 5, rec.get("end", ""))
        lst.cell(r, 6, round(_f(rec.get("duration", 0)), 1)); lst.cell(r, 7, rec.get("activity", ""))
        e = lst.cell(r, 8, f"{IDX_RU.get(idx, idx)} · {IDX_LABEL.get(idx, '')}".strip(" ·")); e.font = Font(size=9, bold=True, color=IDX_COLOR.get(idx, "FF000000"))
        lst.cell(r, 9, REASON_RU.get(reason, "") if idx in LOSS else "").font = Font(size=9)
        lst.cell(r, 10, "да" if rec.get("needs_review") else "").font = Font(size=9, color=AMBER, bold=True)
        lst.cell(r, 11, idx).font = Font(size=9, color=GREY)
        for c in (1, 2, 3, 4, 5, 6): lst.cell(r, c).alignment = Alignment(horizontal="center")
        for c in range(1, 11):
            lst.cell(r, c).border = BORD
            if c not in (8, 9, 10): lst.cell(r, c).font = Font(size=10)
        r += 1
    last = max(r - 1, 4); first = 4
    dur = f"'{LST}'!$F${first}:$F${last}"; idxc = f"'{LST}'!$K${first}:$K${last}"
    lst.freeze_panes = "A4"

    # ── БАЛАНС ───────────────────────────────────────────────────────
    ws.column_dimensions["A"].width = 3
    for col, w in zip("BCDEFG", [42, 12, 10, 12, 10, 26]): ws.column_dimensions[col].width = w
    ws.cell(1, 1, "БАЛАНС РАБОЧЕГО ВРЕМЕНИ · фотография рабочего дня (самофотография)"); ws.merge_cells("A1:G1")
    ws.cell(1, 1).font = Font(bold=True, color=WHITE, size=13); ws.cell(1, 1).fill = PatternFill("solid", fgColor=GREEN); ws.row_dimensions[1].height = 26
    r = 3; bar(ws, r, "ОБЪЕКТ НАБЛЮДЕНИЯ", span=7)
    r += 1; kv(ws, r, "Должность", d.get("position", ""))
    r += 1; kv(ws, r, "Структурное подразделение", d.get("department", ""))
    r += 1; kv(ws, r, "Отрасль", d.get("industry", ""))
    r += 1; kv(ws, r, "Руководитель", d.get("supervisor", ""))
    r += 1; kv(ws, r, "Период наблюдения", d.get("period", ""))
    r += 1; kv(ws, r, "Режим", f"{d.get('schedule_type', 'рабочий день')} {d.get('work_start', '')} – {d.get('work_end', '')}, обед {int(_f(d.get('lunch_minutes', 0)))} мин")
    r += 1; kv(ws, r, "Метод", "самофотография — записи вёл работник" + (" по поручению заказчика" if delegated else ""))
    r += 1; days_row = r; kv(ws, r, "Дней наблюдения", days, fill=SUB)
    r += 1; tot_row = r; kv(ws, r, "Наблюдаемое время всего, мин", f"=SUM({dur})", bold=True, fill=SUB, fmt="0")
    TOT = f"$C${tot_row}"; DAYS = f"$C${days_row}"
    r += 1; kv(ws, r, "В среднем в день, мин", f"=IF({DAYS}=0,0,{TOT}/{DAYS})", fill=SUB, fmt="0.0")

    r += 2; bar(ws, r, "БАЛАНС ПО ИНДЕКСАМ ЗАТРАТ · факт и проект", span=7)
    r += 1; hdr(ws, r, (2, 3, 4, 5, 6, 7), ("Индекс затрат рабочего времени", "Факт, мин", "Факт, %", "Проект, мин", "Проект, %", "Группа"))
    idx_rows = {}
    for code in INDEXES:
        r += 1; idx_rows[code] = r
        ws.cell(r, 2, f"{IDX_RU[code]} · {IDX_LABEL[code]}").font = Font(size=10, bold=True, color=IDX_COLOR[code])
        ws.cell(r, 3, f'=SUMIF({idxc},"{code}",{dur})').number_format = "0"
        ws.cell(r, 4, f"=IF({TOT}=0,0,C{r}/{TOT})").number_format = "0.0%"
        ws.cell(r, 5, round(_f(proj.get(code, 0)), 1)).number_format = "0"
        ws.cell(r, 6, f"=IF({TOT}=0,0,E{r}/{TOT})").number_format = "0.0%"
        ws.cell(r, 7, IDX_GROUP[code]).font = Font(size=9, color=GREY)
        for c in (3, 4, 5, 6): ws.cell(r, c).alignment = Alignment(horizontal="center")
        for c in range(2, 8): ws.cell(r, c).border = BORD
    fr, lr = idx_rows["PZ"], idx_rows["PND"]
    r += 1; ws.cell(r, 2, "ИТОГО").font = Font(bold=True, size=10)
    ws.cell(r, 3, f"=SUM(C{fr}:C{lr})").number_format = "0"; ws.cell(r, 4, f"=SUM(D{fr}:D{lr})").number_format = "0.0%"
    ws.cell(r, 5, f"=SUM(E{fr}:E{lr})").number_format = "0"; ws.cell(r, 6, f"=SUM(F{fr}:F{lr})").number_format = "0.0%"
    for c in range(2, 8): ws.cell(r, c).border = BORD; ws.cell(r, c).fill = PatternFill("solid", fgColor=SUB); ws.cell(r, c).font = Font(bold=True, size=10)
    for c in (3, 4, 5, 6): ws.cell(r, c).alignment = Alignment(horizontal="center")
    total_row = r
    r += 1; ws.cell(r, 2, "Проект: продолжительность дня не меняется, потери (НР, СР, ПНТ, ПНД) переходят в оперативное основное время (ОО).").font = Font(size=9, italic=True, color=GREY)
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)

    def C(code): return f"C{idx_rows[code]}"
    OP = f"({C('OO')}+{C('OV')})"; SERV = f"({C('OBT')}+{C('OBO')})"
    LOSSES = "(" + "+".join(C(c) for c in LOSS) + ")"; NORMS = "(" + "+".join(C(c) for c in NORM) + ")"

    r += 2; bar(ws, r, "АГРЕГАТЫ", span=7)
    r += 1; op_row = r; kv(ws, r, "Топ — оперативное время (ОО + ОВ), мин", f"={OP}", fmt="0")
    r += 1; kv(ws, r, "Тоб — обслуживание рабочего места (ОБТ + ОБО), мин", f"={SERV}", fmt="0")
    r += 1; norm_row = r; kv(ws, r, "Тнорм — нормируемые затраты (ПЗ+ОО+ОВ+ОБТ+ОБО+ОТЛ+ПТ), мин", f"={NORMS}", fmt="0")
    r += 1; loss_row = r; kv(ws, r, "Тпот — потери (НР+СР+ПНТ+ПНД), мин", f"={LOSSES}", fill=YELLOW, fmt="0")
    OPc, NORMc, LOSSc = f"$C${op_row}", f"$C${norm_row}", f"$C${loss_row}"

    # ── ПОКАЗАТЕЛИ ───────────────────────────────────────────────────
    ind.column_dimensions["A"].width = 3
    for col, w in zip("BCDEF", [52, 12, 14, 14, 30]): ind.column_dimensions[col].width = w
    ind.cell(1, 1, "ПОКАЗАТЕЛИ ИСПОЛЬЗОВАНИЯ РАБОЧЕГО ВРЕМЕНИ"); ind.merge_cells("A1:F1")
    ind.cell(1, 1).font = Font(bold=True, color=WHITE, size=13); ind.cell(1, 1).fill = PatternFill("solid", fgColor=GREEN); ind.row_dimensions[1].height = 26
    B = f"'{BAL}'!"
    r = 3; bar(ind, r, "КОЭФФИЦИЕНТЫ (доли наблюдаемого времени)", span=6)
    r += 1; kisp_row = r; kv(ind, r, "Кисп — коэффициент использования рабочего времени", f"=IF({B}{TOT}=0,0,{B}{NORMc}/{B}{TOT})", bold=True, fill=RESULT, fmt="0.0%", note="Тнорм / Тнабл")
    r += 1; kv(ind, r, "Кпнт — потери организационно-технические", f"=IF({B}{TOT}=0,0,{B}{C('PNT')}/{B}{TOT})", fmt="0.0%")
    r += 1; kv(ind, r, "Кпнд — потери по вине работника", f"=IF({B}{TOT}=0,0,{B}{C('PND')}/{B}{TOT})", fmt="0.0%")
    r += 1; kv(ind, r, "Кнр — непроизводительная и случайная работа", f"=IF({B}{TOT}=0,0,({B}{C('NR')}+{B}{C('SR')})/{B}{TOT})", fmt="0.0%")
    r += 1; kv(ind, r, "Контрольная сумма (должна быть 100 %)", f"=SUM(C{kisp_row}:C{r-1})", fill=SUB, fmt="0.0%", note="Кисп + Кпнт + Кпнд + Кнр")

    r += 2; bar(ind, r, "РЕЗЕРВ РОСТА ПРОИЗВОДИТЕЛЬНОСТИ", span=6)
    r += 1; kv(ind, r, "ППТ — возможный рост производительности труда", f"=IF({B}{OPc}=0,0,{B}{LOSSc}/{B}{OPc})", bold=True, fill=RESULT, fmt="0.0%", note="(Тпнт + Тпнд + Тнр + Тср) / Топ")
    r += 1; ind.cell(r, 2, bal.get("pptFormula", "")).font = Font(size=9, italic=True, color=GREY); ind.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
    r += 1; kv(ind, r, "ΔКисп — прирост коэффициента использования (справочно)", f"=IF(C{kisp_row}=0,0,(1-C{kisp_row})/C{kisp_row})", fmt="0.0%", note="(1 − Кисп) / Кисп")
    r += 1; kv(ind, r, "Резерв рабочего времени, мин/день", f"=IF({B}{DAYS}=0,0,{B}{LOSSc}/{B}{DAYS})", bold=True, fill=RESULT, fmt="0.0")

    r += 2; bar(ind, r, "СПРАВОЧНЫЕ ВЕЛИЧИНЫ ДЛЯ НОРМИРОВАНИЯ (по данным наблюдения)", span=6)
    r += 1; kv(ind, r, "аоб — обслуживание рабочего места, % от Топ", f"=IF({B}{OPc}=0,0,({B}{C('OBT')}+{B}{C('OBO')})/{B}{OPc})", fmt="0.0%")
    r += 1; kv(ind, r, "аотл — отдых и личные надобности, % от Топ", f"=IF({B}{OPc}=0,0,{B}{C('OTL')}/{B}{OPc})", fmt="0.0%")
    r += 1; kv(ind, r, "Тпз — подготовительно-заключительное, мин/день", f"=IF({B}{DAYS}=0,0,{B}{C('PZ')}/{B}{DAYS})", fmt="0.0")
    r += 1; ind.cell(r, 2, "Фактические величины по самофотографии. Норматив ОТЛ для условий труда не применялся; превышение фактического ОТЛ над нормативом работнику не вменяется.").font = Font(size=9, italic=True, color=GREY)
    ind.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6); ind.row_dimensions[r].height = 28; ind.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center")

    warns = [w for w in (bal.get("warnings") or []) if w]
    if warns:
        r += 2; bar(ind, r, "ПРЕДУПРЕЖДЕНИЯ РАСЧЁТА", fill=AMBER, span=6)
        for w in warns:
            r += 1; ind.cell(r, 2, "⚠ " + str(w)).font = Font(size=9, color="FF7A5A00"); ind.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
            ind.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center"); ind.row_dimensions[r].height = 28

    # ── ПОТЕРИ ПО ПРИЧИНАМ ───────────────────────────────────────────
    rsn.column_dimensions["A"].width = 3
    for col, w in zip("BCDE", [30, 12, 14, 50]): rsn.column_dimensions[col].width = w
    rsn.cell(1, 1, "ПОТЕРИ РАБОЧЕГО ВРЕМЕНИ ПО ПРИЧИНАМ"); rsn.merge_cells("A1:E1")
    rsn.cell(1, 1).font = Font(bold=True, color=WHITE, size=13); rsn.cell(1, 1).fill = PatternFill("solid", fgColor=GREEN); rsn.row_dimensions[1].height = 26
    r = 3; hdr(rsn, r, (2, 3, 4, 5), ("Причина", "Мин", "% потерь", "Что это"))
    reasons = bal.get("lossesByReason", {}) or {}
    expl = {"ORG": "нет задания, ожидание людей, переделки, дублирование", "TECH": "сбой оборудования и ПО, отсутствие энергии",
            "INFO": "данные не найти, нет доступа, длительный поиск информации", "DISC": "нарушение режима: опоздание, посторонние занятия",
            "EXT": "сторонние организации: смежники, поставщики, клиенты"}
    total_loss = sum(_f(reasons.get(k, 0)) for k in REASON_RU)
    first_r = r + 1
    for k, name in REASON_RU.items():
        r += 1
        rsn.cell(r, 2, name).font = Font(size=10, bold=True)
        rsn.cell(r, 3, round(_f(reasons.get(k, 0)), 1)).number_format = "0"
        rsn.cell(r, 4, f"=IF(SUM($C${first_r}:$C${first_r+4})=0,0,C{r}/SUM($C${first_r}:$C${first_r+4}))").number_format = "0.0%"
        rsn.cell(r, 5, expl[k]).font = Font(size=9, color=GREY)
        for c in (3, 4): rsn.cell(r, c).alignment = Alignment(horizontal="center")
        for c in range(2, 6): rsn.cell(r, c).border = BORD
    r += 1; rsn.cell(r, 2, "Итого потерь, мин").font = Font(bold=True, size=10); rsn.cell(r, 3, f"=SUM(C{first_r}:C{r-1})").number_format = "0"
    rsn.cell(r, 3).alignment = Alignment(horizontal="center")
    for c in range(2, 6): rsn.cell(r, c).border = BORD; rsn.cell(r, c).fill = PatternFill("solid", fgColor=SUB)
    if total_loss <= 0:
        r += 1; rsn.cell(r, 2, "Потери не зафиксированы.").font = Font(size=9, italic=True, color=GREY)

    # крупнейшие потери
    losses = [x for x in recs if str(x.get("wt_index", "")).upper() in LOSS]
    losses.sort(key=lambda x: -_f(x.get("duration", 0)))
    if losses:
        r += 2; bar(rsn, r, "КРУПНЕЙШИЕ ПОТЕРИ", span=5)
        r += 1; hdr(rsn, r, (2, 3, 4, 5), ("Индекс", "Мин", "День", "Действие"))
        for x in losses[:10]:
            r += 1; idx = str(x.get("wt_index", "")).upper()
            rsn.cell(r, 2, f"{IDX_RU.get(idx, idx)} / {REASON_RU.get(str(x.get('loss_reason') or '').upper(), '—')}").font = Font(size=9, bold=True, color=IDX_COLOR.get(idx, "FF000000"))
            rsn.cell(r, 3, round(_f(x.get("duration", 0)), 1)).number_format = "0"; rsn.cell(r, 4, int(_f(x.get("day_number", 1)) or 1))
            rsn.cell(r, 5, x.get("activity", "")).font = Font(size=10)
            for c in (3, 4): rsn.cell(r, c).alignment = Alignment(horizontal="center")
            for c in range(2, 6): rsn.cell(r, c).border = BORD

    # ── ЛЕГЕНДА (на листе Баланс) ────────────────────────────────────
    r = ws.max_row + 2; bar(ws, r, "ЛЕГЕНДА", span=7)
    r += 1; ws.cell(r, 2, "Работа: ПЗ — подготовительно-заключительное · ОО — оперативное основное · ОВ — оперативное вспомогательное · ОБТ/ОБО — обслуживание рабочего места (техническое / организационное)")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7); ws.cell(r, 2).font = Font(size=9, color=GREY); ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center"); ws.row_dimensions[r].height = 30
    r += 1; ws.cell(r, 2, "Регламентированные перерывы: ОТЛ — отдых и личные надобности · ПТ — перерывы по технологии. Потери: НР — непроизводительная работа · СР — случайная работа · ПНТ — организационно-технические · ПНД — по вине работника")
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7); ws.cell(r, 2).font = Font(size=9, color=GREY); ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="center"); ws.row_dimensions[r].height = 30

    for s in (ws, ind, rsn, lst):
        for rowc in s.iter_rows():
            for cell in rowc:
                al = cell.alignment
                cell.alignment = Alignment(horizontal=al.horizontal, vertical="center", wrap_text=(s is not lst or cell.column == 7))
    return wb


if __name__ == "__main__":
    import io, json, sys
    src = sys.argv[1] if len(sys.argv) > 1 else None
    if src:
        d = json.load(open(src, encoding="utf-8"))
    else:
        d = {"position": "Специалист по кадрам", "department": "Отдел кадров", "industry": "Производство", "supervisor": "Начальник отдела кадров",
             "period": "01.09.2026 — 02.09.2026", "schedule_type": "рабочий день", "work_start": "08:30", "work_end": "17:30", "lunch_minutes": 60, "days": 2, "source": "self",
             "records": [
                 {"day_number": 1, "obs_date": "01.09.2026", "start": "08:30", "end": "08:45", "duration": 15, "activity": "Включение компьютера, запуск 1С", "wt_index": "PZ", "loss_reason": None, "needs_review": False},
                 {"day_number": 1, "obs_date": "01.09.2026", "start": "08:45", "end": "10:00", "duration": 75, "activity": "Оформление приказов о приёме", "wt_index": "OO", "loss_reason": None, "needs_review": False},
                 {"day_number": 1, "obs_date": "01.09.2026", "start": "10:00", "end": "10:20", "duration": 20, "activity": "Ожидание согласования у руководителя", "wt_index": "PNT", "loss_reason": "ORG", "needs_review": False},
                 {"day_number": 1, "obs_date": "01.09.2026", "start": "10:20", "end": "10:30", "duration": 10, "activity": "Перекур", "wt_index": "OTL", "loss_reason": None, "needs_review": False},
                 {"day_number": 2, "obs_date": "02.09.2026", "start": "08:30", "end": "09:30", "duration": 60, "activity": "Поиск личного дела в архиве", "wt_index": "NR", "loss_reason": "INFO", "needs_review": True},
                 {"day_number": 2, "obs_date": "02.09.2026", "start": "09:30", "end": "10:30", "duration": 60, "activity": "Ввод данных в 1С", "wt_index": "OO", "loss_reason": None, "needs_review": False},
             ],
             "balance": {"byIndex": {"PZ": 15, "OO": 135, "OV": 0, "OBT": 0, "OBO": 0, "OTL": 10, "PT": 0, "NR": 60, "SR": 0, "PNT": 20, "PND": 0},
                         "projected": {"PZ": 15, "OO": 215, "OV": 0, "OBT": 0, "OBO": 0, "OTL": 10, "PT": 0, "NR": 0, "SR": 0, "PNT": 0, "PND": 0},
                         "total": 240, "perDay": 120, "operative": 135, "servicing": 0, "normirovannoe": 160, "losses": 80,
                         "lossesByReason": {"ORG": 20, "TECH": 0, "INFO": 60, "DISC": 0, "EXT": 0},
                         "kisp": 66.7, "kpnt": 8.3, "kpnd": 0, "knr": 25, "checksum": 100, "ppt": 59.3,
                         "pptFormula": "ППТ = (Тпнт + Тпнд + Тнр + Тср) / Топ × 100 = 80 / 135 × 100 = 59.3 %",
                         "dKisp": 49.9, "reservePerDay": 40, "aob": 0, "aotl": 7.4, "tpzPerDay": 7.5,
                         "warnings": ["Резерв свыше 50 % обычно указывает либо на серьёзные организационные проблемы, либо на неполноту записей наблюдения; требуется повторное наблюдение"]}}
    wb = build_frd(d); buf = io.BytesIO(); wb.save(buf)
    out = sys.argv[2] if len(sys.argv) > 2 else "/tmp/FRD_BALANCE_v3_sample.xlsx"
    wb.save(out); print("OK", len(buf.getvalue()), "bytes |", wb.sheetnames, "→", out)
