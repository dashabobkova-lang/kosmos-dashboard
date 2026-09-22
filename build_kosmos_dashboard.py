#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Автосборка дашборда SALES REPORT WB KOSMOS.

Что делает:
  1. Через Bitrix REST обходит папки "ежедневные отчеты КОСМОС/<DD.MM.YY>"
  2. В каждой берёт "отчет по продажам *.XLSX" (Отчёт по данным поставщика WB)
  3. Из самой свежей папки берёт "отчет по остаткам*.xlsx" (снимок остатков/в пути)
  4. Собирает двухлистовой xlsx (SALES REPORT по дням + Запас)
  5. Заливает его обратно в Bitrix, заменяя существующий файл
  6. Шлёт краткую сводку за последний день в Telegram

Запуск:
  python3 build_kosmos_dashboard.py
Переменные окружения (обязательные):
  BITRIX_WEBHOOK   - https://<portal>/rest/<user>/<token>
  TG_TOKEN         - токен Telegram-бота
  TG_CHAT          - chat_id группы
Необязательные:
  DAILY_FOLDER_ID  - id папки с дневными отчётами (по умолчанию 101236)
  TARGET_FOLDER_ID - id папки, куда класть дашборд (по умолчанию 101232)
  TARGET_NAME      - имя файла дашборда
  ART_PREFIX       - префикс артикулов для фильтра (по умолчанию ДАБ)
"""

import io
import os
import re
import sys
import json
import zipfile
import datetime
import urllib.parse
import urllib.request

WEBHOOK = os.environ.get("BITRIX_WEBHOOK", "").rstrip("/")
TG_TOKEN = os.environ.get("TG_TOKEN", "")
TG_CHAT = os.environ.get("TG_CHAT", "")
DAILY_FOLDER_ID = os.environ.get("DAILY_FOLDER_ID", "101236")
TARGET_FOLDER_ID = os.environ.get("TARGET_FOLDER_ID", "101232")
TARGET_NAME = os.environ.get("TARGET_NAME", "SALES REPORT WB KOSMOS.xlsx")
ART_PREFIX = os.environ.get("ART_PREFIX", "ДАБ")

DASH_URL = ("https://bylulu.bitrix24.ru/docs/file/"
            "%D0%94%D0%90%D0%A8%D0%91%D0%9E%D0%A0%D0%94%D0%AB%20%D0%98%20"
            "%D0%9E%D0%A2%D0%A7%D0%95%D0%A2%D0%AB/"
            "SALES%20REPORT%20WB%20KOSMOS.xlsx")


def log(msg):
    print("[%s] %s" % (datetime.datetime.now().strftime("%H:%M:%S"), msg), flush=True)


def die(msg):
    log("ОШИБКА: " + msg)
    notify_error(msg)
    sys.exit(1)


# ---------------------------------------------------------------- Bitrix REST
def rest(method, params=None):
    url = "%s/%s.json" % (WEBHOOK, method)
    data = urllib.parse.urlencode(params or {}).encode("utf-8")
    req = urllib.request.Request(url, data=data)
    with urllib.request.urlopen(req, timeout=120) as r:
        return json.loads(r.read().decode("utf-8"))


def get_children(folder_id):
    out, start = [], 0
    while True:
        r = rest("disk.folder.getchildren", {"id": folder_id, "start": start})
        out.extend(r.get("result", []))
        if "next" not in r:
            break
        start = r["next"]
    return out


def download(url):
    with urllib.request.urlopen(url, timeout=300) as r:
        return r.read()


def upload_file(folder_id, name, blob):
    """Загружает файл в папку. Если файл с таким именем есть - перезаписывает."""
    existing = None
    for it in get_children(folder_id):
        if it.get("TYPE") == "file" and it.get("NAME") == name:
            existing = it["ID"]
            break
    if existing:
        r = rest("disk.file.uploadversion", {"id": existing})
    else:
        r = rest("disk.folder.uploadfile", {"id": folder_id, "data[NAME]": name})
    res = r.get("result") or {}
    upload_url, field = res.get("uploadUrl"), res.get("field", "file")
    if not upload_url:
        die("не получен uploadUrl: %s" % json.dumps(r, ensure_ascii=False)[:300])

    boundary = "----kosmos%d" % int(datetime.datetime.now().timestamp())
    body = b""
    body += ("--%s\r\n" % boundary).encode()
    body += ('Content-Disposition: form-data; name="%s"; filename="%s"\r\n'
             % (field, name)).encode("utf-8")
    body += b"Content-Type: application/octet-stream\r\n\r\n"
    body += blob
    body += ("\r\n--%s--\r\n" % boundary).encode()
    req = urllib.request.Request(upload_url, data=body)
    req.add_header("Content-Type", "multipart/form-data; boundary=%s" % boundary)
    with urllib.request.urlopen(req, timeout=300) as r:
        return json.loads(r.read().decode("utf-8"))


# ------------------------------------------------------------- xlsx чтение
CELL_RX = re.compile(
    r'<c r="([A-Z]+)\d+"[^>]*?(?:/>|>(?:<is><t[^>]*>(.*?)</t></is>|<v>(.*?)</v>)?</c>)',
    re.S)
ROW_RX = re.compile(r'<row r="(\d+)"[^>]*>(.*?)</row>', re.S)


def unescape(s):
    if not s:
        return ""
    return (s.replace("&lt;", "<").replace("&gt;", ">")
             .replace("&quot;", '"').replace("&apos;", "'").replace("&amp;", "&"))


def read_sheet_rows(blob):
    """Возвращает список (row_number, {col_letter: value})."""
    with zipfile.ZipFile(io.BytesIO(blob)) as z:
        names = [n for n in z.namelist() if n.startswith("xl/worksheets/sheet")]
        names.sort()
        xml = z.read(names[0]).decode("utf-8", "ignore")
    i, j = xml.find("<sheetData>"), xml.find("</sheetData>")
    body = xml[i:j] if i >= 0 and j > i else xml
    out = []
    for m in ROW_RX.finditer(body):
        rn = int(m.group(1))
        cells = {}
        for c in CELL_RX.finditer(m.group(2)):
            col = c.group(1)
            val = c.group(2) if c.group(2) is not None else (c.group(3) or "")
            cells[col] = unescape(val)
        out.append((rn, cells))
    return out


def num(s):
    if s is None or s == "":
        return 0.0
    try:
        return float(str(s).replace(",", ".").replace("\xa0", "").replace(" ", ""))
    except ValueError:
        return 0.0


# ------------------------------------------------------------ сбор данных
def collect():
    log("читаю список папок с дневными отчётами...")
    folders = []
    for it in get_children(DAILY_FOLDER_ID):
        if it.get("TYPE") != "folder":
            continue
        m = re.match(r"^(\d{2})\.(\d{2})\.(\d{2})$", it["NAME"].strip())
        if not m:
            log("  пропускаю папку с нестандартным именем: %s" % it["NAME"])
            continue
        d, mo, y = m.groups()
        folders.append(("20%s-%s-%s" % (y, mo, d), it["ID"], it["NAME"]))
    folders.sort()
    if not folders:
        die("не найдено ни одной папки вида DD.MM.YY в папке %s" % DAILY_FOLDER_ID)

    agg, dates, missing = {}, [], []
    stock_src = None  # (date, blob) самый свежий отчёт по остаткам

    for date, fid, fname in folders:
        kids = get_children(fid)
        sales = [k for k in kids if k.get("TYPE") == "file"
                 and re.search(r"по\s*продажам", k["NAME"], re.I)]
        stocks = [k for k in kids if k.get("TYPE") == "file"
                  and re.search(r"по\s*остатк", k["NAME"], re.I)]
        if stocks:
            stock_src = (date, download(stocks[0]["DOWNLOAD_URL"]))
        if not sales:
            missing.append(fname)
            continue
        blob = download(sales[0]["DOWNLOAD_URL"])
        rows = read_sheet_rows(blob)
        # схема supplier-goods: F артикул, G артикул WB, A бренд, B предмет,
        # K склад, L заказано шт, M заказы руб, N выкуплено шт, O к перечислению, P остаток
        cnt = 0
        for rn, c in rows:
            if rn < 3:
                continue
            art = (c.get("F") or "").strip()
            if not art or not art.startswith(ART_PREFIX):
                continue
            a = agg.setdefault(art, {"art": art, "nm": c.get("G", ""),
                                     "brand": c.get("A", ""), "subj": c.get("B", ""),
                                     "o": {}, "orub": {}, "b": {}, "pay": {},
                                     "stock": {}, "fbs": {}})
            a["o"][date] = a["o"].get(date, 0) + num(c.get("L"))
            a["orub"][date] = a["orub"].get(date, 0) + num(c.get("M"))
            a["b"][date] = a["b"].get(date, 0) + num(c.get("N"))
            a["pay"][date] = a["pay"].get(date, 0) + num(c.get("O"))
            a["stock"][date] = a["stock"].get(date, 0) + num(c.get("P"))
            if "Свой" in (c.get("K") or ""):
                a["fbs"][date] = a["fbs"].get(date, 0) + num(c.get("P"))
            cnt += 1
        dates.append(date)
        log("  %s: строк %d" % (fname, cnt))

    if missing:
        log("ВНИМАНИЕ: нет отчёта по продажам в папках: %s" % ", ".join(missing))
    if not dates:
        die("ни в одной папке не найден 'отчет по продажам'")

    # остатки в пути из самого свежего отчёта по остаткам
    transit = {}
    if stock_src:
        sd, sblob = stock_src
        for rn, c in read_sheet_rows(sblob):
            if rn < 2:
                continue
            art = (c.get("C") or "").strip()
            if not art.startswith(ART_PREFIX):
                continue
            t = transit.setdefault(art, {"to": 0.0, "from": 0.0})
            t["to"] += num(c.get("H"))     # В пути до получателей
            t["from"] += num(c.get("I"))   # В пути возвраты на склад WB
        log("остатки в пути: из отчёта за %s, артикулов %d" % (sd, len(transit)))
    else:
        log("ВНИМАНИЕ: отчёт по остаткам не найден, колонки 'в пути' будут пустыми")

    return agg, sorted(set(dates)), transit, missing


# ------------------------------------------------------------ сборка xlsx
def colname(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def esc(t):
    return (str(t or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def serial(d):
    return (datetime.date(*map(int, d.split("-"))) - datetime.date(1899, 12, 30)).days


STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="3"><numFmt numFmtId="164" formatCode="dd.mm"/><numFmt numFmtId="165" formatCode="0%"/><numFmt numFmtId="166" formatCode="0.0"/></numFmts>
<fonts count="3"><font><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><name val="Arial"/></font><font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts>
<fills count="8">
<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFC2327A"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFF2DCE9"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFECECEC"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFE8F1FB"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
<fill><patternFill patternType="solid"><fgColor rgb="FFFAE3D0"/><bgColor indexed="64"/></patternFill></fill>
</fills>
<borders count="3"><border><left/><right/><top/><bottom/><diagonal/></border>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="thin"><color rgb="FFBFBFBF"/></bottom><diagonal/></border>
<border><left style="thin"><color rgb="FFBFBFBF"/></left><right style="thin"><color rgb="FFBFBFBF"/></right><top style="thin"><color rgb="FFBFBFBF"/></top><bottom style="medium"><color rgb="FF7F7F7F"/></bottom><diagonal/></border></borders>
<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
<cellXfs count="29">
<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
<xf numFmtId="3" fontId="1" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="165" fontId="1" fillId="4" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="164" fontId="1" fillId="3" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center"/></xf>
<xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf>
<xf numFmtId="49" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="1" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="3" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="166" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="49" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="1" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="3" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="166" fontId="0" fillId="5" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="49" fontId="0" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="1" fontId="0" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="3" fontId="0" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="166" fontId="0" fillId="6" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="49" fontId="0" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="1" fontId="0" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="3" fontId="0" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="165" fontId="0" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="166" fontId="0" fillId="5" borderId="2" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="49" fontId="1" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="1" fontId="1" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
<xf numFmtId="166" fontId="1" fillId="7" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/>
</cellXfs>
<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>'''

# индексы стилей
SA = {"t": 6, "i": 7, "m": 8, "p": 9, "a": 10}
SB = {"t": 11, "i": 12, "m": 13, "p": 14, "a": 15}
SAT = {"t": 16, "i": 17, "m": 18, "p": 19, "a": 20}
SBT = {"t": 21, "i": 22, "m": 23, "p": 24, "a": 25}
CAT_T, CAT_I, CAT_A = 26, 27, 28

CATEGORIES = [
    ("ВЕТРОВКИ", ["*VETR-*"]),
    ("КУРТКИ", ["*JACK-*"]),
    ("ЛЕН", ["*LN-*"]),
    ("ПОДАРКИ", ["*gift-*"]),
    ("СПОРТИВНЫЕ КОСТЮМЫ", ["*MOD-*"]),
    ("ВЯЗАНЫЕ КОСТЮМЫ", ["*ORB-*", "*SJP-*"]),
]


def group_key(art):
    a = art.strip()
    return a.split("/")[0].strip() if "/" in a else a


def build_sheet1(rows, dates, transit):
    nd = len(dates)
    C = {"photo": 1, "art": 2, "wb": 3, "subj": 4, "brand": 5,
         "stock": 6, "fbs": 7, "toc": 8, "fromc": 9}
    ord_cols = list(range(10, 10 + nd))
    c_otot, c_oavg, c_o7 = 10 + nd, 11 + nd, 12 + nd
    buy_cols = list(range(13 + nd, 13 + 2 * nd))
    c_btot, c_bavg, c_b7 = 13 + 2 * nd, 14 + 2 * nd, 15 + 2 * nd
    c_pc, c_pcdel = 16 + 2 * nd, 17 + 2 * nd
    c_chk, c_days, c_share = 18 + 2 * nd, 19 + 2 * nd, 20 + 2 * nd
    last = c_share
    r0, r1 = 5, 4 + len(rows)
    oL, oR = colname(ord_cols[0]), colname(ord_cols[-1])
    o7c = colname(ord_cols[max(0, nd - 7)])
    bL, bR = colname(buy_cols[0]), colname(buy_cols[-1])
    b7c = colname(buy_cols[max(0, nd - 7)])
    oT, bT, o7L = colname(c_otot), colname(c_btot), colname(c_o7)
    mat_end = colname(ord_cols[max(0, nd - 8)])

    last_date = dates[-1]
    g_stock = sum(r["stock"].get(last_date, 0) for r in rows)
    g_fbs = sum(r["fbs"].get(last_date, 0) for r in rows)
    g_to = sum(transit.get(r["art"], {}).get("to", 0) for r in rows)
    g_from = sum(transit.get(r["art"], {}).get("from", 0) for r in rows)
    g_orub = sum(sum(r["orub"].values()) for r in rows)
    g_osht = sum(sum(r["o"].values()) for r in rows)
    g_chk = round(g_orub / g_osht) if g_osht else 0

    L = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
         '<sheetViews><sheetView workbookViewId="0"><pane xSplit="9" ySplit="4" '
         'topLeftCell="%s5" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>' % oL,
         '<sheetFormatPr defaultRowHeight="14"/>']
    L.append('<cols><col min="1" max="1" width="8" customWidth="1"/>'
             '<col min="2" max="2" width="24" customWidth="1"/>'
             '<col min="3" max="3" width="12" customWidth="1"/>'
             '<col min="4" max="4" width="16" customWidth="1"/>'
             '<col min="5" max="5" width="12" customWidth="1"/>'
             '<col min="6" max="9" width="9" customWidth="1"/>'
             '<col min="10" max="%d" width="10" customWidth="1"/>'
             '<col min="%d" max="%d" width="9" customWidth="1"/>'
             '<col min="%d" max="%d" width="11" customWidth="1"/>'
             '<col min="%d" max="%d" width="14" customWidth="1"/></cols>'
             % (buy_cols[-1], c_otot, c_o7, c_btot, c_days, c_share, c_share))
    L.append("<sheetData>")

    title = ("SALES REPORT по дням — данные поставщика ООО «КОСМОС», %s–%s (только %s)"
             % (dates[0][8:10] + "." + dates[0][5:7],
                dates[-1][8:10] + "." + dates[-1][5:7] + "." + dates[-1][:4], ART_PREFIX))
    row = ('<row r="1"><c r="A1" s="0" t="inlineStr"><is><t>%s</t></is></c>' % esc(title))
    for n in range(2, last + 1):
        if n == ord_cols[0]:
            row += '<c r="%s1" s="5" t="inlineStr"><is><t>ЗАКАЗЫ</t></is></c>' % colname(n)
        elif n == buy_cols[0]:
            row += '<c r="%s1" s="5" t="inlineStr"><is><t>ВЫКУПЫ</t></is></c>' % colname(n)
        else:
            row += '<c r="%s1" s="0"/>' % colname(n)
    L.append(row + "</row>")

    heads = {1: "Фото", 2: "Артикул продавца", 3: "Артикул WB", 4: "Предмет", 5: "Бренд",
             6: "Остаток, шт", 7: "Остаток FBS, шт", 8: "В пути до клиента, шт",
             9: "В пути от клиента, шт",
             c_otot: "Итого", c_oavg: "Ср/день", c_o7: "Посл. 7 дн",
             c_btot: "Итого", c_bavg: "Ср/день", c_b7: "Посл. 7 дн",
             c_pc: "% выкупа", c_pcdel: "% выкупа (учёт доставки 7 дн)",
             c_chk: "Ср. чек, ₽", c_days: "Хватит на, дн", c_share: "Доля в продажах, %"}
    row = '<row r="2" ht="34">'
    for n in range(1, 10):
        row += '<c r="%s2" s="1" t="inlineStr"><is><t>%s</t></is></c>' % (colname(n), esc(heads[n]))
    for i, d in enumerate(dates):
        row += '<c r="%s2" s="4"><v>%d</v></c>' % (colname(ord_cols[i]), serial(d))
    for n in (c_otot, c_oavg, c_o7):
        row += '<c r="%s2" s="1" t="inlineStr"><is><t>%s</t></is></c>' % (colname(n), esc(heads[n]))
    for i, d in enumerate(dates):
        row += '<c r="%s2" s="4"><v>%d</v></c>' % (colname(buy_cols[i]), serial(d))
    for n in (c_btot, c_bavg, c_b7, c_pc, c_pcdel, c_chk, c_days, c_share):
        row += '<c r="%s2" s="1" t="inlineStr"><is><t>%s</t></is></c>' % (colname(n), esc(heads[n]))
    L.append(row + "</row>")

    # строка 3 - ИТОГО в рублях по дням
    row = '<row r="3" ht="15"><c r="A3" s="2" t="inlineStr"><is><t>ИТОГО, ₽</t></is></c>'
    for n in range(2, 10):
        row += '<c r="%s3" s="2"/>' % colname(n)
    for i, d in enumerate(dates):
        tot = sum(r["orub"].get(d, 0) for r in rows)
        row += '<c r="%s3" s="2"><v>%d</v></c>' % (colname(ord_cols[i]), round(tot))
    row += '<c r="%s3" s="2"><f>SUM(%s3:%s3)</f></c><c r="%s3" s="2"/><c r="%s3" s="2"/>' % (oT, oL, oR, colname(c_oavg), o7L)
    for i, d in enumerate(dates):
        tot = sum(r["pay"].get(d, 0) for r in rows)
        row += '<c r="%s3" s="2"><v>%d</v></c>' % (colname(buy_cols[i]), round(tot))
    row += '<c r="%s3" s="2"><f>SUM(%s3:%s3)</f></c>' % (bT, bL, bR)
    for n in (c_bavg, c_b7, c_pc, c_pcdel, c_chk, c_days, c_share):
        row += '<c r="%s3" s="2"/>' % colname(n)
    L.append(row + "</row>")

    # строка 4 - ИТОГО в штуках
    row = '<row r="4" ht="15"><c r="A4" s="2" t="inlineStr"><is><t>ИТОГО, шт</t></is></c>'
    for n in range(2, 6):
        row += '<c r="%s4" s="2"/>' % colname(n)
    for n, v in ((C["stock"], g_stock), (C["fbs"], g_fbs), (C["toc"], g_to), (C["fromc"], g_from)):
        row += '<c r="%s4" s="2"><v>%d</v></c>' % (colname(n), round(v))
    for n in ord_cols + [c_otot]:
        cl = colname(n)
        row += '<c r="%s4" s="2"><f>SUM(%s%d:%s%d)</f></c>' % (cl, cl, r0, cl, r1)
    row += '<c r="%s4" s="2"><f>ROUND(%s4/%d,1)</f></c>' % (colname(c_oavg), oT, nd)
    row += '<c r="%s4" s="2"><f>SUM(%s%d:%s%d)</f></c>' % (o7L, o7L, r0, o7L, r1)
    for n in buy_cols + [c_btot]:
        cl = colname(n)
        row += '<c r="%s4" s="2"><f>SUM(%s%d:%s%d)</f></c>' % (cl, cl, r0, cl, r1)
    row += '<c r="%s4" s="2"><f>ROUND(%s4/%d,1)</f></c>' % (colname(c_bavg), bT, nd)
    b7L = colname(c_b7)
    row += '<c r="%s4" s="2"><f>SUM(%s%d:%s%d)</f></c>' % (b7L, b7L, r0, b7L, r1)
    row += '<c r="%s4" s="3"><f>IFERROR(%s4/%s4,&quot;&quot;)</f></c>' % (colname(c_pc), bT, oT)
    row += ('<c r="%s4" s="3"><f>IFERROR(%s4/SUM(%s4:%s4),&quot;&quot;)</f></c>'
            % (colname(c_pcdel), bT, oL, mat_end))
    row += '<c r="%s4" s="2"><v>%d</v></c>' % (colname(c_chk), g_chk)
    row += ('<c r="%s4" s="2"><f>IFERROR(%s4/(%s4/%d),&quot;&quot;)</f></c>'
            % (colname(c_days), colname(C["stock"]), oT, nd))
    row += '<c r="%s4" s="3"><f>IFERROR(%s4/%s4,&quot;&quot;)</f></c>' % (colname(c_share), bT, bT)
    L.append(row + "</row>")

    prev_g, band = None, False
    for idx, r in enumerate(rows):
        rn = r0 + idx
        g = group_key(r["art"])
        if idx == 0:
            prev_g = g
        elif g != prev_g:
            band = not band
            prev_g = g
        last_of_group = (idx == len(rows) - 1) or (group_key(rows[idx + 1]["art"]) != g)
        S = (SBT if last_of_group else SB) if band else (SAT if last_of_group else SA)
        card = "https://www.wildberries.ru/catalog/%s/detail.aspx" % r["nm"]
        tr = transit.get(r["art"], {"to": 0, "from": 0})
        osht = sum(r["o"].values())
        orub = sum(r["orub"].values())
        chk = round(orub / osht) if osht else ""
        row = '<row r="%d"><c r="A%d" s="%d"><f>HYPERLINK(&quot;%s&quot;,&quot;фото ↗&quot;)</f></c>' % (rn, rn, S["t"], card)
        for cl, val in (("B", r["art"]), ("C", r["nm"]), ("D", r["subj"]), ("E", r["brand"])):
            row += '<c r="%s%d" s="%d" t="inlineStr"><is><t>%s</t></is></c>' % (cl, rn, S["t"], esc(val))
        for n, v in ((C["stock"], r["stock"].get(last_date, 0)), (C["fbs"], r["fbs"].get(last_date, 0)),
                     (C["toc"], tr["to"]), (C["fromc"], tr["from"])):
            row += '<c r="%s%d" s="%d"><v>%d</v></c>' % (colname(n), rn, S["i"], round(v))
        for i, d in enumerate(dates):
            row += '<c r="%s%d" s="%d"><v>%d</v></c>' % (colname(ord_cols[i]), rn, S["i"], round(r["o"].get(d, 0)))
        row += '<c r="%s%d" s="%d"><f>SUM(%s%d:%s%d)</f></c>' % (oT, rn, S["i"], oL, rn, oR, rn)
        row += '<c r="%s%d" s="%d"><f>ROUND(SUM(%s%d:%s%d)/%d,1)</f></c>' % (colname(c_oavg), rn, S["a"], oL, rn, oR, rn, nd)
        row += '<c r="%s%d" s="%d"><f>SUM(%s%d:%s%d)</f></c>' % (o7L, rn, S["i"], o7c, rn, oR, rn)
        for i, d in enumerate(dates):
            row += '<c r="%s%d" s="%d"><v>%d</v></c>' % (colname(buy_cols[i]), rn, S["i"], round(r["b"].get(d, 0)))
        row += '<c r="%s%d" s="%d"><f>SUM(%s%d:%s%d)</f></c>' % (bT, rn, S["i"], bL, rn, bR, rn)
        row += '<c r="%s%d" s="%d"><f>ROUND(SUM(%s%d:%s%d)/%d,1)</f></c>' % (colname(c_bavg), rn, S["a"], bL, rn, bR, rn, nd)
        row += '<c r="%s%d" s="%d"><f>SUM(%s%d:%s%d)</f></c>' % (b7L, rn, S["i"], b7c, rn, bR, rn)
        row += ('<c r="%s%d" s="%d"><f>IFERROR(SUM(%s%d:%s%d)/SUM(%s%d:%s%d),&quot;&quot;)</f></c>'
                % (colname(c_pc), rn, S["p"], bL, rn, bR, rn, oL, rn, oR, rn))
        row += ('<c r="%s%d" s="%d"><f>IFERROR(SUM(%s%d:%s%d)/SUM(%s%d:%s%d),&quot;&quot;)</f></c>'
                % (colname(c_pcdel), rn, S["p"], bL, rn, bR, rn, oL, rn, mat_end, rn))
        if chk != "":
            row += '<c r="%s%d" s="%d"><v>%d</v></c>' % (colname(c_chk), rn, S["m"], chk)
        else:
            row += '<c r="%s%d" s="%d"/>' % (colname(c_chk), rn, S["m"])
        row += ('<c r="%s%d" s="%d"><f>IFERROR(%s%d/(%s%d/%d),&quot;&quot;)</f></c>'
                % (colname(c_days), rn, S["a"], colname(C["stock"]), rn, oT, rn, nd))
        row += ('<c r="%s%d" s="%d"><f>IFERROR(%s%d/$%s$4,&quot;&quot;)</f></c>'
                % (colname(c_share), rn, S["p"], bT, rn, bT))
        L.append(row + "</row>")

    L.append("</sheetData>")
    L.append('<mergeCells count="2"><mergeCell ref="%s1:%s1"/><mergeCell ref="%s1:%s1"/></mergeCells>'
             % (oL, oR, bL, bR))
    scale = ('<colorScale><cfvo type="min"/><cfvo type="max"/>'
             '<color rgb="FFFFF3E6"/><color rgb="FFF57C00"/></colorScale>')
    for pr, (a, b) in enumerate([(oL + str(r0), oR + str(r1)),
                                 (bL + str(r0), bR + str(r1)),
                                 (bT + str(r0), bT + str(r1))], start=1):
        L.append('<conditionalFormatting sqref="%s:%s"><cfRule type="colorScale" priority="%d">%s</cfRule></conditionalFormatting>'
                 % (a, b, pr, scale))
    L.append('<ignoredErrors><ignoredError sqref="A1:%s%d" numberStoredAsText="1" formula="1" '
             'formulaRange="1" emptyCellReference="1" unlockedFormula="1" calculatedColumn="1"/></ignoredErrors>'
             % (colname(last), r1))
    L.append("</worksheet>")
    return "".join(L)


def build_sheet2(rows, dates, transit):
    nd = len(dates)
    last7 = dates[max(0, nd - 7):]
    last_date = dates[-1]
    r0, r1 = 3, 2 + len(rows)
    heads = ["Фото", "Артикул продавца", "Предмет", "Бренд", "Остаток FBS, шт",
             "Остаток FBO, шт", "К клиенту (в пути), шт", "От клиента (в пути), шт",
             "Остатки ИТОГО, шт", "Продажи за посл. 7 дней, шт", "Запас, нед"]
    L = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
         '<sheetViews><sheetView workbookViewId="0"><pane xSplit="4" ySplit="2" topLeftCell="E3" '
         'activePane="bottomRight" state="frozen"/></sheetView></sheetViews>',
         '<sheetFormatPr defaultRowHeight="14"/>',
         '<cols><col min="1" max="1" width="8" customWidth="1"/>'
         '<col min="2" max="2" width="24" customWidth="1"/>'
         '<col min="3" max="3" width="16" customWidth="1"/>'
         '<col min="4" max="4" width="22" customWidth="1"/>'
         '<col min="5" max="11" width="13" customWidth="1"/></cols>',
         "<sheetData>"]
    row = '<row r="1" ht="30">'
    for i, h in enumerate(heads):
        row += '<c r="%s1" s="1" t="inlineStr"><is><t>%s</t></is></c>' % (colname(i + 1), esc(h))
    L.append(row + "</row>")
    row = ('<row r="2"><c r="A2" s="2" t="inlineStr"><is><t>ИТОГО</t></is></c>'
           '<c r="B2" s="2"/><c r="C2" s="2"/><c r="D2" s="2"/>')
    for cl in "EFGHIJ":
        row += '<c r="%s2" s="2"><f>SUM(%s%d:%s%d)</f></c>' % (cl, cl, r0, cl, r1)
    row += '<c r="K2" s="2"><f>IFERROR(I2/J2,&quot;&quot;)</f></c></row>'
    L.append(row)

    for idx, r in enumerate(rows):
        rn = r0 + idx
        S = SB if rn % 2 == 0 else SA
        card = "https://www.wildberries.ru/catalog/%s/detail.aspx" % r["nm"]
        tr = transit.get(r["art"], {"to": 0, "from": 0})
        total = r["stock"].get(last_date, 0)
        fbs = r["fbs"].get(last_date, 0)
        fbo = total - fbs
        s7 = sum(r["b"].get(d, 0) for d in last7)
        row = '<row r="%d"><c r="A%d" s="%d"><f>HYPERLINK(&quot;%s&quot;,&quot;фото ↗&quot;)</f></c>' % (rn, rn, S["t"], card)
        for cl, val in (("B", r["art"]), ("C", r["subj"]), ("D", r["brand"])):
            row += '<c r="%s%d" s="%d" t="inlineStr"><is><t>%s</t></is></c>' % (cl, rn, S["t"], esc(val))
        for cl, v in (("E", fbs), ("F", fbo), ("G", tr["to"]), ("H", tr["from"])):
            row += '<c r="%s%d" s="%d"><v>%d</v></c>' % (cl, rn, S["i"], round(v))
        row += '<c r="I%d" s="%d"><f>E%d+F%d+G%d+H%d</f></c>' % (rn, S["i"], rn, rn, rn, rn)
        row += '<c r="J%d" s="%d"><v>%d</v></c>' % (rn, S["i"], round(s7))
        row += '<c r="K%d" s="%d"><f>IFERROR(I%d/J%d,&quot;&quot;)</f></c></row>' % (rn, S["a"], rn, rn)
        L.append(row)

    cs = r1 + 2
    cat_rows = []
    for i, (name, pats) in enumerate(CATEGORIES):
        rr = cs + i
        cat_rows.append(rr)
        row = ('<row r="%d"><c r="A%d" s="%d"/><c r="B%d" s="%d"/><c r="C%d" s="%d"/>'
               '<c r="D%d" s="%d" t="inlineStr"><is><t>%s</t></is></c>'
               % (rr, rr, CAT_T, rr, CAT_T, rr, CAT_T, rr, CAT_T, esc(name)))
        for cl in "EFGHIJ":
            parts = ['SUMIF($B$%d:$B$%d,&quot;%s&quot;,%s$%d:%s$%d)' % (r0, r1, p, cl, r0, cl, r1)
                     for p in pats]
            row += '<c r="%s%d" s="%d"><f>%s</f></c>' % (cl, rr, CAT_I, "+".join(parts))
        row += '<c r="K%d" s="%d"><f>IFERROR(I%d/J%d,&quot;&quot;)</f></c></row>' % (rr, CAT_A, rr, rr)
        L.append(row)

    ro = cs + len(CATEGORIES)
    row = ('<row r="%d"><c r="A%d" s="%d"/><c r="B%d" s="%d"/><c r="C%d" s="%d"/>'
           '<c r="D%d" s="%d" t="inlineStr"><is><t>ПРОЧЕЕ</t></is></c>'
           % (ro, ro, CAT_T, ro, CAT_T, ro, CAT_T, ro, CAT_T))
    for cl in "EFGHIJ":
        subs = "+".join("%s%d" % (cl, x) for x in cat_rows)
        row += '<c r="%s%d" s="%d"><f>%s2-(%s)</f></c>' % (cl, ro, CAT_I, cl, subs)
    row += '<c r="K%d" s="%d"><f>IFERROR(I%d/J%d,&quot;&quot;)</f></c></row>' % (ro, CAT_A, ro, ro)
    L.append(row)

    L.append("</sheetData>")
    L.append('<ignoredErrors><ignoredError sqref="A1:K%d" numberStoredAsText="1" formula="1" '
             'formulaRange="1" emptyCellReference="1" unlockedFormula="1" calculatedColumn="1"/></ignoredErrors>' % ro)
    L.append("</worksheet>")
    return "".join(L)


def build_xlsx(rows, dates, transit):
    now = datetime.datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")
    parts = {
        "[Content_Types].xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
            '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
            '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
            "</Types>",
        "_rels/.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
            "</Relationships>",
        "docProps/core.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
            'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            '<dc:title>SALES REPORT WB KOSMOS</dc:title><dc:creator>auto</dc:creator>'
            '<cp:lastModifiedBy>auto</cp:lastModifiedBy>'
            '<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>'
            '<dcterms:modified xsi:type="dcterms:W3CDTF">%s</dcterms:modified></cp:coreProperties>' % (now, now),
        "docProps/app.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
            'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
            "<Application>Microsoft Excel</Application><DocSecurity>0</DocSecurity><ScaleCrop>false</ScaleCrop>"
            '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Листы</vt:lpstr></vt:variant>'
            "<vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs>"
            '<TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>SALES REPORT</vt:lpstr>'
            "<vt:lpstr>Запас</vt:lpstr></vt:vector></TitlesOfParts>"
            "<LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc>"
            "<HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion></Properties>",
        "xl/workbook.xml":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="SALES REPORT" sheetId="1" r:id="rId1"/>'
            '<sheet name="Запас" sheetId="2" r:id="rId2"/></sheets>'
            '<calcPr fullCalcOnLoad="1"/></workbook>',
        "xl/_rels/workbook.xml.rels":
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>'
            '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            "</Relationships>",
        "xl/styles.xml": STYLES,
        "xl/worksheets/sheet1.xml": build_sheet1(rows, dates, transit),
        "xl/worksheets/sheet2.xml": build_sheet2(rows, dates, transit),
    }
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for name, content in parts.items():
            z.writestr(name, content.encode("utf-8"))
    return buf.getvalue()


# ---------------------------------------------------------------- Telegram
def tg_send(text):
    if not TG_TOKEN or not TG_CHAT:
        log("Telegram не настроен - пропускаю отправку")
        return
    payload = json.dumps({"chat_id": TG_CHAT, "text": text, "parse_mode": "HTML"}).encode("utf-8")
    req = urllib.request.Request("https://api.telegram.org/bot%s/sendMessage" % TG_TOKEN, data=payload)
    req.add_header("Content-Type", "application/json; charset=utf-8")
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            res = json.loads(r.read().decode("utf-8"))
            log("Telegram: ok=%s" % res.get("ok"))
    except Exception as e:
        log("Telegram ошибка: %s" % e)


def notify_error(msg):
    try:
        tg_send("⚠️ <b>Дашборд КОСМОС не собрался</b>\n\n<pre>%s</pre>" % esc(msg)[:900])
    except Exception:
        pass


def fmt(n):
    return "{:,.0f}".format(n).replace(",", " ")


def delta(cur, prev):
    if not prev:
        return "⚪ —"
    p = round(100.0 * (cur - prev) / prev, 1)
    if p > 0:
        return "🟢 +%s%%" % p
    if p < 0:
        return "🔴 %s%%" % p
    return "⚪ 0%"


def send_summary(rows, dates, missing):
    d, prev = dates[-1], (dates[-2] if len(dates) > 1 else None)
    cur = {k: sum(r[k].get(d, 0) for r in rows) for k in ("o", "orub", "b", "pay")}
    old = {k: sum(r[k].get(prev, 0) for r in rows) for k in ("o", "orub", "b", "pay")} if prev else None
    lines = []
    for label, key, unit in (("ЗАКАЗЫ", "o", "шт."), ("СУММА ЗАКАЗОВ", "orub", "₽"),
                             ("ПРОДАЖИ", "b", "шт."), ("СУММА ПРОДАЖ", "pay", "₽")):
        val = "%s %s" % (fmt(cur[key]), unit)
        dl = delta(cur[key], old[key]) if old else ""
        lines.append("%-15s%-13s%s" % (label, val, dl))
    dd = "%s.%s.%s" % (d[8:10], d[5:7], d[:4])
    text = ("<b>КОСМОС</b> — сводка за %s\n\n<pre>%s</pre>\n\nДашборд:\n%s"
            % (dd, "\n".join(lines), DASH_URL))
    if missing:
        text += "\n\n⚠️ Нет отчёта по продажам за: %s" % ", ".join(missing)
    tg_send(text)


# ---------------------------------------------------------------- main
def main():
    if not WEBHOOK:
        die("не задана переменная BITRIX_WEBHOOK")
    agg, dates, transit, missing = collect()
    rows = sorted(agg.values(), key=lambda r: r["art"])
    log("итог: дней %d (%s..%s), артикулов %d" % (len(dates), dates[0], dates[-1], len(rows)))
    blob = build_xlsx(rows, dates, transit)
    log("xlsx собран: %d байт" % len(blob))
    res = upload_file(TARGET_FOLDER_ID, TARGET_NAME, blob)
    fid = (res.get("result") or {}).get("ID")
    if not fid:
        die("не удалось загрузить файл: %s" % json.dumps(res, ensure_ascii=False)[:300])
    log("залит в Bitrix, file id=%s" % fid)
    send_summary(rows, dates, missing)
    log("готово")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        die("%s: %s" % (type(e).__name__, e))
