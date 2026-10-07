# -*- coding: utf-8 -*-
"""Полный ежедневный цикл дашборда КОСМОС — без чужих сервисов и без участия Дарьи.

  1. обходит публичную папку Google Диска с выгрузками WB (вместе с подпапками);
  2. качает новые файлы, тип определяет по содержимому, а не по имени;
  3. собирает xlsx с тремя листами;
  4. заливает его в Google Таблицу (сервисный аккаунт);
  5. шлёт сводку за вчера в Телеграм — один раз за день,
     повторно только если цифры изменились.

  python kosmos_daily.py           обычный прогон
  python kosmos_daily.py --force   собрать и отправить, даже если ничего не изменилось
"""
import datetime
import io
import json
import os
import re
import sys
import urllib.request
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import build_kosmos_dashboard as B          # noqa: E402
import build_local                          # noqa: E402

FOLDER_ID = "1zDp7Tp5InWwgWzr7ahvlpZVX5dZZldTl"
SHEET_URL = "https://docs.google.com/spreadsheets/d/1mLPTkAl84cYwKKPZN_IsnrSuMPYA3vXweajwoup5e80/edit"
WORK = os.environ.get("KOSMOS_WORK", os.path.join(HERE, "данные"))
OUT = os.path.join(HERE, "SALES REPORT WB KOSMOS.xlsx")
STATE = os.path.join(HERE, "kosmos_state.json")
LOG = os.path.join(HERE, "kosmos_daily.log")
ENV = os.path.join(os.path.dirname(HERE), ".env")

MSK = datetime.timezone(datetime.timedelta(hours=3))
LAST_TG = []   # причина, по которой не ушла сводка
HEAD_RX = re.compile(r"с\s+(\d{2})\.(\d{2})\.(\d{4})\s+по\s+(\d{2})\.(\d{2})\.(\d{4})")


def log(msg):
    line = "[%s] %s" % (datetime.datetime.now(MSK).strftime("%d.%m %H:%M:%S"), msg)
    print(line, flush=True)
    try:
        with io.open(LOG, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


B.log = log


# ----------------------------------------------------------------- Google Диск

def folder_items(folder_id):
    """Файлы и подпапки публичной папки: [(kind, id, name)]."""
    # &t=... и заголовки против кэша: иначе Google отдаёт список папки с задержкой
    url = ("https://drive.google.com/embeddedfolderview?id=%s&t=%d#list"
           % (folder_id, int(datetime.datetime.now().timestamp())))
    req = urllib.request.Request(url, headers={
        "User-Agent": "Mozilla/5.0",
        "Cache-Control": "no-cache, max-age=0",
        "Pragma": "no-cache"})
    with urllib.request.urlopen(req, timeout=120) as r:
        html = r.read().decode("utf-8", "ignore")
    out = []
    for m in re.finditer(r'https://drive\.google\.com/(file/d/|drive/folders/)([\w-]+)', html):
        kind = "file" if m.group(1).startswith("file") else "folder"
        item = (kind, m.group(2))
        if item not in [(k, i) for k, i, _ in out]:
            out.append((kind, m.group(2), ""))
    return out


def walk(folder_id, depth=0):
    """Все файлы папки и подпапок."""
    files = []
    for kind, fid, _ in folder_items(folder_id):
        if kind == "file":
            files.append(fid)
        elif depth < 3:
            files.extend(walk(fid, depth + 1))
    return files


def download(file_id):
    url = "https://drive.google.com/uc?export=download&id=%s" % file_id
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=300) as r:
        return r.read()


# ------------------------------------------------------------ распознавание

def sheet_head(blob, limit=4000):
    """Начало первого листа книги — по нему понимаем, что это за файл."""
    try:
        with zipfile.ZipFile(io.BytesIO(blob)) as z:
            names = sorted(n for n in z.namelist() if n.startswith("xl/worksheets/sheet"))
            xml = z.read(names[0]).decode("utf-8", "ignore")[:limit]
            shared = ""
            if "xl/sharedStrings.xml" in z.namelist():
                shared = z.read("xl/sharedStrings.xml").decode("utf-8", "ignore")[:limit]
    except Exception:
        return "", ""
    return xml, shared


def classify(blob):
    """('sales', дата) | ('stocks', None) | (None, None) — по содержимому файла."""
    xml, shared = sheet_head(blob)
    text = xml + shared
    if "данным поставщика" in text or "Выкупленные товары" in text:
        m = HEAD_RX.search(text)
        if not m:
            return None, None
        if m.group(1, 2, 3) != m.group(4, 5, 6):
            return None, None            # многодневная выгрузка
        return "sales", "%s-%s-%s" % (m.group(3), m.group(2), m.group(1))
    if "В пути до получателей" in text or "Всего находится на складах" in text:
        return "stocks", None
    return None, None


def made_at(blob):
    """Когда отчёт сформирован — чтобы из двух выгрузок за день брать свежую."""
    xml, shared = sheet_head(blob)
    m = re.search(r"сформирован\s+(\d{2})\.(\d{2})\.(\d{4})\s+(\d{2}):(\d{2}):(\d{2})", xml + shared)
    return "".join(m.groups()[::-1]) if m else ""


# ---------------------------------------------------------------- Telegram

def clean_token(raw):
    """Из секрета достаёт сам токен: туда легко попадает лишний текст из письма BotFather."""
    m = re.search(r"\d{6,12}:[A-Za-z0-9_-]{30,}", raw or "")
    return m.group(0) if m else (raw or "").strip()


def clean_chat(raw):
    """chat_id — число, возможно со знаком минус."""
    m = re.search(r"-?\d{5,}", raw or "")
    return m.group(0) if m else (raw or "").strip()


def tg_send(text):
    tok = clean_token(os.environ.get("TG_TOKEN", ""))
    chat = clean_chat(os.environ.get("TG_CHAT", ""))
    if not tok or not chat:
        try:
            for line in io.open(ENV, encoding="utf-8-sig"):
                if line.startswith("TELEGRAM_BOT_TOKEN="):
                    tok = tok or clean_token(line.split("=", 1)[1])
                elif line.startswith("TELEGRAM_CHAT_ID="):
                    chat = chat or clean_chat(line.split("=", 1)[1])
        except Exception:
            pass
    if not tok or not chat:
        LAST_TG.append("нет токена (%s) или chat_id (%s)" % (bool(tok), bool(chat)))
        log("Телеграм не настроен")
        return False
    payload = json.dumps({"chat_id": chat, "text": text, "parse_mode": "HTML",
                          "disable_web_page_preview": True}).encode("utf-8")
    req = urllib.request.Request("https://api.telegram.org/bot%s/sendMessage" % tok, data=payload)
    req.add_header("Content-Type", "application/json; charset=utf-8")
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            res = json.loads(r.read().decode("utf-8"))
        ok = res.get("ok")
        if not ok:
            LAST_TG.append(str(res)[:200])
        log("Телеграм: ok=%s" % ok)
        return bool(ok)
    except Exception as e:
        detail = ""
        try:
            detail = e.read().decode("utf-8", "ignore")[:200]
        except Exception:
            pass
        LAST_TG.append("%s %s" % (e, detail))
        log("Телеграм ошибка: %s %s" % (e, detail))
        return False


# ---------------------------------------------------------------- состояние

STATE_SHEET = "_состояние"


def _state_ws():
    import gspread
    import kosmos_gsheet as G
    gc = gspread.service_account(filename=G.KEY_PATH)
    sh = gc.open_by_key(G.SHEET_ID)
    try:
        return sh.worksheet(STATE_SHEET)
    except Exception:
        ws = sh.add_worksheet(title=STATE_SHEET, rows=400, cols=2)
        ws.update(range_name="A1", values=[["день", "итоги"]])
        return ws


def load_state():
    """Память о разосланных сводках живёт в самой таблице — она одна
    и для облака, и для запуска с компьютера, поэтому дублей не бывает."""
    try:
        rows = _state_ws().get_all_values()
    except Exception as e:
        log("состояние не прочитано (%s) — выхожу, чтобы не слать дубль" % e)
        sys.exit(1)
    out = {}
    for r in rows[1:]:
        if len(r) >= 2 and r[0]:
            try:
                out[r[0].strip()] = json.loads(r[1])
            except Exception:
                pass
    return out


def save_state(st):
    try:
        ws = _state_ws()
        table = [["день", "итоги"]] + [["'" + d, json.dumps(st[d], ensure_ascii=False)]
                                       for d in sorted(st)]
        if ws.row_count < len(table):
            ws.resize(rows=len(table) + 20, cols=2)
        ws.batch_clear(["A1:B%d" % ws.row_count])
        ws.update(range_name="A1", values=table, value_input_option="USER_ENTERED")
    except Exception as e:
        log("состояние не сохранено: %s" % e)


# ---------------------------------------------------------------- прогон

def fetch_all():
    """Скачивает выгрузки и раскладывает по папкам DD.MM.YY. Возвращает число дней."""
    os.makedirs(WORK, exist_ok=True)
    seen_path = os.path.join(WORK, "_скачано.json")
    try:
        seen = json.load(io.open(seen_path, encoding="utf-8"))
    except Exception:
        seen = {}

    ids = walk(FOLDER_ID)
    log("файлов в папке: %d" % len(ids))
    best_sales, newest_stocks = {}, None
    added = 0

    for fid in ids:
        cached = seen.get(fid)
        if cached and os.path.exists(os.path.join(WORK, cached["path"])):
            blob = None
        else:
            try:
                blob = download(fid)
            except Exception as e:
                log("не скачан %s: %s" % (fid[:12], e))
                continue
            if blob[:2] != b"PK":
                seen[fid] = {"kind": "skip", "path": ""}
                continue
            kind, date = classify(blob)
            if kind is None:
                seen[fid] = {"kind": "skip", "path": ""}
                continue
            if kind == "sales":
                folder = "%s.%s.%s" % (date[8:10], date[5:7], date[2:4])
                rel = os.path.join(folder, "отчет по продажам.xlsx")
            else:
                rel = os.path.join("_остатки", fid + ".xlsx")
            os.makedirs(os.path.join(WORK, os.path.dirname(rel)), exist_ok=True)
            with open(os.path.join(WORK, rel), "wb") as f:
                f.write(blob)
            seen[fid] = {"kind": kind, "path": rel, "date": date or "",
                         "made": made_at(blob)}
            added += 1
            log("  + %s %s" % (kind, date or ""))

        info = seen[fid]
        if info["kind"] == "sales":
            d = info.get("date")
            if d and (d not in best_sales or info.get("made", "") > best_sales[d].get("made", "")):
                best_sales[d] = info
        elif info["kind"] == "stocks":
            if newest_stocks is None or info.get("path", "") > newest_stocks.get("path", ""):
                newest_stocks = info

    # самый свежий отчёт по остаткам кладём в папку последнего дня
    if best_sales and newest_stocks:
        last = max(best_sales)
        folder = "%s.%s.%s" % (last[8:10], last[5:7], last[2:4])
        dst = os.path.join(WORK, folder, "отчет по остаткам.xlsx")
        src = os.path.join(WORK, newest_stocks["path"])
        if os.path.exists(src):
            with open(src, "rb") as a, open(dst, "wb") as b:
                b.write(a.read())

    json.dump(seen, io.open(seen_path, "w", encoding="utf-8"), ensure_ascii=False)
    log("новых файлов: %d, дней с продажами: %d" % (added, len(best_sales)))
    return best_sales


def push_run_log(lines):
    """Пишет отчёт о прогоне в лист _прогоны — это единственный способ увидеть,
    что происходило в облаке: логи GitHub без токена недоступны."""
    try:
        import gspread
        import kosmos_gsheet as G
        gc = gspread.service_account(filename=G.KEY_PATH)
        sh = gc.open_by_key(G.SHEET_ID)
        try:
            ws = sh.worksheet("_прогоны")
        except Exception:
            ws = sh.add_worksheet(title="_прогоны", rows=200, cols=3)
            ws.update(range_name="A1", values=[["когда", "где", "что произошло"]])
        where = "облако" if os.environ.get("GITHUB_ACTIONS") else "компьютер"
        when = datetime.datetime.now(MSK).strftime("%d.%m %H:%M")
        ws.append_row([when, where, " | ".join(lines)[:40000]], value_input_option="RAW")
    except Exception as e:
        log("отчёт о прогоне не записан: %s" % e)


def push_cache(rows, rows_size, dates):
    """Заполняет служебный лист _данные, который читает старый скрипт в таблице.

    Он ищет выгрузки по именам файлов и из-за этого регулярно решал, что отчёта
    нет, и слал в группу ложные предупреждения. Если его кэш ведём мы, он видит
    данные на месте и помалкивает.
    """
    import gspread
    import kosmos_gsheet as G

    gc = gspread.service_account(filename=G.KEY_PATH)
    sh = gc.open_by_key(G.SHEET_ID)
    try:
        ws = sh.worksheet("_данные")
    except Exception:
        ws = sh.add_worksheet(title="_данные", rows=len(dates) + 10, cols=5)

    table = [["дата", "файл", "изменён", "данные", "размеры"]]
    for d in dates:
        day = {}
        for r in rows:
            if not (r["o"].get(d) or r["b"].get(d) or r["stock"].get(d)):
                continue
            day[r["art"]] = {"nm": r["nm"], "brand": r["brand"], "subj": r["subj"],
                             "o": r["o"].get(d, 0), "orub": r["orub"].get(d, 0),
                             "b": r["b"].get(d, 0), "pay": r["pay"].get(d, 0),
                             "stock": r["stock"].get(d, 0), "fbs": r["fbs"].get(d, 0)}
        size = {}
        for r in rows_size:
            v = [r["o"].get(d, 0), r["orub"].get(d, 0), r["b"].get(d, 0),
                 r["pay"].get(d, 0), r["stock"].get(d, 0)]
            if any(v):
                size[r["art"]] = v
        table.append(["'" + d, "ведёт kosmos_daily", 0,
                      json.dumps(day, ensure_ascii=False),
                      json.dumps(size, ensure_ascii=False)])

    if ws.row_count < len(table):
        ws.resize(rows=len(table) + 10, cols=5)
    ws.batch_clear(["A1:E%d" % ws.row_count])
    ws.update(range_name="A1", values=table, value_input_option="USER_ENTERED")
    log("служебный лист обновлён: %d дней" % (len(table) - 1))


def main():
    force = "--force" in sys.argv
    trace = []
    best = fetch_all()
    trace.append("файлов-дней найдено: %d" % len(best))

    build_local.ROOT = WORK
    agg, dates, transit, missing, agg_size = build_local.collect_local()
    if not dates:
        trace.append("разобранных дней нет")
        push_run_log(trace)
        tg_send("⚠️ <b>КОСМОС</b>: в папке нет ни одной выгрузки продаж.")
        return
    trace.append("дней в сборке: %d, последний %s" % (len(dates), dates[-1]))
    rows = sorted(agg.values(), key=lambda r: r["art"])
    rows_size = sorted(agg_size.values(), key=lambda r: r["art"])
    target = dates[-1]

    cur = {k: round(sum(r[k].get(target, 0) for r in rows))
           for k in ("o", "orub", "b", "pay")}
    st = load_state()
    if not force and st.get(target) == cur:
        trace.append("данные за %s не изменились — выход" % target)
        push_run_log(trace)
        log("данные за %s не изменились — ничего не делаю" % target)
        return

    with open(OUT, "wb") as f:
        f.write(B.build_xlsx(rows, dates, transit, rows_size))
    log("собран %s: дней %d, артикулов %d, строк по размерам %d"
        % (OUT, len(dates), len(rows), len(rows_size)))

    import kosmos_gsheet
    sys.argv = ["kosmos_gsheet.py", OUT]
    kosmos_gsheet.main()

    try:
        push_cache(rows, rows_size, dates)
    except Exception as e:
        log("служебный лист не обновлён: %s" % e)

    prev_date = dates[-2] if len(dates) > 1 else None
    prev = ({k: round(sum(r[k].get(prev_date, 0) for r in rows))
             for k in ("o", "orub", "b", "pay")} if prev_date else None)
    lines = []
    for label, key, unit in (("ЗАКАЗЫ", "o", "шт"), ("СУММА ЗАКАЗОВ", "orub", "₽"),
                             ("ПРОДАЖИ", "b", "шт"), ("СУММА ПРОДАЖ", "pay", "₽")):
        val = "%s %s" % (B.fmt(cur[key]), unit)
        lines.append("%-15s%-14s%s" % (label, val, B.delta(cur[key], prev[key]) if prev else ""))
    corrected = target in st
    dd = "%s.%s.%s" % (target[8:10], target[5:7], target[:4])
    text = ("<b>КОСМОС</b> — сводка за %s%s\n\n<pre>%s</pre>\n\n"
            "<i>Собрано дней: %d (%s–%s), артикулов %d.</i>\n\nДашборд:\n%s"
            % (dd, " (исправленная)" if corrected else "", "\n".join(lines),
               len(dates), dates[0][8:10] + "." + dates[0][5:7], dd, len(rows), SHEET_URL))
    ok = tg_send(text)
    trace.append("сводка за %s отправлена: %s%s"
                 % (target, ok, "" if ok else " — " + "; ".join(LAST_TG)))
    push_run_log(trace)
    if ok:
        st[target] = cur
        save_state(st)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        traceback.print_exc()
        log("СБОЙ: %s: %s" % (type(e).__name__, e))
        tg_send("⚠️ <b>Дашборд КОСМОС не собрался</b>\n\n<pre>%s: %s</pre>"
                % (type(e).__name__, str(e)[:400]))
        sys.exit(1)
