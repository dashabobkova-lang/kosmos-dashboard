#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Ежедневный прогон дашборда КОСМОС в GitHub Actions.

Логика одного запуска (идемпотентна — можно запускать сколько угодно раз):
  1. target = вчерашняя дата по Москве
  2. тянем данные из Bitrix; если последний день в данных != target -> выход
  3. сравниваем итоги за target с state.json
       нет записи        -> собрать, залить, отправить сводку
       запись отличается -> собрать, залить, отправить ИСПРАВЛЕННУЮ сводку
       совпадает         -> ничего не делать
  4. state.json коммитится обратно в репозиторий шагом workflow

Переменные окружения: BITRIX_WEBHOOK, TG_TOKEN, TG_CHAT
"""
import os
import sys
import json
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_kosmos_dashboard as B

STATE = os.path.join(HERE, "state.json")
MSK = datetime.timezone(datetime.timedelta(hours=3))
KEYS = ("o", "orub", "b", "pay")


def load_state():
    try:
        with open(STATE, encoding="utf-8-sig") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except Exception as e:
        B.log("state.json повреждён (%s) — выхожу, чтобы не слать дубль сводки" % e)
        sys.exit(1)


def save_state(st):
    with open(STATE, "w", encoding="utf-8") as f:
        json.dump(st, f, ensure_ascii=False, indent=1, sort_keys=True)


def totals(rows, d):
    return {k: round(sum(r[k].get(d, 0) for r in rows)) for k in KEYS}


def summary_text(cur, prev, d, corrected, missing):
    lines = []
    for label, key, unit in (("ЗАКАЗЫ", "o", "шт"), ("СУММА ЗАКАЗОВ", "orub", "₽"),
                             ("ПРОДАЖИ", "b", "шт"), ("СУММА ПРОДАЖ", "pay", "₽")):
        val = "%s %s" % (B.fmt(cur[key]), unit)
        dl = B.delta(cur[key], prev[key]) if prev else ""
        lines.append("%-15s%-14s%s" % (label, val, dl))
    dd = "%s.%s.%s" % (d[8:10], d[5:7], d[:4])
    head = "<b>КОСМОС</b> — сводка за %s%s" % (dd, " (исправленная)" if corrected else "")
    tail = "\n\nОтчёт был перевыгружен, цифры уточнены." if corrected else ""
    text = "%s\n\n<pre>%s</pre>%s\n\nДашборд:\n%s" % (head, "\n".join(lines), tail, B.DASH_URL)
    if missing:
        text += "\n\n⚠️ Нет отчёта по продажам за: %s" % ", ".join(missing)
    return text


def main():
    if not B.WEBHOOK:
        B.die("не задана переменная BITRIX_WEBHOOK")

    target = (datetime.datetime.now(MSK).date() - datetime.timedelta(days=1)).isoformat()
    B.log("цель: %s" % target)

    agg, dates, transit, missing = B.collect()
    rows = sorted(agg.values(), key=lambda r: r["art"])

    # ранний прогон (9:00 МСК) - только забор и проверка данных, без публикации.
    # Если Битрикс недоступен, узнаем об этом за полчаса до публикации.
    now = datetime.datetime.now(MSK)
    if "--probe" in sys.argv or (now.hour == 9 and now.minute < 25):
        if dates and dates[-1] == target:
            B.log("забор: отчёт за %s на месте, дней %d, артикулов %d — публикация в 9:30"
                  % (target, len(dates), len(rows)))
        else:
            B.log("забор: отчёта за %s ещё нет (последний день %s)"
                  % (target, dates[-1] if dates else "—"))
        return

    if not dates or dates[-1] != target:
        B.log("последний день в данных %s != %s — отчёта ещё нет"
              % (dates[-1] if dates else "—", target))
        # поздний прогон: данных так и не появилось - предупреждаем, один раз за день
        st = load_state()
        key = "missing:" + target
        if datetime.datetime.now(MSK).hour >= 11 and key not in st:
            dd = "%s.%s.%s" % (target[8:10], target[5:7], target[:4])
            folder = "%s.%s.%s" % (target[8:10], target[5:7], target[2:4])
            msg = ("⚠️ <b>КОСМОС: нет отчёта за %s</b>\n\n"
                   "В папке «ежедневные отчеты КОСМОС/%s» не появился «отчет по продажам» — "
                   "дашборд и сводку собрать не из чего.\n"
                   "Продолжаю проверять папку до вечера: как только файл появится, "
                   "дашборд соберётся и сводка придёт сама." % (dd, folder))
            if B.tg_send(msg):
                st[key] = datetime.datetime.now(MSK).isoformat(timespec="seconds")
                save_state(st)
        return

    # папку за сегодня, которая ещё пуста, в предупреждения не тащим
    today_folder = "%s.%s.%s" % tuple(datetime.datetime.now(MSK).strftime("%d.%m.%y").split("."))
    missing = [m for m in missing if m != today_folder]

    cur = totals(rows, target)
    prev = totals(rows, dates[-2]) if len(dates) > 1 else None

    st = load_state()
    rec = st.get(target)
    if rec and rec.get("totals") == cur:
        B.log("данные за %s не изменились — ничего не делаю" % target)
        return

    blob = B.build_xlsx(rows, dates, transit)
    B.log("xlsx собран: %d дней, %d артикулов, %d байт" % (len(dates), len(rows), len(blob)))
    res = B.upload_file(B.TARGET_FOLDER_ID, B.TARGET_NAME, blob)
    fid = (res.get("result") or {}).get("ID")
    if not fid:
        B.die("не удалось залить файл: %s" % json.dumps(res, ensure_ascii=False)[:300])
    B.log("залит в Bitrix, file id=%s" % fid)

    corrected = bool(rec)
    if corrected:
        B.log("РАСХОЖДЕНИЕ: было %s, стало %s" % (rec.get("totals"), cur))
    B.tg_send(summary_text(cur, prev, target, corrected, missing))
    st[target] = {"totals": cur,
                  "sent": datetime.datetime.now(MSK).isoformat(timespec="seconds")}
    save_state(st)
    B.log("готово")


if __name__ == "__main__":
    try:
        main()
    except SystemExit:
        raise
    except Exception as e:
        import traceback
        traceback.print_exc()
        B.die("%s: %s" % (type(e).__name__, e))
