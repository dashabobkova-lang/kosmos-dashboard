#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Локальная сборка дашборда КОСМОС из папок на диске Z: (Bitrix WebDAV).

Использует логику build_kosmos_dashboard.py, но читает файлы с диска,
а не через Bitrix REST. Результат кладёт в указанный путь.

  python build_local.py <выходной .xlsx>
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import build_kosmos_dashboard as B

ROOT = r"Z:\ДАШБОРДЫ И ОТЧЕТЫ\ежедневные отчеты КОСМОС"
ART_PREFIX = B.ART_PREFIX


def collect_local():
    folders = []
    for name in os.listdir(ROOT):
        p = os.path.join(ROOT, name)
        if not os.path.isdir(p):
            continue
        m = re.match(r"^(\d{2})\.(\d{2})\.(\d{2})$", name.strip())
        if not m:
            B.log("пропускаю папку: %s" % name)
            continue
        d, mo, y = m.groups()
        folders.append(("20%s-%s-%s" % (y, mo, d), p, name))
    folders.sort()

    agg, agg_size, dates, missing = {}, {}, [], []
    stock_src = None

    for date, path, fname in folders:
        kids = os.listdir(path)
        sales = [k for k in kids if re.search(r"по\s*продажам", k, re.I)]
        stocks = [k for k in kids if re.search(r"по\s*остатк", k, re.I)]
        if stocks:
            with open(os.path.join(path, stocks[0]), "rb") as f:
                stock_src = (date, f.read())
        if not sales:
            missing.append(fname)
            continue
        with open(os.path.join(path, sales[0]), "rb") as f:
            blob = f.read()
        rows = B.read_sheet_rows(blob)
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
            a["o"][date] = a["o"].get(date, 0) + B.num(c.get("L"))
            a["orub"][date] = a["orub"].get(date, 0) + B.num(c.get("M"))
            a["b"][date] = a["b"].get(date, 0) + B.num(c.get("N"))
            a["pay"][date] = a["pay"].get(date, 0) + B.num(c.get("O"))
            a["stock"][date] = a["stock"].get(date, 0) + B.num(c.get("P"))
            if "Свой" in (c.get("K") or ""):
                a["fbs"][date] = a["fbs"].get(date, 0) + B.num(c.get("P"))

            # тот же разрез, но по размерам: колонка I
            size = (c.get("I") or "").strip() or "б/р"
            key = "%s · %s" % (art, size)
            z = agg_size.setdefault(key, {"art": key, "nm": c.get("G", ""),
                                          "brand": c.get("A", ""), "subj": c.get("B", ""),
                                          "o": {}, "orub": {}, "b": {}, "pay": {},
                                          "stock": {}, "fbs": {}})
            z["o"][date] = z["o"].get(date, 0) + B.num(c.get("L"))
            z["orub"][date] = z["orub"].get(date, 0) + B.num(c.get("M"))
            z["b"][date] = z["b"].get(date, 0) + B.num(c.get("N"))
            z["pay"][date] = z["pay"].get(date, 0) + B.num(c.get("O"))
            z["stock"][date] = z["stock"].get(date, 0) + B.num(c.get("P"))
            if "Свой" in (c.get("K") or ""):
                z["fbs"][date] = z["fbs"].get(date, 0) + B.num(c.get("P"))
            cnt += 1
        dates.append(date)
        B.log("  %s: строк ДАБ %d, заказано %d шт" %
              (fname, cnt, round(sum(a2["o"].get(date, 0) for a2 in agg.values()))))

    if missing:
        B.log("нет отчёта по продажам в: %s" % ", ".join(missing))

    transit = {}
    if stock_src:
        sd, sblob = stock_src
        for rn, c in B.read_sheet_rows(sblob):
            if rn < 2:
                continue
            art = (c.get("C") or "").strip()
            if not art.startswith(ART_PREFIX):
                continue
            t = transit.setdefault(art, {"to": 0.0, "from": 0.0})
            t["to"] += B.num(c.get("H"))
            t["from"] += B.num(c.get("I"))
        B.log("остатки в пути: отчёт за %s, артикулов %d" % (sd, len(transit)))

    return agg, sorted(set(dates)), transit, missing, agg_size


def main():
    out = sys.argv[1]
    agg, dates, transit, missing, agg_size = collect_local()
    rows = sorted(agg.values(), key=lambda r: r["art"])
    B.log("итог: дней %d (%s..%s), артикулов %d" % (len(dates), dates[0], dates[-1], len(rows)))
    blob = B.build_xlsx(rows, dates, transit)
    with open(out, "wb") as f:
        f.write(blob)
    B.log("записан %s (%d байт)" % (out, len(blob)))

    # сводка в консоль
    d = dates[-1]
    prev = dates[-2]
    for label, key, unit in (("ЗАКАЗЫ", "o", "шт"), ("СУММА ЗАКАЗОВ", "orub", "руб"),
                             ("ПРОДАЖИ", "b", "шт"), ("СУММА ПРОДАЖ", "pay", "руб")):
        cur = sum(r[key].get(d, 0) for r in rows)
        old = sum(r[key].get(prev, 0) for r in rows)
        B.log("%-16s %12s   (пред. день %s, %s)" %
              (label, B.fmt(cur) + " " + unit, B.fmt(old), B.delta(cur, old)))


if __name__ == "__main__":
    main()
