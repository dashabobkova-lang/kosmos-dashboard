# -*- coding: utf-8 -*-
"""Показывает, что лежит в папках ежедневных отчётов КОСМОС в Битриксе.

  python check.py [сколько последних папок]
"""
import re
import sys

import env_local
env_local.load()

import build_kosmos_dashboard as B

n = int(sys.argv[1]) if len(sys.argv) > 1 else 4

folders = []
for it in B.get_children(B.DAILY_FOLDER_ID):
    if it.get("TYPE") == "folder" and re.match(r"^\d{2}\.\d{2}\.\d{2}$", it["NAME"].strip()):
        d, mo, y = it["NAME"].split(".")
        folders.append(("20%s-%s-%s" % (y, mo, d), it["ID"], it["NAME"]))
folders.sort()

for date, fid, name in folders[-n:]:
    kids = [k for k in B.get_children(fid) if k.get("TYPE") == "file"]
    print("%s — файлов %d" % (name, len(kids)))
    for k in kids:
        mark = "  <-- отчёт по продажам" if re.search(r"по\s*продажам", k["NAME"], re.I) else ""
        print("    %-52s %s%s" % (k["NAME"], k.get("UPDATE_TIME", "")[:19], mark))
    if not kids:
        print("    (пусто)")
