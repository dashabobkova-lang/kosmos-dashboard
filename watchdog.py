#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Сторож: убеждается, что дневной прогон вообще состоялся.

Запускается позже основного. Смотрит state.json: за вчерашний день там должна быть
либо запись с итогами (сводка ушла), либо пометка "missing:" (данных не было и об этом
предупредили). Если нет ни того, ни другого — значит основной прогон не отработал
и об этом никто не узнал. Тогда шлём алерт.
"""
import os
import sys
import json
import datetime

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import build_kosmos_dashboard as B
from daily import load_state, MSK

RUN_URL = "https://github.com/dashabobkova-lang/kosmos-dashboard/actions"


def main():
    target = (datetime.datetime.now(MSK).date() - datetime.timedelta(days=1)).isoformat()
    st = load_state()
    if target in st:
        B.log("за %s сводка отправлена — всё в порядке" % target)
        return
    if ("missing:" + target) in st:
        B.log("за %s предупреждение об отсутствии отчёта отправлено — всё в порядке" % target)
        return
    dd = "%s.%s.%s" % (target[8:10], target[5:7], target[:4])
    B.log("ни сводки, ни предупреждения за %s — шлю алерт" % target)
    B.tg_send("⚠️ <b>КОСМОС: утренний прогон не состоялся</b>\n\n"
              "За %s не отправлено ни сводки, ни предупреждения — похоже, "
              "задача не запустилась или упала.\n%s" % (dd, RUN_URL))


if __name__ == "__main__":
    main()
