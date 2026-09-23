# -*- coding: utf-8 -*-
"""Подхватывает BITRIX_WEBHOOK / TG_TOKEN / TG_CHAT из локального .env.

Значения кладутся только в os.environ и на экран не выводятся.
Нужен для ручных прогонов на машине Дарьи; в GitHub Actions не используется.
"""
import io
import os

ENV = os.path.join(r"C:\Users\DASHA\Desktop\Claude", ".env")
MAP = {"BITRIX_WEBHOOK": "BITRIX_WEBHOOK",
       "TELEGRAM_BOT_TOKEN": "TG_TOKEN",
       "TELEGRAM_CHAT_ID": "TG_CHAT"}


def load():
    found = []
    try:
        for line in io.open(ENV, encoding="utf-8-sig"):
            if "=" not in line:
                continue
            k, v = line.split("=", 1)
            k, v = k.strip(), v.strip()
            if k in MAP and v:
                os.environ[MAP[k]] = v
                found.append(MAP[k])
    except Exception as e:
        print("не прочитан .env: %s" % e)
    return found


if __name__ == "__main__":
    print("подхвачено:", ", ".join(load()) or "ничего")
