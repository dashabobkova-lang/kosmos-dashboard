# -*- coding: utf-8 -*-
"""Диагностика строки BITRIX_WEBHOOK без раскрытия её значения."""
import os

import env_local
env_local.load()

w = os.environ.get("BITRIX_WEBHOOK", "")
print("длина: %d" % len(w))
print("начинается с https://: %s" % w.startswith("https://"))
print("домен: %s" % (w.split("/")[2] if w.count("/") >= 2 else "—"))
parts = w.split("/")
print("сегментов пути: %d" % max(0, len(parts) - 3))
print("схема пути: /%s" % "/".join(
    ("<%d симв.>" % len(p)) if i >= 1 else p for i, p in enumerate(parts[3:]) if p))
bad = [(i, ch, hex(ord(ch))) for i, ch in enumerate(w) if ord(ch) > 127]
print("не-ASCII символов: %d" % len(bad))
for i, ch, code in bad[:10]:
    print("   позиция %d: %r (%s)" % (i, ch, code))
ws = [(i, hex(ord(ch))) for i, ch in enumerate(w) if ch.isspace()]
print("пробельных символов: %d %s" % (len(ws), ws[:5]))
