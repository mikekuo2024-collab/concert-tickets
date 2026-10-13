#!/usr/bin/env python3
"""從 index.html 產生各老師的團購頁面：<slug>/index.html

用法：在儲存庫根目錄執行  python3 tools/build_teachers.py
修改 index.html 後務必重新執行，三個老師頁面才會同步更新。
"""
import json
import os

# sessions：此老師表單開放的場次 id（1=台中 11/28，2=高雄 12/26 午場，3=高雄 12/26 晚場）
# deadline：結單日（含當日，台灣時間），須與 apps-script/Code.gs 的 TEACHERS 設定一致
TEACHERS = [
    {"id": "wang", "name": "王裕文", "sessions": [1], "deadline": "2026-11-20"},
    {"id": "wei", "name": "魏鴻達", "sessions": [1], "deadline": "2026-11-20"},
    {"id": "hsieh", "name": "謝政良", "sessions": [2, 3], "deadline": "2026-12-20"},
]

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src = open(os.path.join(ROOT, "index.html"), encoding="utf-8").read()

MARK = "const TEACHER = null;"
assert src.count(MARK) == 1, "index.html 找不到 TEACHER 設定"
assert src.count('src="logo.png"') == 1, "index.html 找不到 logo"

for t in TEACHERS:
    page = src.replace(MARK, "const TEACHER = " + json.dumps(t, ensure_ascii=False) + ";")
    page = page.replace('src="logo.png"', 'src="../logo.png"')
    out_dir = os.path.join(ROOT, t["id"])
    os.makedirs(out_dir, exist_ok=True)
    with open(os.path.join(out_dir, "index.html"), "w", encoding="utf-8") as f:
        f.write(page)
    print("產生", t["id"] + "/index.html", "→", t["name"] + "老師")
