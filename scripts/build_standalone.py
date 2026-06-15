#!/usr/bin/env python3
"""
prices/data/latest.json を scripts/dashboard_template.html に埋め込み、
外部リクエストゼロの自己完結ダッシュボード prices/index.html を生成する。

社内/非公開利用向け：生成物は CDN・外部API・解析タグを一切含まず、
ダブルクリックでオフライン表示できる単一HTMLになる。
fetch_prices.py の直後（毎朝06:00 JST）に実行される。
"""
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATE = os.path.join(ROOT, "scripts", "dashboard_template.html")
LATEST = os.path.join(ROOT, "prices", "data", "latest.json")
OUT = os.path.join(ROOT, "prices", "index.html")


def main():
    with open(TEMPLATE, encoding="utf-8") as f:
        tpl = f.read()
    with open(LATEST, encoding="utf-8") as f:
        data = f.read().strip()

    if "__DATA__" not in tpl:
        raise SystemExit("テンプレートに __DATA__ プレースホルダがありません")
    json.loads(data)  # 妥当性チェック

    html = tpl.replace("__DATA__", data)
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"[ok] 自己完結HTMLを生成 → {OUT} ({len(html):,} bytes)")


if __name__ == "__main__":
    main()
