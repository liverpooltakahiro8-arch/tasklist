#!/usr/bin/env python3
"""
船舶燃料 価格取得スクリプト（無料ソースのみ・標準ライブラリのみ）

無料で取得できる市場指標（Yahoo Finance / Stooq、任意で EIA）を組み合わせ、
各船舶燃料の価格目安を透明な換算式で算出して prices/data/latest.json を生成する。
GitHub Actions から毎朝 06:00 JST に実行される。

出所・換算ロジックの詳細は docs/DATA_SOURCES.md を参照。
"""

import csv
import io
import json
import os
import sys
import time
import urllib.request
import urllib.error
from datetime import datetime, timezone, timedelta

JST = timezone(timedelta(hours=9))
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT, "prices", "data")
LATEST = os.path.join(DATA_DIR, "latest.json")
HISTORY = os.path.join(DATA_DIR, "history.csv")

# ---------------------------------------------------------------------------
# 設定（係数はここに集約。docs/DATA_SOURCES.md と一致させること）
# ---------------------------------------------------------------------------
CONFIG = {
    # 換算係数（市場指標 → 船舶燃料 $/t, $/kg）
    "vlsfo": {"brent_mult": 7.33, "premium": 70},
    "hsfo": {"vlsfo_discount": 95},
    "mgo": {"ho_gal_per_t": 312, "premium": 40},
    "biodiesel": {"lb_per_t": 2204.62, "processing": 120},
    "ethanol": {"gal_per_t": 335},
    "ammonia_model": {"ng_mult": 34, "opex": 180},
    "h2_grey": {"ng_mult": 0.18, "opex": 0.9},     # $/kg
    "h2_green": {"power_usd_kwh": 0.06, "kwh_per_kg": 52, "opex": 1.0},  # $/kg

    # 月次ポステッド／参考値（手動更新。Methanex・World Bank等を反映）
    "methanol_posted_usd_t": 740,        # Methanex アジア契約価格（2026-04時点 参考）
    "methanol_as_of": "2026-04",
    "ammonia_ref_usd_t": 420,            # アンモニア（CFR Far East 参考）
    "ammonia_as_of": "2026-04",

    # 低位発熱量 LHV (MJ/kg) — $/GJ 換算に使用
    "lhv": {
        "mgo": 42.7, "vlsfo": 40.5, "hsfo": 40.5, "biodiesel": 37.2,
        "ethanol": 26.8, "methanol": 19.9, "ammonia": 18.6, "hydrogen": 120.0,
    },
}

# Yahoo Finance のシンボル（キー不要）
YF = {
    "brent": "BZ=F", "wti": "CL=F", "natgas": "NG=F",
    "heatoil": "HO=F", "rbob": "RB=F", "soyoil": "ZL=F", "ethanol": "EH=F",
}
# Stooq フォールバック用シンボル
STOOQ = {
    "brent": "cb.f", "wti": "cl.f", "natgas": "ng.f",
    "heatoil": "ho.f", "rbob": "rb.f", "soyoil": "zl.f",
}

UA = {"User-Agent": "Mozilla/5.0 (compatible; FuelPriceBot/1.0; +https://github.com)"}


def http_get(url, timeout=20):
    req = urllib.request.Request(url, headers=UA)
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return r.read().decode("utf-8", "replace")


def fetch_yahoo(symbol):
    """Yahoo Finance Chart API から最新終値と前日比を取得。"""
    url = (f"https://query1.finance.yahoo.com/v8/finance/chart/{symbol}"
           f"?interval=1d&range=7d")
    data = json.loads(http_get(url))
    res = data["chart"]["result"][0]
    meta = res["meta"]
    price = meta.get("regularMarketPrice")
    prev = meta.get("chartPreviousClose") or meta.get("previousClose")
    if price is None:
        # フォールバック：close 配列の最後の有効値
        closes = [c for c in res["indicators"]["quote"][0]["close"] if c is not None]
        price = closes[-1]
        prev = closes[-2] if len(closes) > 1 else price
    return float(price), (float(prev) if prev else None)


def fetch_stooq(symbol):
    """Stooq CSV から終値を取得（フォールバック）。"""
    url = f"https://stooq.com/q/l/?s={symbol}&f=sd2t2ohlcv&h&e=csv"
    txt = http_get(url)
    row = list(csv.DictReader(io.StringIO(txt)))[0]
    close = row.get("Close")
    if close in (None, "", "N/D"):
        raise ValueError("stooq no close")
    return float(close), None


def get_market(name):
    """市場指標を Yahoo→Stooq の順で取得。(value, prev, source) を返す。"""
    errors = []
    if name in YF:
        try:
            v, p = fetch_yahoo(YF[name])
            return v, p, "Yahoo Finance"
        except Exception as e:  # noqa: BLE001
            errors.append(f"yahoo:{e}")
    if name in STOOQ:
        try:
            v, p = fetch_stooq(STOOQ[name])
            return v, p, "Stooq"
        except Exception as e:  # noqa: BLE001
            errors.append(f"stooq:{e}")
    print(f"[warn] {name} 取得失敗: {'; '.join(errors)}", file=sys.stderr)
    return None, None, None


def pct(cur, prev):
    if cur is None or prev in (None, 0):
        return None
    return round((cur - prev) / prev * 100, 2)


def per_gj(usd_per_t, lhv_mj_kg):
    """$/t と LHV(MJ/kg) から $/GJ を算出。1 t = 1000 kg, LHV*1000kg / 1000 = GJ/t."""
    if usd_per_t is None:
        return None
    gj_per_t = lhv_mj_kg * 1000 / 1000  # MJ/kg * 1000 kg = MJ/t /1000 = GJ/t
    return round(usd_per_t / gj_per_t, 2)


def main():
    now = datetime.now(JST)
    c = CONFIG
    lhv = c["lhv"]

    # --- 市場指標の取得 -----------------------------------------------------
    mk = {}
    for name in ["brent", "wti", "natgas", "heatoil", "soyoil", "ethanol"]:
        v, p, src = get_market(name)
        mk[name] = {"value": v, "prev": p, "source": src}

    fuels = []

    def add(key, name_ja, unit, price, change, source, basis, kind, freq,
            extra=None, lhv_key=None):
        lk = lhv_key or key
        gj = per_gj(price, lhv[lk]) if price is not None and lk in lhv else None
        item = {
            "key": key, "name": name_ja, "unit": unit,
            "price": (round(price, 2) if price is not None else None),
            "usd_per_gj": gj, "change_pct": change,
            "source": source, "basis": basis,
            "kind": kind, "freq": freq,
            "available": price is not None,
        }
        if extra:
            item.update(extra)
        fuels.append(item)

    # VLSFO（Brent連動）
    brent = mk["brent"]["value"]
    vlsfo = None
    if brent is not None:
        vlsfo = brent * c["vlsfo"]["brent_mult"] + c["vlsfo"]["premium"]
    add("vlsfo", "VLSFO（低硫黄重油）", "USD/t", vlsfo,
        pct(vlsfo, (mk["brent"]["prev"] * c["vlsfo"]["brent_mult"] + c["vlsfo"]["premium"])
            if mk["brent"]["prev"] else None),
        mk["brent"]["source"], "Brent×7.33+70", "Proxy", "Daily")

    # HSFO / IFO380
    hsfo = (vlsfo - c["hsfo"]["vlsfo_discount"]) if vlsfo is not None else None
    add("hsfo", "HSFO / IFO380（高硫黄重油）", "USD/t", hsfo,
        fuels[-1]["change_pct"], mk["brent"]["source"], "VLSFO−95", "Proxy", "Daily")

    # MGO（ガスオイル連動）
    ho = mk["heatoil"]["value"]
    mgo = (ho * c["mgo"]["ho_gal_per_t"] + c["mgo"]["premium"]) if ho is not None else None
    mgo_prev = (mk["heatoil"]["prev"] * c["mgo"]["ho_gal_per_t"] + c["mgo"]["premium"]) \
        if mk["heatoil"]["prev"] else None
    add("mgo", "MGO（舶用軽油）", "USD/t", mgo, pct(mgo, mgo_prev),
        mk["heatoil"]["source"], "HeatingOil×312+40", "Proxy", "Daily")

    # バイオディーゼル（大豆油連動, ¢/lb）
    soy = mk["soyoil"]["value"]
    bio = ((soy / 100) * c["biodiesel"]["lb_per_t"] + c["biodiesel"]["processing"]) \
        if soy is not None else None
    bio_prev = ((mk["soyoil"]["prev"] / 100) * c["biodiesel"]["lb_per_t"]
                + c["biodiesel"]["processing"]) if mk["soyoil"]["prev"] else None
    add("biodiesel", "バイオディーゼル（B100/FAME）", "USD/t", bio, pct(bio, bio_prev),
        mk["soyoil"]["source"], "大豆油×2204.6+120", "Proxy", "Daily")

    # エタノール（先物 $/gal）
    eth = mk["ethanol"]["value"]
    eth_t = (eth * c["ethanol"]["gal_per_t"]) if eth is not None else None
    eth_prev = (mk["ethanol"]["prev"] * c["ethanol"]["gal_per_t"]) \
        if mk["ethanol"]["prev"] else None
    add("ethanol", "エタノール", "USD/t", eth_t, pct(eth_t, eth_prev),
        mk["ethanol"]["source"] or "—",
        "Ethanol先物×335", "Direct" if mk["ethanol"]["source"] else "—", "Daily")

    # メタノール（月次ポステッド参考）
    add("methanol", "メタノール", "USD/t", float(c["methanol_posted_usd_t"]), None,
        "Methanex ポステッド価格", f"アジア契約価格 ({c['methanol_as_of']})",
        "Reference", "Monthly")

    # アンモニア（月次参考値）
    add("ammonia", "アンモニア（参考値）", "USD/t", float(c["ammonia_ref_usd_t"]), None,
        "World Bank / 公開コメンタリ", f"CFR Far East 参考 ({c['ammonia_as_of']})",
        "Reference", "Monthly")

    # アンモニア（ガス連動 生産コスト推定）
    ng = mk["natgas"]["value"]
    nh3_model = (ng * c["ammonia_model"]["ng_mult"] + c["ammonia_model"]["opex"]) \
        if ng is not None else None
    nh3_prev = (mk["natgas"]["prev"] * c["ammonia_model"]["ng_mult"]
                + c["ammonia_model"]["opex"]) if mk["natgas"]["prev"] else None
    add("ammonia_model", "アンモニア（ガス連動 生産コスト推定）", "USD/t",
        nh3_model, pct(nh3_model, nh3_prev), mk["natgas"]["source"] or "—",
        "HenryHub×34+180", "Model", "Daily", lhv_key="ammonia")

    # 水素 グレー（$/kg, ガス連動）
    h2g = (ng * c["h2_grey"]["ng_mult"] + c["h2_grey"]["opex"]) if ng is not None else None
    h2g_prev = (mk["natgas"]["prev"] * c["h2_grey"]["ng_mult"] + c["h2_grey"]["opex"]) \
        if mk["natgas"]["prev"] else None
    # $/kg → $/t = ×1000、$/GJ は hydrogen LHV
    add("hydrogen_grey", "水素 グレー（化石由来 推定）", "USD/kg", h2g,
        pct(h2g, h2g_prev), mk["natgas"]["source"] or "—",
        "HenryHub×0.18+0.9", "Model", "Daily",
        extra={"usd_per_t": (round(h2g * 1000, 1) if h2g else None),
               "usd_per_gj": (round(h2g * 1000 / lhv["hydrogen"], 2) if h2g else None)})

    # 水素 グリーン（$/kg, 電力連動）
    h2gr = (c["h2_green"]["power_usd_kwh"] * c["h2_green"]["kwh_per_kg"]
            + c["h2_green"]["opex"])
    add("hydrogen_green", "水素 グリーン（再エネ由来 推定）", "USD/kg", h2gr, None,
        f"電力単価 ${c['h2_green']['power_usd_kwh']}/kWh",
        "電力×52+1.0", "Model", "Daily",
        extra={"usd_per_t": round(h2gr * 1000, 1),
               "usd_per_gj": round(h2gr * 1000 / lhv["hydrogen"], 2)})

    # --- 出力組み立て -------------------------------------------------------
    ok = sum(1 for f in fuels if f["available"])
    payload = {
        "updated_at_jst": now.strftime("%Y-%m-%d %H:%M JST"),
        "updated_at_iso": now.isoformat(),
        "next_update_jst": (now.replace(hour=6, minute=0, second=0, microsecond=0)
                            + timedelta(days=1)).strftime("%Y-%m-%d 06:00 JST"),
        "available_count": ok,
        "total_count": len(fuels),
        "markets": {k: {"value": (round(v["value"], 4) if v["value"] is not None else None),
                        "source": v["source"]} for k, v in mk.items()},
        "fuels": fuels,
        "config": {
            "coefficients": {k: v for k, v in c.items() if k != "lhv"},
            "lhv_mj_kg": lhv,
        },
        "disclaimer": ("価格は無料市場指標に基づく参考目安。実取引・契約価格"
                       "（Platts/Argus/Ship & Bunker等）とは異なる。"),
    }

    # 全滅した場合は前回値を保持（stale）
    if ok == 0 and os.path.exists(LATEST):
        with open(LATEST, encoding="utf-8") as f:
            old = json.load(f)
        old["stale"] = True
        old["stale_checked_at_jst"] = now.strftime("%Y-%m-%d %H:%M JST")
        payload = old
        print("[warn] 全ソース取得失敗。前回値を stale として保持。", file=sys.stderr)

    os.makedirs(DATA_DIR, exist_ok=True)
    with open(LATEST, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)

    # 履歴 CSV へ追記
    write_header = not os.path.exists(HISTORY)
    with open(HISTORY, "a", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        if write_header:
            w.writerow(["datetime_jst"] + [fu["key"] for fu in fuels])
        w.writerow([now.strftime("%Y-%m-%d %H:%M")]
                   + [(fu["price"] if fu["available"] else "") for fu in fuels])

    print(f"[ok] {ok}/{len(fuels)} 燃料を更新 → {LATEST}")


if __name__ == "__main__":
    main()
