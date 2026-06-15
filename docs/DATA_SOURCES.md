# 船舶燃料 価格データソース整理（無料・タイムリー取得可能なもの）

本サイト（`prices/index.html`）が表示する各燃料価格の **出所・更新頻度・取得方法・換算ロジック** をまとめる。
船舶燃料には「単一の無料公式API」が存在しないため、**無料で取得できる市場指標を組み合わせ、透明な換算式で各燃料の価格目安を算出**している。
各値には必ず以下のラベルを付与する。

- **取得区分**：`直接 (Direct)` / `連動推定 (Proxy)` / `モデル推定 (Model)`
- **更新頻度**：`日次 (Daily)` / `月次 (Monthly)` / `参考 (Reference)`

> ⚠️ 重要：本サイトの価格は**意思決定の参考目安**であり、実際の入札・契約価格（Platts / Argus / Ship & Bunker 実取引値）とは異なる。
> 商用の確報値が必要な場合は有料サービス（OPIS / Platts / Argus / Ship & Bunker Premium）を参照すること。

---

## 1. 無料で利用できる一次ソース一覧

| ソース | 内容 | 形式 | APIキー | 頻度 | 用途 |
|---|---|---|---|---|---|
| **Yahoo Finance Chart API** `query1.finance.yahoo.com/v8/finance/chart/{sym}` | 商品先物の終値・前日比 | JSON | 不要 | 日次 | 主データ源（WTI/Brent/天然ガス/ガスオイル/ガソリン/大豆油 等） |
| **Stooq** `stooq.com/q/l/?s={sym}&f=sd2t2ohlcv&e=csv` | 商品先物・FXの日足 | CSV | 不要 | 日次 | Yahoo障害時のフォールバック |
| **EIA Open Data API v2** `api.eia.gov/v2/...` | 米エネルギー省の現物価格（原油/留出油/天然ガス/残渣油 等） | JSON | **無料キー必要** | 日次/週次/月次 | 任意。`EIA_API_KEY` を Secrets 登録で有効化 |
| **World Bank "Pink Sheet"** `thedocs.worldbank.org/.../CMO-Pink-Sheet-*.pdf` | 国際商品月次価格（原油/天然ガス/肥料=アンモニア関連） | PDF/XLSX | 不要 | 月次 | アンモニア参考値の根拠 |
| **Methanex 価格シート** `methanex.com/our-products/about-methanol/pricing/` | メタノール地域別ポステッド価格 | Web/PDF | 不要 | 月次 | メタノール参考値の根拠 |
| **Ship & Bunker** `shipandbunker.com/prices` | 主要港の VLSFO/MGO/IFO380 実勢 | Web | 不要(閲覧) | 日次 | 重油系の実勢クロスチェック（自動取得はToS確認要） |

### APIキー不要で自動化が確実なのは Yahoo / Stooq
GitHub Actions のランナーは外部接続が自由なため、**キー不要の Yahoo / Stooq を主軸**にする。
EIA はより正確な現物価格を返すが無料キー登録が必要なため、**任意の高精度オプション**として組み込む。

---

## 2. 燃料 → 価格の対応（換算ロジック）

市場で直接の無料先物が立っていない船舶燃料は、関連する原料・指標から**透明な線形式**で目安を算出する。
係数は `scripts/fetch_prices.py` の `CONFIG` に集約し、いつでも調整可能。

| 燃料 | 取得方法 | 元データ | 換算ロジック（既定係数） | 区分 / 頻度 |
|---|---|---|---|---|
| **VLSFO（低硫黄重油）** | Brent連動 | Brent (BZ=F) $/bbl | `Brent×7.33 + 70` → $/t | Proxy / Daily |
| **HSFO / IFO380（高硫黄重油）** | VLSFO連動 | 上記 | `VLSFO − 95` → $/t | Proxy / Daily |
| **MGO（舶用軽油）** | ガスオイル連動 | Heating Oil (HO=F) $/gal | `HO×312 + 40` → $/t | Proxy / Daily |
| **バイオディーゼル（B100/FAME）** | 大豆油連動 | Soybean Oil (ZL=F) ¢/lb | `(¢/100)×2204.6 + 120` → $/t | Proxy / Daily |
| **エタノール** | 先物直接/連動 | Ethanol (EH=F) または砂糖/トウモロコシ | `$/gal×335` → $/t | Direct/Proxy / Daily |
| **メタノール** | ポステッド参考 | Methanex アジア契約価格 | 月次ポステッド値（CONFIG） | Reference / Monthly |
| **アンモニア（参考値）** | 月次参考 | World Bank / 公開コメンタリ | 月次参考値（CONFIG） | Reference / Monthly |
| **アンモニア（生産コスト推定）** | ガス連動 | Henry Hub (NG=F) $/MMBtu | `NG×34 + 180` → $/t | Model / Daily |
| **水素 グレー（化石由来）** | ガス連動 | Henry Hub (NG=F) | `NG×0.18 + 0.9` → $/kg | Model / Daily |
| **水素 グリーン（再エネ由来）** | 電力連動 | 電力単価（CONFIG） | `電力$/kWh×52 + 1.0` → $/kg | Model / Daily |

### エネルギー等価比較（$/GJ）
燃料は発熱量が大きく異なるため、`$/t` だけでなく **低位発熱量(LHV)で割った `$/GJ`** を併記する。
これにより「同じエネルギーあたりのコスト」で横並び比較でき、既存の燃料転換シミュレーターと接続できる。

| 燃料 | LHV (MJ/kg) |
|---|---|
| MGO | 42.7 |
| VLSFO / HSFO | 40.5 |
| バイオディーゼル | 37.2 |
| エタノール | 26.8 |
| メタノール | 19.9 |
| アンモニア | 18.6 |
| 水素 | 120.0 |

---

## 3. 自動更新の仕組み

- **スケジュール**：GitHub Actions `cron: "0 21 * * *"`（UTC 21:00 = **JST 06:00**）
- **処理**：`scripts/fetch_prices.py` が各ソースを取得 → `prices/data/latest.json` を更新し、`prices/data/history.csv` に追記 → 自動コミット
- **耐障害性**：いずれかのソースが失敗しても、直近の取得値（`history.csv`／前回 `latest.json`）を保持し、`stale: true` と取得時刻を明示
- **表示**：`prices/index.html` が `latest.json` を読み込み、最終更新時刻（JST）と次回更新予定（翌06:00 JST）を表示

手動実行も可能：リポジトリの Actions タブから "Update Fuel Prices" を `workflow_dispatch` で起動。

---

## 4. 既知の限界と注意

- **水素・アンモニア**には流動的な無料スポット市場が存在しない。本サイトの値は**原料連動のモデル推定**であり、実際の供給契約価格とは乖離しうる。
- **メタノール・アンモニア参考値**は月次更新のため、`latest.json` の `as_of` 日付を確認すること。
- **重油系（VLSFO/MGO）**は Brent / ガスオイル連動の推定。港別の実勢は Ship & Bunker を併用してクロスチェックすることを推奨。
- 商用利用・売買判断には必ず確報の有料データを使用すること。
