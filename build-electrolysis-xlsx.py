# -*- coding: utf-8 -*-
import math
import xlsxwriter
from xlsxwriter.utility import xl_rowcol_to_cell as cell

OUT_PATH = "/home/user/tasklist/electrolysis-cost-calculator.xlsx"

NM3 = 11.13
HHV = 39.4
HRS = 8760

PRESETS = {
    "アルカリ": dict(capexPerKW=130000, secNm3=4.8, stackShare=35, stackLifeH=80000, omPct=3.0),
    "PEM":      dict(capexPerKW=165000, secNm3=5.0, stackShare=45, stackLifeH=70000, omPct=3.5),
    "SOEC":     dict(capexPerKW=330000, secNm3=3.8, stackShare=50, stackLifeH=40000, omPct=4.0),
    "AEM":      dict(capexPerKW=140000, secNm3=4.8, stackShare=40, stackLifeH=30000, omPct=3.0),
}
TECHS = list(PRESETS.keys())

DEFAULTS = dict(
    capacityKW=1000, capexPerKW=165000, installFactor=20, stackShare=45,
    cf=46, lifetime=20, degradation=0, elecPrice=15, secNm3=5.0,
    omPct=3.5, waterUse=1.4, waterPrice=300, stackLifeH=70000,
    wacc=7, fx=150, subsidy=0, o2Credit=0, inclCompr=0, comprCapex=50000000, comprSEC=3.0,
)

# ---------- Python port of compute() for cached values ----------
def compute(s):
    r = s["wacc"] / 100.0; N = s["lifetime"]
    CRF = 1.0 / N if r <= 0 else r * (1 + r) ** N / ((1 + r) ** N - 1)
    hours = s["cf"] / 100.0 * HRS
    secKg = s["secNm3"] * NM3
    degF = 1 + (s["degradation"] / 100.0) * (N / 2.0)
    secEff = secKg * degF
    aH2 = s["capacityKW"] * hours / secEff if hours > 0 else 0
    capexSys = s["capacityKW"] * s["capexPerKW"]
    elx = capexSys * (1 + s["installFactor"] / 100.0)
    aCapex = elx * CRF
    fOM = elx * s["omPct"] / 100.0
    stackCost = capexSys * s["stackShare"] / 100.0
    yps = s["stackLifeH"] / hours if hours > 0 else 1e9
    nRepl = max(0, math.ceil(N / yps) - 1) if yps > 0 else 0
    if r <= 0:
        pv = stackCost * nRepl
    else:
        q = (1 + r) ** (-yps)
        pv = 0 if nRepl <= 0 else stackCost * q * (1 - q ** nRepl) / (1 - q)
    aStack = pv * CRF
    compr = 0.0
    if s["inclCompr"]:
        compr = ((s["comprCapex"] * CRF + s["comprCapex"] * s["omPct"] / 100.0) / aH2 if aH2 > 0 else 0) + s["elecPrice"] * s["comprSEC"]
    p = {}
    p["capex"] = aCapex / aH2 if aH2 > 0 else 0
    p["stack"] = aStack / aH2 if aH2 > 0 else 0
    p["om"] = fOM / aH2 if aH2 > 0 else 0
    p["elec"] = s["elecPrice"] * secEff
    p["water"] = s["waterUse"] * s["waterPrice"] / 1000.0 * NM3
    p["compr"] = compr
    p["subsidy"] = -(s["subsidy"] * NM3)
    p["o2"] = -(s["o2Credit"] * NM3)
    total = sum(p.values())
    return dict(p=p, total=total, aH2=aH2, CRF=CRF, secEff=secEff,
                aCapex=aCapex, fOM=fOM, stackCost=stackCost, r=r,
                aStack=aStack, pv=pv)

def scenario_state(tech):
    s = dict(DEFAULTS)
    s.update(PRESETS[tech])
    return s

# ====================================================================
wb = xlsxwriter.Workbook(OUT_PATH, {"nan_inf_to_errors": True})
wb.set_calc_mode("auto")

# ---------- formats ----------
C_INPUT = "#FFF3C4"; C_CONST = "#DCECFB"; C_CALC = "#FFFFFF"
C_HEAD = "#1F2A44"; C_SUB = "#33415C"; C_KPI = "#0EA5E9"
f_title = wb.add_format({"bold": True, "font_size": 16, "font_color": "#0B1120"})
f_sub   = wb.add_format({"font_size": 10, "font_color": "#475569"})
f_h2    = wb.add_format({"bold": True, "font_size": 11, "font_color": "white", "bg_color": C_HEAD, "align": "left", "valign": "vcenter", "border": 1, "border_color": "#0B1120"})
f_label = wb.add_format({"font_size": 10, "align": "left", "valign": "vcenter"})
f_note  = wb.add_format({"font_size": 9, "font_color": "#64748B", "italic": True, "align": "left"})
f_unit  = wb.add_format({"font_size": 9, "font_color": "#475569", "align": "left"})
def numfmt(bg, nf="#,##0.###", bold=False, color="#0B1120"):
    return wb.add_format({"bg_color": bg, "border": 1, "border_color": "#CBD5E1",
                          "num_format": nf, "align": "right", "bold": bold, "font_color": color})
f_in    = numfmt(C_INPUT); f_in0 = numfmt(C_INPUT, "#,##0"); f_in2 = numfmt(C_INPUT, "#,##0.0")
f_const = numfmt(C_CONST); f_const0 = numfmt(C_CONST, "#,##0")
f_calc  = numfmt(C_CALC, "#,##0.######"); f_calc2 = numfmt(C_CALC, "#,##0.00")
f_legend_in = wb.add_format({"bg_color": C_INPUT, "border": 1, "border_color": "#CBD5E1", "align": "center", "font_size": 9})
f_legend_co = wb.add_format({"bg_color": C_CONST, "border": 1, "border_color": "#CBD5E1", "align": "center", "font_size": 9})
f_legend_ca = wb.add_format({"bg_color": C_CALC, "border": 1, "border_color": "#CBD5E1", "align": "center", "font_size": 9})
f_tbl_h = wb.add_format({"bold": True, "font_color": "white", "bg_color": C_SUB, "border": 1, "border_color": "#0B1120", "align": "center"})
f_tbl_l = wb.add_format({"border": 1, "border_color": "#CBD5E1", "align": "left"})
f_tbl_n = wb.add_format({"border": 1, "border_color": "#CBD5E1", "num_format": "#,##0.00", "align": "right"})
f_tbl_p = wb.add_format({"border": 1, "border_color": "#CBD5E1", "num_format": "0.0%", "align": "right"})
f_tbl_tot = wb.add_format({"bold": True, "border": 1, "border_color": "#0B1120", "bg_color": "#EEF2FF", "num_format": "#,##0.00", "align": "right"})
f_tbl_totl = wb.add_format({"bold": True, "border": 1, "border_color": "#0B1120", "bg_color": "#EEF2FF", "align": "left"})
f_kpi_lbl = wb.add_format({"font_size": 10, "font_color": "#475569"})
f_kpi_big = wb.add_format({"bold": True, "font_size": 28, "font_color": C_KPI, "num_format": "#,##0.0"})
f_kpi_med = wb.add_format({"bold": True, "font_size": 14, "font_color": "#0B1120", "num_format": "#,##0.0"})
f_kpi_unit = wb.add_format({"font_size": 10, "font_color": "#475569"})

# sheet names
S_IN = "入力"; S_CO = "定数・前提"; S_CA = "計算"; S_OU = "結果"
ws_in = wb.add_worksheet(S_IN)
ws_ou = wb.add_worksheet(S_OU)
ws_co = wb.add_worksheet(S_CO)
ws_ca = wb.add_worksheet(S_CA)

def name(n, sheet, r, c):
    wb.define_name(n, "='{}'!${}${}".format(sheet, xlsxwriter.utility.xl_col_to_name(c), r + 1))

# ====================================================================
# 定数・前提 (CONST)
# ====================================================================
ws_co.set_column("A:A", 24); ws_co.set_column("B:B", 16); ws_co.set_column("C:C", 12); ws_co.set_column("D:G", 16)
ws_co.hide_gridlines(2)
ws_co.write("A1", "定数・前提（固定値／プリセット）", f_title)
ws_co.write("A2", "青セルは定数・前提。通常は変更不要（必要なら編集可）。方式プリセットは『入力』の参考列で参照されます。", f_sub)
r = 3
ws_co.merge_range(r, 0, r, 3, "物理定数・換算", f_h2); r += 1
const_rows = [
    ("NM3perKG", "kg→Nm³ 換算", NM3, "Nm³/kg", "0℃,1atm 密度0.0899kg/Nm³"),
    ("HHV", "水素 高位発熱量", HHV, "kWh/kg", "効率算定用 (141.8 MJ/kg)"),
    ("hoursPerYear", "年間時間", HRS, "h/年", "8760 = 24×365"),
    ("target2030", "日本 供給コスト目標 2030", 30, "円/Nm³", "経産省/NEDO"),
    ("target2050", "日本 供給コスト目標 2050", 20, "円/Nm³", "経産省/NEDO"),
]
for nm, lbl, val, unit, note in const_rows:
    ws_co.write(r, 0, lbl, f_label)
    ws_co.write_number(r, 1, val, f_const0 if val >= 1000 else f_const)
    ws_co.write(r, 2, unit, f_unit); ws_co.write_string(r, 3, note, f_note)
    name(nm, S_CO, r, 1); r += 1

r += 1
ws_co.merge_range(r, 0, r, 5, "方式プリセット（代表値・編集可）", f_h2); r += 1
hdr = ["方式", "CAPEX単価(円/kW)", "SEC(kWh/Nm³)", "スタック比率(%)", "スタック寿命(h)", "O&M(%/年)"]
for c, h in enumerate(hdr):
    ws_co.write(r, c, h, f_tbl_h)
r += 1
preset_first = r
for tech in TECHS:
    p = PRESETS[tech]
    ws_co.write(r, 0, tech, f_tbl_l)
    ws_co.write_number(r, 1, p["capexPerKW"], f_const0)
    ws_co.write_number(r, 2, p["secNm3"], f_const)
    ws_co.write_number(r, 3, p["stackShare"], f_const)
    ws_co.write_number(r, 4, p["stackLifeH"], f_const0)
    ws_co.write_number(r, 5, p["omPct"], f_const)
    r += 1
preset_last = r - 1
# named ranges for preset columns (for INDEX/MATCH)
col_letter = xlsxwriter.utility.xl_col_to_name
def colrange(c):
    return "='{}'!${}${}:${}${}".format(S_CO, col_letter(c), preset_first + 1, col_letter(c), preset_last + 1)
wb.define_name("presetNames", colrange(0))
wb.define_name("presetCapex", colrange(1))
wb.define_name("presetSEC", colrange(2))
wb.define_name("presetStackShare", colrange(3))
wb.define_name("presetStackLife", colrange(4))
wb.define_name("presetOM", colrange(5))

# ====================================================================
# 入力 (INPUT)
# ====================================================================
ws_in.set_column("A:A", 30); ws_in.set_column("B:B", 14); ws_in.set_column("C:C", 11); ws_in.set_column("D:D", 34)
ws_in.hide_gridlines(2)
ws_in.write("A1", "水電解 水素製造コスト 試算ツール ― 入力", f_title)
ws_in.write("A2", "黄色セルだけ編集してください。結果は『結果』シートに自動反映（グラフ付き）。", f_sub)
# legend
ws_in.write("A3", "凡例:", f_note)
ws_in.write("B3", "入力", f_legend_in); ws_in.write("C3", "定数", f_legend_co); ws_in.write("D3", "計算結果（自動）", f_legend_ca)

# field groups: (group title) then list of (key,label,unit,min,max,note_or_None,is_tech)
groups = [
    ("① 装置・規模", [
        ("capacityKW", "定格電力", "kW", 10, 100000, "システムへの定格入力電力", False),
    ]),
    ("② CAPEX（設備費）", [
        ("capexPerKW", "システムCAPEX単価", "円/kW", 20000, 500000, "据付除くシステム本体", True),
        ("installFactor", "据付・建設係数", "%", 0, 100, "システム費に上乗せ", False),
        ("stackShare", "スタック比率（対システム）", "%", 10, 70, "交換費の基礎", True),
    ]),
    ("③ 稼働プロファイル", [
        ("cf", "設備利用率（稼働率）", "%", 5, 95, "稼働率×8760 = 年間稼働時間", False),
        ("lifetime", "設備寿命", "年", 5, 30, None, False),
        ("degradation", "劣化率（電力原単位の上昇）", "%/年", 0, 3, "0で劣化なし", False),
    ]),
    ("④ 電力", [
        ("elecPrice", "電力単価", "円/kWh", 1, 40, "LCOHの最大要素", False),
        ("secNm3", "電力原単位（SEC）", "kWh/Nm³", 3, 7, "方式で変わる", True),
    ]),
    ("⑤ O&M・水", [
        ("omPct", "固定O&M（対CAPEX）", "%/年", 0, 8, None, True),
        ("waterUse", "水消費", "L/Nm³", 0, 5, None, False),
        ("waterPrice", "水単価", "円/m³", 0, 2000, None, False),
    ]),
    ("⑥ スタック", [
        ("stackLifeH", "スタック寿命", "時間(h)", 10000, 120000, "交換周期の基礎", True),
    ]),
    ("⑦ 財務", [
        ("wacc", "割引率（WACC）", "%", 0, 15, "資本回収係数CRFに使用", False),
        ("fx", "為替", "円/$", 80, 250, "$表示用", False),
    ]),
    ("⑧ 補助金・副生酸素（任意）", [
        ("subsidy", "補助金", "円/Nm³", 0, 100, "コストから差引", False),
        ("o2Credit", "副生酸素 販売益", "円/Nm³", 0, 50, "コストから差引", False),
    ]),
    ("⑨ 圧縮・出荷（任意）", [
        ("inclCompr", "圧縮・出荷を含める", "1=含/0=否", 0, 1, "1で下の2項目を反映", False),
        ("comprCapex", "圧縮機等CAPEX（総額）", "円", 0, 2000000000, None, False),
        ("comprSEC", "圧縮 電力原単位", "kWh/kg", 0, 10, None, False),
    ]),
]

r = 4
# tech selector (reference only)
ws_in.merge_range(r, 0, r, 3, "方式プリセット参照", f_h2); r += 1
ws_in.write(r, 0, "方式（参考値の表示用）", f_label)
ws_in.write(r, 1, "PEM", wb.add_format({"bg_color": C_INPUT, "border": 1, "border_color": "#CBD5E1", "align": "center"}))
name("tech", S_IN, r, 1)
ws_in.data_validation(r, 1, r, 1, {"validate": "list", "source": TECHS,
                                   "input_title": "方式", "input_message": "代表値を右列に参照表示します（入力欄は手動編集）"})
ws_in.write(r, 3, "右の入力欄に手動で値を入れます。選択方式の代表値は各行の備考に表示。", f_note)
r += 2

addr = {}  # key -> (row, col=1)
for title, fields in groups:
    ws_in.merge_range(r, 0, r, 3, title, f_h2); r += 1
    for key, lbl, unit, mn, mx, note, is_tech in fields:
        ws_in.write(r, 0, lbl, f_label)
        val = DEFAULTS[key]
        fmt = f_in0 if abs(val) >= 1000 or key in ("capacityKW", "capexPerKW", "stackLifeH", "comprCapex", "waterPrice") else f_in
        ws_in.write_number(r, 1, val, fmt)
        name(key, S_IN, r, 1)
        addr[key] = (r, 1)
        ws_in.write(r, 2, unit, f_unit)
        # note / reference
        if is_tech:
            colmap = {"capexPerKW": "presetCapex", "secNm3": "presetSEC", "stackShare": "presetStackShare",
                      "stackLifeH": "presetStackLife", "omPct": "presetOM"}
            ref = colmap[key]
            f = '=("参考("&tech&"): "&TEXT(INDEX({},MATCH(tech,presetNames,0)),"#,##0.###")&" {}")'.format(ref, unit)
            ws_in.write_formula(r, 3, f, f_note)
        elif note:
            ws_in.write_string(r, 3, note, f_note)
        # data validation
        dv = {"validate": "decimal", "criteria": "between", "minimum": mn, "maximum": mx,
              "input_title": lbl, "input_message": "範囲: {} 〜 {} {}".format(mn, mx, unit),
              "error_message": "範囲 {}〜{} で入力してください".format(mn, mx)}
        ws_in.data_validation(r, 1, r, 1, dv)
        r += 1
    r += 1

# ====================================================================
# 計算 (CALC) — scenario grid + sensitivity
# ====================================================================
ws_ca.set_column("A:A", 26); ws_ca.set_column("B:B", 3); ws_ca.set_column("C:I", 14)
ws_ca.hide_gridlines(2)
ws_ca.write("A1", "計算（LCOHエンジン）", f_title)
ws_ca.write("A2", "数式は『入力』『定数・前提』の名前を参照。編集不要。", f_sub)

# scenario columns: 現在 + 4 techs
scen_cols = list(range(2, 7))  # C,D,E,F,G  -> 現在, アルカリ, PEM, SOEC, AEM
scen_names = ["現在"] + TECHS
hdr_row = 3
for i, nmrow in enumerate(scen_names):
    ws_ca.write(hdr_row, scen_cols[i], nmrow, f_tbl_h)
ws_ca.write(hdr_row, 0, "項目", f_tbl_h)

# line items with (label, key) ; formulas built per column
items = ["r", "N", "CRF", "hours", "secKg", "degF", "secEff", "aH2", "capexSys",
         "elx", "aCapex", "fOM", "stackCost", "yps", "nRepl", "q", "pvRepl", "aStack", "compr",
         "p_capex", "p_stack", "p_om", "p_elec", "p_water", "p_compr", "p_subsidy", "p_o2",
         "grossPos", "totalPerKg", "totalNm3", "totalUSD"]
labels = {"r": "r=WACC", "N": "寿命N", "CRF": "資本回収係数CRF", "hours": "年間稼働時間",
          "secKg": "SEC(kWh/kg)", "degF": "劣化係数", "secEff": "実効SEC", "aH2": "年間H2(kg)",
          "capexSys": "システムCAPEX", "elx": "総CAPEX(据付込)", "aCapex": "CAPEX年経費",
          "fOM": "固定O&M(円/年)", "stackCost": "スタック費/回", "yps": "交換周期(年)",
          "nRepl": "交換回数", "q": "割引率^(-周期)", "pvRepl": "交換費PV", "aStack": "スタック年経費",
          "compr": "圧縮(円/kg)", "p_capex": "設備費(円/kg)", "p_stack": "スタック交換(円/kg)",
          "p_om": "固定O&M(円/kg)", "p_elec": "電力費(円/kg)", "p_water": "水費(円/kg)",
          "p_compr": "圧縮・出荷(円/kg)", "p_subsidy": "補助金(円/kg)", "p_o2": "副生酸素(円/kg)",
          "grossPos": "粗コスト合計(円/kg)", "totalPerKg": "LCOH(円/kg)", "totalNm3": "LCOH(円/Nm³)",
          "totalUSD": "LCOH($/kg)"}
row_of = {it: hdr_row + 1 + i for i, it in enumerate(items)}
for it in items:
    ws_ca.write(row_of[it], 0, labels[it], f_label)

cached = {"現在": compute(DEFAULTS)}
for t in TECHS:
    cached[t] = compute(scenario_state(t))

def cref(it, col):  # same-column cell ref
    return cell(row_of[it], col)

for si, scol in enumerate(scen_cols):
    nmrow = scen_names[si]
    if si == 0:
        T = dict(capex="capexPerKW", sec="secNm3", share="stackShare", life="stackLifeH", om="omPct")
        cv = cached["現在"]; st = DEFAULTS
    else:
        tech = TECHS[si - 1]
        # tech params -> preset cell refs
        pr = preset_first + (si - 1)
        T = dict(capex="'{}'!${}${}".format(S_CO, "B", pr + 1),
                 sec="'{}'!${}${}".format(S_CO, "C", pr + 1),
                 share="'{}'!${}${}".format(S_CO, "D", pr + 1),
                 life="'{}'!${}${}".format(S_CO, "E", pr + 1),
                 om="'{}'!${}${}".format(S_CO, "F", pr + 1))
        cv = cached[tech]; st = scenario_state(tech)
    p = cv["p"]
    F = {
        "r": ("=wacc/100", cv["r"]),
        "N": ("=lifetime", st["lifetime"]),
        "CRF": ("=IF({r}<=0,1/{N},{r}*(1+{r})^{N}/((1+{r})^{N}-1))".format(r=cref("r", scol), N=cref("N", scol)), cv["CRF"]),
        "hours": ("=cf/100*hoursPerYear", st["cf"]/100*HRS),
        "secKg": ("={}*NM3perKG".format(T["sec"]), st["secNm3"]*NM3),
        "degF": ("=1+(degradation/100)*({}/2)".format(cref("N", scol)), 1+(st["degradation"]/100)*(st["lifetime"]/2)),
        "secEff": ("={}*{}".format(cref("secKg", scol), cref("degF", scol)), cv["secEff"]),
        "aH2": ("=IF({h}>0,capacityKW*{h}/{se},0)".format(h=cref("hours", scol), se=cref("secEff", scol)), cv["aH2"]),
        "capexSys": ("=capacityKW*{}".format(T["capex"]), st["capacityKW"]*st["capexPerKW"]),
        "elx": ("={}*(1+installFactor/100)".format(cref("capexSys", scol)), st["capacityKW"]*st["capexPerKW"]*(1+st["installFactor"]/100)),
        "aCapex": ("={}*{}".format(cref("elx", scol), cref("CRF", scol)), cv["aCapex"]),
        "fOM": ("={}*{}/100".format(cref("elx", scol), T["om"]), cv["fOM"]),
        "stackCost": ("={}*{}/100".format(cref("capexSys", scol), T["share"]), cv["stackCost"]),
        "yps": ("=IF({h}>0,{life}/{h},1E+09)".format(h=cref("hours", scol), life=T["life"]),
                (st["stackLifeH"]/(st["cf"]/100*HRS)) if st["cf"] > 0 else 1e9),
        "nRepl": ("=MAX(0,CEILING({N}/{yps},1)-1)".format(N=cref("N", scol), yps=cref("yps", scol)),
                  max(0, math.ceil(st["lifetime"]/(st["stackLifeH"]/(st["cf"]/100*HRS)))-1)),
        "q": ("=(1+{r})^(-{yps})".format(r=cref("r", scol), yps=cref("yps", scol)),
              (1+cv["r"])**(-(st["stackLifeH"]/(st["cf"]/100*HRS)))),
        "pvRepl": ("=IF({r}<=0,{sc}*{n},IF({n}<=0,0,{sc}*{q}*(1-{q}^{n})/(1-{q})))".format(
            r=cref("r", scol), sc=cref("stackCost", scol), n=cref("nRepl", scol), q=cref("q", scol)),
            cv["aStack"]/cv["CRF"] if cv["CRF"] else 0),
        "aStack": ("={}*{}".format(cref("pvRepl", scol), cref("CRF", scol)), cv["aStack"]),
        "compr": ("=IF(inclCompr=1,IF({a}>0,(comprCapex*{crf}+comprCapex*{om}/100)/{a},0)+elecPrice*comprSEC,0)".format(
            a=cref("aH2", scol), crf=cref("CRF", scol), om=T["om"]), p["compr"]),
        "p_capex": ("=IF({a}>0,{ac}/{a},0)".format(a=cref("aH2", scol), ac=cref("aCapex", scol)), p["capex"]),
        "p_stack": ("=IF({a}>0,{as_}/{a},0)".format(a=cref("aH2", scol), as_=cref("aStack", scol)), p["stack"]),
        "p_om": ("=IF({a}>0,{fo}/{a},0)".format(a=cref("aH2", scol), fo=cref("fOM", scol)), p["om"]),
        "p_elec": ("=elecPrice*{}".format(cref("secEff", scol)), p["elec"]),
        "p_water": ("=waterUse*waterPrice/1000*NM3perKG", p["water"]),
        "p_compr": ("={}".format(cref("compr", scol)), p["compr"]),
        "p_subsidy": ("=-(subsidy*NM3perKG)", p["subsidy"]),
        "p_o2": ("=-(o2Credit*NM3perKG)", p["o2"]),
        "grossPos": ("={}+{}+{}+{}+{}+{}".format(cref("p_capex", scol), cref("p_stack", scol), cref("p_om", scol),
                     cref("p_elec", scol), cref("p_water", scol), cref("p_compr", scol)),
                     p["capex"]+p["stack"]+p["om"]+p["elec"]+p["water"]+p["compr"]),
        "totalPerKg": ("={}+{}+{}+{}+{}+{}+{}+{}".format(*[cref(k, scol) for k in
                       ["p_capex", "p_stack", "p_om", "p_elec", "p_water", "p_compr", "p_subsidy", "p_o2"]]), cv["total"]),
        "totalNm3": ("={}/NM3perKG".format(cref("totalPerKg", scol)), cv["total"]/NM3),
        "totalUSD": ("={}/fx".format(cref("totalPerKg", scol)), cv["total"]/st["fx"]),
    }
    for it in items:
        formula, value = F[it]
        fmt = f_calc2 if it in ("totalNm3", "totalUSD", "p_capex", "p_stack", "p_om", "p_elec", "p_water", "p_compr") else f_calc
        ws_ca.write_formula(row_of[it], scol, formula, fmt, value)

# named ranges from 現在 (col C = index 2)
NOW = 2
name("lcohPerKg", S_CA, row_of["totalPerKg"], NOW)
name("lcohNm3", S_CA, row_of["totalNm3"], NOW)
name("lcohUSD", S_CA, row_of["totalUSD"], NOW)
name("annualH2kg", S_CA, row_of["aH2"], NOW)
name("grossPos", S_CA, row_of["grossPos"], NOW)
for comp in ["p_capex", "p_stack", "p_om", "p_elec", "p_water", "p_compr", "p_subsidy", "p_o2"]:
    name("now_" + comp, S_CA, row_of[comp], NOW)

# --- stacked-bar data: components in 円/Nm³ across scenarios ---
bar_first = row_of["totalUSD"] + 3
comp_keys = ["p_capex", "p_stack", "p_om", "p_elec", "p_water", "p_compr", "p_subsidy", "p_o2"]
comp_lbl = {"p_capex": "設備費(CAPEX)", "p_stack": "スタック交換", "p_om": "固定O&M",
            "p_elec": "電力費", "p_water": "水費", "p_compr": "圧縮・出荷", "p_subsidy": "補助金", "p_o2": "副生酸素"}
ws_ca.write(bar_first - 1, 0, "▼ 積み上げグラフ用（円/Nm³）", f_h2)
for i, sc in enumerate(scen_cols):
    ws_ca.write(bar_first - 0 if False else bar_first - 1, sc, "", f_h2)
# header of scenarios for bar block
bar_hdr = bar_first
ws_ca.write(bar_hdr, 0, "成分＼方式", f_tbl_h)
for i, sc in enumerate(scen_cols):
    ws_ca.write(bar_hdr, sc, scen_names[i], f_tbl_h)
bar_data_first = bar_hdr + 1
bar_rows = {}
for j, ck in enumerate(comp_keys):
    rr = bar_data_first + j
    bar_rows[ck] = rr
    ws_ca.write(rr, 0, comp_lbl[ck], f_tbl_l)
    for i, sc in enumerate(scen_cols):
        src = cell(row_of[ck], sc)
        val = cached[scen_names[i]]["p"][ck.replace("p_", "")] / NM3
        ws_ca.write_formula(rr, sc, "={}/NM3perKG".format(src), f_calc2, val)
# target lines rows
t2030_row = bar_data_first + len(comp_keys)
t2050_row = t2030_row + 1
ws_ca.write(t2030_row, 0, "2030目標", f_tbl_l)
ws_ca.write(t2050_row, 0, "2050目標", f_tbl_l)
for i, sc in enumerate(scen_cols):
    ws_ca.write_formula(t2030_row, sc, "=target2030", f_calc2, 30)
    ws_ca.write_formula(t2050_row, sc, "=target2050", f_calc2, 20)

# --- sensitivity table: LCOH vs capacity factor ---
sens_first = t2050_row + 3
ws_ca.write(sens_first - 1, 0, "▼ 感度分析：設備利用率 vs LCOH", f_h2)
sh = sens_first
sens_hdrs = ["設備利用率(%)", "hours", "aH2", "p_capex", "yps", "nRepl", "q", "pvRepl", "aStack", "p_stack", "compr", "LCOH(円/Nm³)"]
for c, h in enumerate(sens_hdrs):
    ws_ca.write(sh, c, h, f_tbl_h)
sens_data_first = sh + 1
cf_values = list(range(10, 96, 5))
# references to 現在 column cells (cf-independent parts)
now = NOW
ref_secEff = cell(row_of["secEff"], now); ref_aCapex = cell(row_of["aCapex"], now)
ref_fOM = cell(row_of["fOM"], now); ref_stackCost = cell(row_of["stackCost"], now)
ref_CRF = cell(row_of["CRF"], now); ref_r = cell(row_of["r"], now); ref_N = cell(row_of["N"], now)
ref_pelec = cell(row_of["p_elec"], now); ref_pwater = cell(row_of["p_water"], now)
ref_psub = cell(row_of["p_subsidy"], now); ref_po2 = cell(row_of["p_o2"], now)
for k, cfv in enumerate(cf_values):
    rr = sens_data_first + k
    A = cell(rr, 0); HH = cell(rr, 1); AH = cell(rr, 2); PC = cell(rr, 3)
    YP = cell(rr, 4); NR = cell(rr, 5); QQ = cell(rr, 6); PVc = cell(rr, 7)
    ASc = cell(rr, 8); PSc = cell(rr, 9); CMc = cell(rr, 10)
    s_cf = dict(DEFAULTS); s_cf["cf"] = cfv
    cvk = compute(s_cf)
    pk = cvk["p"]
    ws_ca.write_number(rr, 0, cfv, f_calc)
    ws_ca.write_formula(rr, 1, "={}/100*hoursPerYear".format(A), f_calc, cfv/100*HRS)
    ws_ca.write_formula(rr, 2, "=IF({h}>0,capacityKW*{h}/{se},0)".format(h=HH, se=ref_secEff), f_calc, cvk["aH2"])
    ws_ca.write_formula(rr, 3, "=IF({a}>0,{ac}/{a},0)".format(a=AH, ac=ref_aCapex), f_calc2, pk["capex"])
    ws_ca.write_formula(rr, 4, "=IF({h}>0,stackLifeH/{h},1E+09)".format(h=HH), f_calc, (s_cf["stackLifeH"]/(cfv/100*HRS))),
    ws_ca.write_formula(rr, 5, "=MAX(0,CEILING({N}/{yp},1)-1)".format(N=ref_N, yp=YP), f_calc, max(0, math.ceil(s_cf["lifetime"]/(s_cf["stackLifeH"]/(cfv/100*HRS)))-1))
    ws_ca.write_formula(rr, 6, "=(1+{r})^(-{yp})".format(r=ref_r, yp=YP), f_calc, (1+cvk["r"])**(-(s_cf["stackLifeH"]/(cfv/100*HRS))))
    ws_ca.write_formula(rr, 7, "=IF({r}<=0,{sc}*{n},IF({n}<=0,0,{sc}*{q}*(1-{q}^{n})/(1-{q})))".format(r=ref_r, sc=ref_stackCost, n=NR, q=QQ), f_calc, cvk["aStack"]/cvk["CRF"] if cvk["CRF"] else 0)
    ws_ca.write_formula(rr, 8, "={}*{}".format(PVc, ref_CRF), f_calc, cvk["aStack"])
    ws_ca.write_formula(rr, 9, "=IF({a}>0,{as_}/{a},0)".format(a=AH, as_=ASc), f_calc2, pk["stack"])
    ws_ca.write_formula(rr, 10, "=IF(inclCompr=1,IF({a}>0,(comprCapex*{crf}+comprCapex*omPct/100)/{a},0)+elecPrice*comprSEC,0)".format(a=AH, crf=ref_CRF), f_calc2, pk["compr"])
    # LCOH(円/Nm³)
    lf = "=({pc}+{ps}+IF({a}>0,{fo}/{a},0)+{cm}+{pe}+{pw}+{psub}+{po2})/NM3perKG".format(
        pc=PC, ps=PSc, a=AH, fo=ref_fOM, cm=CMc, pe=ref_pelec, pw=ref_pwater, psub=ref_psub, po2=ref_po2)
    ws_ca.write_formula(rr, 11, lf, f_calc2, cvk["total"]/NM3)
sens_data_last = sens_data_first + len(cf_values) - 1

# ====================================================================
# 結果 (OUTPUT) — KPIs, table, charts
# ====================================================================
ws_ou.set_column("A:A", 18); ws_ou.set_column("B:B", 14); ws_ou.set_column("C:C", 12)
ws_ou.set_column("D:D", 4); ws_ou.set_column("E:H", 13)
ws_ou.hide_gridlines(2)
ws_ou.write("A1", "水電解 水素製造コスト 試算結果", f_title)
ws_ou.write("A2", "数値は『入力』を変えると自動更新。方式比較の積み上げ・構成比・感度をグラフ表示。", f_sub)

# KPI block
ws_ou.write("A4", "水素製造コスト (LCOH)", f_kpi_lbl)
ws_ou.write_formula("A5", "=lcohNm3", f_kpi_big, cached["現在"]["total"]/NM3)
ws_ou.write("C5", "円/Nm³", f_kpi_unit)
ws_ou.write("A6", "円/kg", f_kpi_lbl); ws_ou.write_formula("B6", "=lcohPerKg", f_kpi_med, cached["現在"]["total"])
ws_ou.write("A7", "$/kg", f_kpi_lbl); ws_ou.write_formula("B7", "=lcohUSD", wb.add_format({"bold": True, "font_size": 14, "num_format": "$#,##0.00"}), cached["現在"]["total"]/150)
ws_ou.write("A9", "年間生産量", f_kpi_lbl)
ws_ou.write_formula("B9", "=annualH2kg*NM3perKG", wb.add_format({"num_format": "#,##0", "bold": True}), cached["現在"]["aH2"]*NM3)
ws_ou.write("C9", "Nm³/年", f_kpi_unit)
ws_ou.write_formula("B10", "=annualH2kg/1000", wb.add_format({"num_format": "#,##0.0"}), cached["現在"]["aH2"]/1000)
ws_ou.write("C10", "t/年", f_kpi_unit)
ws_ou.write("A12", "電力費シェア", f_kpi_lbl)
ws_ou.write_formula("B12", "=now_p_elec/grossPos", wb.add_format({"num_format": "0%", "bold": True, "font_color": C_KPI}), cached["現在"]["p"]["elec"]/(cached["現在"]["total"]-cached["現在"]["p"]["subsidy"]-cached["現在"]["p"]["o2"]))
ws_ou.write("A13", "2030目標差", f_kpi_lbl)
ws_ou.write_formula("B13", "=lcohNm3-target2030", wb.add_format({"num_format": '+#,##0.0;-#,##0.0', "bold": True}), cached["現在"]["total"]/NM3-30)
ws_ou.write("C13", "円/Nm³", f_kpi_unit)

# breakdown table
tr = 15
ws_ou.write(tr, 0, "内訳", f_tbl_h); ws_ou.write(tr, 1, "円/Nm³", f_tbl_h); ws_ou.write(tr, 2, "円/kg", f_tbl_h)
ws_ou.write(tr, 3, "", f_tbl_h)
ws_ou.write(tr, 4, "$/kg", f_tbl_h); ws_ou.write(tr, 5, "構成比", f_tbl_h)
comp_order = [("now_p_capex", "設備費(CAPEX)"), ("now_p_stack", "スタック交換"), ("now_p_om", "固定O&M"),
              ("now_p_elec", "電力費"), ("now_p_water", "水費"), ("now_p_compr", "圧縮・出荷"),
              ("now_p_subsidy", "補助金"), ("now_p_o2", "副生酸素")]
ck2 = ["capex", "stack", "om", "elec", "water", "compr", "subsidy", "o2"]
rr = tr + 1
for (nmf, lbl), k in zip(comp_order, ck2):
    pv = cached["現在"]["p"][k]
    ws_ou.write(rr, 0, lbl, f_tbl_l)
    ws_ou.write_formula(rr, 1, "={}/NM3perKG".format(nmf), f_tbl_n, pv/NM3)
    ws_ou.write_formula(rr, 2, "={}".format(nmf), f_tbl_n, pv)
    ws_ou.write_formula(rr, 4, "={}/fx".format(nmf), wb.add_format({"border": 1, "border_color": "#CBD5E1", "num_format": "$#,##0.000", "align": "right"}), pv/150)
    ws_ou.write_formula(rr, 5, "={}/grossPos".format(nmf), f_tbl_p, pv/(cached["現在"]["total"]-cached["現在"]["p"]["subsidy"]-cached["現在"]["p"]["o2"]))
    rr += 1
ws_ou.write(rr, 0, "合計 LCOH", f_tbl_totl)
ws_ou.write_formula(rr, 1, "=lcohNm3", f_tbl_tot, cached["現在"]["total"]/NM3)
ws_ou.write_formula(rr, 2, "=lcohPerKg", f_tbl_tot, cached["現在"]["total"])
ws_ou.write_formula(rr, 4, "=lcohUSD", wb.add_format({"bold": True, "border": 1, "bg_color": "#EEF2FF", "num_format": "$#,##0.000", "align": "right"}), cached["現在"]["total"]/150)
ws_ou.write(rr, 5, "", f_tbl_tot)
tbl_last = rr

# ---------- charts ----------
COLORS = ["#38BDF8", "#818CF8", "#2DD4BF", "#FBBF24", "#22D3EE", "#A78BFA", "#34D399", "#4ADE80"]

# 1) stacked column + target lines (combo)
colchart = wb.add_chart({"type": "column", "subtype": "stacked"})
for j, ck in enumerate(comp_keys):
    colchart.add_series({
        "name":       [S_CA, bar_hdr, 0] if False else comp_lbl[ck],
        "categories": [S_CA, bar_hdr, scen_cols[0], bar_hdr, scen_cols[-1]],
        "values":     [S_CA, bar_rows[ck], scen_cols[0], bar_rows[ck], scen_cols[-1]],
        "fill":       {"color": COLORS[j]},
        "border":     {"none": True},
        "gap": 80,
    })
linechart = wb.add_chart({"type": "line"})
for trow, tname, tcol in [(t2030_row, "2030目標 (30)", "#F59E0B"), (t2050_row, "2050目標 (20)", "#10B981")]:
    linechart.add_series({
        "name": tname,
        "categories": [S_CA, bar_hdr, scen_cols[0], bar_hdr, scen_cols[-1]],
        "values": [S_CA, trow, scen_cols[0], trow, scen_cols[-1]],
        "line": {"color": tcol, "width": 1.5, "dash_type": "dash"},
        "marker": {"type": "none"},
    })
colchart.combine(linechart)
colchart.set_title({"name": "コスト内訳（積み上げ）と目標 ― 方式比較"})
colchart.set_x_axis({"name": "方式"})
colchart.set_y_axis({"name": "LCOH (円/Nm³)", "major_gridlines": {"visible": True}})
colchart.set_size({"width": 560, "height": 360})
colchart.set_legend({"position": "bottom"})
ws_ou.insert_chart("H4", colchart)

# 2) doughnut (現在 share)
dough = wb.add_chart({"type": "doughnut"})
dough.add_series({
    "name": "構成比（現在・円/Nm³）",
    "categories": [S_CA, bar_data_first, 0, bar_data_first + 5, 0],   # 6 positive components labels
    "values":     [S_CA, bar_data_first, NOW, bar_data_first + 5, NOW],
    "points": [{"fill": {"color": COLORS[i]}} for i in range(6)],
})
dough.set_title({"name": "構成比（現在）"})
dough.set_size({"width": 360, "height": 300})
dough.set_legend({"position": "right", "font": {"size": 9}})
ws_ou.insert_chart("H24", dough)

# 3) sensitivity line
sens = wb.add_chart({"type": "line"})
sens.add_series({
    "name": "LCOH vs 設備利用率",
    "categories": [S_CA, sens_data_first, 0, sens_data_last, 0],
    "values":     [S_CA, sens_data_first, 11, sens_data_last, 11],
    "line": {"color": "#38BDF8", "width": 2.25},
    "marker": {"type": "circle", "size": 4, "fill": {"color": "#38BDF8"}},
})
sens.set_title({"name": "感度分析：設備利用率 → LCOH"})
sens.set_x_axis({"name": "設備利用率 (%)"})
sens.set_y_axis({"name": "LCOH (円/Nm³)"})
sens.set_size({"width": 560, "height": 300})
sens.set_legend({"none": True})
ws_ou.insert_chart("O24", sens)

ws_ou.activate()
wb.close()
print("written:", OUT_PATH)
print("cached 現在 LCOH 円/Nm³ =", round(cached["現在"]["total"]/NM3, 2))
for t in TECHS:
    print("  ", t, "LCOH 円/Nm³ =", round(cached[t]["total"]/NM3, 2))
