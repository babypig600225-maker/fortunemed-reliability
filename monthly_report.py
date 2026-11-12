"""
月度分析：系列（合計）＋ 機台（個別）表、狀態判定、分析與建議、Excel 匯出
"""
import io
import pandas as pd

FAULT_CATS = [
    ("溫度異常", ["溫度", "過熱", "溫控", "感溫"],
     "建議檢查溫控器、感溫線與加熱元件，並納入定期點檢"),
    ("漏油/漏水", ["漏油", "漏水", "滲油", "漏氣"],
     "建議加強油路／水路與接頭的定期檢查，預防洩漏"),
    ("電氣/電線", ["電線", "電路", "跳電", "斷電", "短路", "配電", "電源"],
     "建議檢查線材固定與端子，並安排電氣定期點檢"),
    ("機械磨損", ["斷裂", "磨損", "螺絲", "異音", "鬆脫", "皮帶", "軸承"],
     "建議定期檢查磨耗件與鎖固狀態，備妥關鍵備品"),
    ("感測/控制", ["感應", "感測", "PLC", "電磁閥", "參數", "速度"],
     "建議校正感測器並確認控制參數設定"),
    ("無法啟動", ["無法啟動", "無法使用", "無法運作", "不動作", "當機"],
     "建議追查根本原因，並建立異常排除步驟"),
    ("保養/更換", ["保養", "更換", "清洗", "清理"],
     "建議納入預防保養計畫，依週期執行"),
]

STATUS_ORDER = {"需重點改善": 0, "需改善": 1, "正常": 2, "表現卓越": 3}
STATUS_STYLE = {
    "需重點改善": ("#fcebeb", "#a32d2d"),
    "需改善":     ("#fff2cc", "#7b4f00"),
    "正常":       ("#e8f0fc", "#2563c9"),
    "表現卓越":   ("#e2efda", "#1f7a4d"),
}


def fault_category(text):
    t = str(text or "")
    for name, kws, sug in FAULT_CATS:
        if any(k in t for k in kws):
            return name, sug
    return "其他", "建議持續追蹤故障原因"


def _fmt(v, unit="H"):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{v:g}{unit}"


def _seq(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        return str(int(float(v)))
    except (TypeError, ValueError):
        return str(v)


def _item_text(r):
    need = str(r.get("需求", "") or "").strip().replace("\n", " ")
    if len(need) > 18:
        need = need[:18] + "…"
    down = float(r.get("停工時數(h)", 0) or 0)
    rep = float(r.get("維修時數", 0) or 0)
    parts = []
    if down > 0:
        parts.append(f"{down:g}H停機")
    parts.append(f"維修{rep:g}H")
    sq = _seq(r.get("seq"))
    if sq:
        parts.append(f"seq={sq}")
    return f"{r['機台代碼']}「{need}」（{'，'.join(parts)}）"


def series_comment(p, recs):
    """p: pool() 結果；recs: 該系列該月報修單"""
    if not p["faults"]:
        return "本月無故障記錄，設備運作正常。"
    txt = (f"本月故障{p['faults']}次（在用{p['n']}台中{p['n_fault']}台），"
           f"合計MTBF {_fmt(p['mtbf'])}、MTTR {_fmt(p['mttr'])}。")
    if len(recs):
        r = recs.sort_values(["停工時數(h)", "維修時數"], ascending=False)
        top = [_item_text(x) for _, x in r.head(2).iterrows()]
        txt += "主要項目：" + "；".join(top) + "。"
        cnt = recs.groupby("機台代碼").size()
        rep_m = cnt[cnt >= 2].sort_values(ascending=False)
        if len(rep_m):
            txt += "重複故障：" + "、".join(f"{c}（{n}次）" for c, n in rep_m.head(3).items()) + "。"
        _, sug = fault_category(r.iloc[0].get("需求", ""))
        sugs = [sug]
        if p["mttr"] is not None and p["mttr"] > 8:
            sugs.append("修復時間偏長，建議檢討備品與維修流程")
        if len(rep_m):
            sugs.append("針對重複故障機台追查根本原因")
        txt += "；".join(sugs) + "。"
    return txt


def build_month_tables(month, dk, dr, pool, judge, series_name):
    """回傳 (系列表, 機台表)；dk = build_full 結果（單月），dr = 該月報修單"""
    dkm = dk[dk["月份"] == month]
    drm = dr[dr["月份"] == month]
    hours = float(dkm["月可用時數(h)"].iloc[0]) if len(dkm) else 0

    mach_rows = []
    for _, m in dkm.sort_values("機台代碼").iterrows():
        f = int(m["故障次數"])
        mtbf = None if pd.isna(m["MTBF(h)"]) else float(m["MTBF(h)"])
        mttr = None if pd.isna(m["MTTR(h)"]) else float(m["MTTR(h)"])
        recs = drm[drm["機台代碼"] == m["機台代碼"]]
        main = ""
        if len(recs):
            r = recs.sort_values(["停工時數(h)", "維修時數"], ascending=False)
            main = "；".join(_item_text(x) for _, x in r.head(2).iterrows())
        mach_rows.append({
            "月份": month.replace("2026-", "2026-"), "月標準時數": hours,
            "設備類型": m["設備類型"], "設備系列": m["系列"],
            "設備代號": m["機台代碼"], "設備名稱": m["機台名稱"],
            "使用狀態": m.get("使用狀態", ""),
            "故障次數": f, "總維修工時(H)": round(float(m["總維修時數(h)"]), 2),
            "總維修停工時間(H)": round(float(m["停工時數(h)"]), 2),
            "MTBF(小時/次)": mtbf if f else "無故障",
            "MTTR(小時/次)": mttr if f else 0,
            "狀態判定": judge(f, mtbf, mttr),
            "主要故障": main,
        })
    df_m = pd.DataFrame(mach_rows)

    ser_rows = []
    for sc, g in dkm.groupby("系列"):
        p = pool(g)
        st_ = judge(p["faults"], p["mtbf"], p["mttr"])
        watch = df_m[(df_m["設備系列"] == sc) & df_m["狀態判定"].isin(["需重點改善", "需改善"])]
        ser_rows.append({
            "月份": month, "月標準時數": hours,
            "設備類型": g["設備類型"].iloc[0], "設備名稱": series_name.get(sc, ""),
            "設備系列": sc, "在用台數": p["n"], "故障台數": p["n_fault"],
            "故障總次數": p["faults"], "總維修工時(H)": p["repair"],
            "總維修停工時間(H)": p["downtime"],
            "MTBF(小時/次)": p["mtbf"] if p["faults"] else "無故障",
            "MTTR(小時/次)": p["mttr"] if p["faults"] else 0,
            "狀態判定": st_,
            "需關注機台": "、".join(f"{r['設備代號']}({r['狀態判定']})" for _, r in watch.iterrows()),
            "分析與建議": series_comment(p, drm[drm["系列"] == sc]),
        })
    df_s = pd.DataFrame(ser_rows)
    return df_s, df_m


def to_excel(df_s, df_m, month):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        df_s.to_excel(xw, sheet_name="系列分析(合計)", index=False)
        df_m.to_excel(xw, sheet_name="機台明細(個別)", index=False)
        thin = Side(style="thin", color="BFC5CE")
        hdr_fill = PatternFill("solid", fgColor="9BC2E6")
        for ws, df, widths in [
            (xw.sheets["系列分析(合計)"], df_s,
             {"設備名稱": 16, "設備系列": 12, "分析與建議": 60, "需關注機台": 28, "狀態判定": 12}),
            (xw.sheets["機台明細(個別)"], df_m,
             {"設備名稱": 22, "設備代號": 15, "主要故障": 55, "狀態判定": 12}),
        ]:
            for j, col in enumerate(df.columns, start=1):
                c = ws.cell(row=1, column=j)
                c.fill = hdr_fill
                c.font = Font(bold=True)
                c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ws.column_dimensions[get_column_letter(j)].width = widths.get(col, 11)
            ws.row_dimensions[1].height = 34
            st_col = list(df.columns).index("狀態判定") + 1 if "狀態判定" in df.columns else None
            for i in range(2, len(df) + 2):
                for j in range(1, len(df.columns) + 1):
                    c = ws.cell(row=i, column=j)
                    c.border = Border(top=thin, bottom=thin, left=thin, right=thin)
                    wrap = df.columns[j - 1] in ("分析與建議", "主要故障", "需關注機台")
                    c.alignment = Alignment(vertical="center", wrap_text=wrap,
                                            horizontal="left" if wrap else "center")
                if st_col:
                    v = ws.cell(row=i, column=st_col).value
                    if v in STATUS_STYLE:
                        bg, fg = STATUS_STYLE[v]
                        ws.cell(row=i, column=st_col).fill = PatternFill("solid", fgColor=bg[1:])
                        ws.cell(row=i, column=st_col).font = Font(bold=True, color=fg[1:])
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
    return buf.getvalue()
