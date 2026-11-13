"""
目標設定說明、每月達成情形、每月分析摘要
"""
import pandas as pd


def _ok(v, target, higher_better=True):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    return v >= target if higher_better else v <= target


def unit_table(dkm, by, pool):
    """單位（系列 or 機台）× 單月的合計 KPI"""
    rows = []
    for key, g in dkm.groupby(by):
        p = pool(g)
        rows.append({"key": key, "name": g["機台名稱"].iloc[0] if by == "機台代碼" else "",
                     **p})
    return pd.DataFrame(rows)


def month_achievement(dkm, pool, T):
    """單月達成情形；T = dict(mtbf, mttr, avail)。無故障的單位視為 MTBF/MTTR 達標。"""
    fleet = pool(dkm)
    ser = unit_table(dkm, "系列", pool)
    mac = unit_table(dkm, "機台代碼", pool)

    def rate(df, col, target, higher):
        if df.empty:
            return 0, 0
        ok = df[col].apply(lambda v: True if v is None or pd.isna(v)
                           else (v >= target if higher else v <= target))
        return int(ok.sum()), len(df)

    return {
        "fleet": fleet, "series": ser, "machines": mac,
        "ser_mtbf": rate(ser, "mtbf", T["mtbf"], True),
        "ser_mttr": rate(ser, "mttr", T["mttr"], False),
        "ser_avail": rate(ser, "avail", T["avail"], True),
        "mac_mtbf": rate(mac, "mtbf", T["mtbf"], True),
        "mac_mttr": rate(mac, "mttr", T["mttr"], False),
        "fleet_mttr_ok": _ok(fleet["mttr"], T["mttr"], False),
        "fleet_avail_ok": _ok(fleet["avail"], T["avail"], True),
    }


def achievement_table(dk, months, pool, T):
    rows, cache = [], {}
    for mo in months:
        a = month_achievement(dk[dk["月份"] == mo], pool, T)
        cache[mo] = a
        f = a["fleet"]
        fr = lambda t: f"{t[0]}/{t[1]}（{t[0]/t[1]*100:.0f}%）" if t[1] else "—"
        mark = lambda ok: "—" if ok is None else ("✓" if ok else "✗")
        rows.append({
            "月份": mo.replace("2026-", ""),
            "在用台數": f["n"], "故障次數": f["faults"],
            "全廠MTTR(h)": f["mttr"], "MTTR達標": mark(a["fleet_mttr_ok"]),
            "全廠可用率(%)": f["avail"], "可用率達標": mark(a["fleet_avail_ok"]),
            "全廠MTBF(h)（參考）": f["mtbf"],
            "每台平均故障次數（參考）": round(f["faults"] / f["n"], 3) if f["n"] else None,
            "系列MTBF達成": fr(a["ser_mtbf"]), "系列MTTR達成": fr(a["ser_mttr"]),
            "系列可用率達成": fr(a["ser_avail"]),
            "單台MTBF達成": fr(a["mac_mtbf"]), "單台MTTR達成": fr(a["mac_mttr"]),
            "_ser_mtbf": a["ser_mtbf"][0] / a["ser_mtbf"][1] * 100 if a["ser_mtbf"][1] else None,
            "_ser_mttr": a["ser_mttr"][0] / a["ser_mttr"][1] * 100 if a["ser_mttr"][1] else None,
            "_ser_avail": a["ser_avail"][0] / a["ser_avail"][1] * 100 if a["ser_avail"][1] else None,
            "_mac_mtbf": a["mac_mtbf"][0] / a["mac_mtbf"][1] * 100 if a["mac_mtbf"][1] else None,
            "_mac_mttr": a["mac_mttr"][0] / a["mac_mttr"][1] * 100 if a["mac_mttr"][1] else None,
        })
    return pd.DataFrame(rows), cache


def _diff(cur, prev, unit="", fmt="{:+.2f}", better_low=True):
    if cur is None or prev is None:
        return ""
    d = cur - prev
    if abs(d) < 1e-9:
        return "，與上月持平"
    word = ("改善" if (d < 0) == better_low else "變差")
    return f"（較上月 {fmt.format(d)}{unit}，{word}）"


def summary_text(mo, prev_mo, cache, dr, series_name, T):
    """回傳 list[(標題, 內容)]"""
    a = cache[mo]
    p = cache.get(prev_mo) if prev_mo else None
    f = a["fleet"]
    pf = p["fleet"] if p else None
    m_label = mo.replace("2026-", "")
    out = []

    # 1. 全廠
    t = f"在用 {f['n']} 台，故障 {f['faults']} 次"
    if pf:
        d = f["faults"] - pf["faults"]
        t += f"（較上月 {d:+d} 次）"
    t += f"，每台平均 {f['faults']/f['n']:.2f} 次。" if f["n"] else "。"
    if f["mttr"] is not None:
        t += (f"MTTR {f['mttr']:.2f}h，{'達標' if f['mttr'] <= T['mttr'] else '未達標'}（≤{T['mttr']:g}h）"
              + _diff(f["mttr"], pf["mttr"] if pf else None, "h") + "。")
    if f["avail"] is not None:
        t += (f"可用率 {f['avail']:.2f}%，{'達標' if f['avail'] >= T['avail'] else '未達標'}（≥{T['avail']:g}%）"
              + _diff(f["avail"], pf["avail"] if pf else None, "%", better_low=False) + "。")
    if f["mtbf"] is not None:
        t += f"MTBF {f['mtbf']:.1f}h（參考值）。"
    out.append(("全廠", t))

    # 2. 系列
    ser = a["series"]
    bad_mtbf = ser[ser["mtbf"].notna() & (ser["mtbf"] < T["mtbf"])].sort_values("mtbf")
    bad_mttr = ser[ser["mttr"].notna() & (ser["mttr"] > T["mttr"])].sort_values("mttr", ascending=False)
    bad_av = ser[ser["avail"].notna() & (ser["avail"] < T["avail"])].sort_values("avail")
    nm = lambda k: f"{series_name.get(k, k)}（{k}）"
    parts = [f"共 {len(ser)} 個系列。MTBF 達成 {a['ser_mtbf'][0]}、MTTR 達成 {a['ser_mttr'][0]}、"
             f"可用率達成 {a['ser_avail'][0]}。"]
    if len(bad_mtbf):
        parts.append("MTBF 未達：" + "、".join(f"{nm(r.key)} {r.mtbf:.1f}h" for r in bad_mtbf.head(3).itertuples())
                     + ("等" if len(bad_mtbf) > 3 else "") + "。")
    if len(bad_mttr):
        parts.append("MTTR 超標：" + "、".join(f"{nm(r.key)} {r.mttr:.1f}h" for r in bad_mttr.head(3).itertuples())
                     + ("等" if len(bad_mttr) > 3 else "") + "。")
    if len(bad_av):
        parts.append("可用率未達：" + "、".join(f"{nm(r.key)} {r.avail:.2f}%" for r in bad_av.head(3).itertuples())
                     + ("等" if len(bad_av) > 3 else "") + "。")
    if not (len(bad_mtbf) or len(bad_mttr) or len(bad_av)):
        parts.append("各系列皆達標。")
    out.append(("系列", "".join(parts)))

    # 3. 單台
    mac = a["machines"]
    bm = mac[mac["mtbf"].notna() & (mac["mtbf"] < T["mtbf"])].sort_values("mtbf")
    bt = mac[mac["mttr"].notna() & (mac["mttr"] > T["mttr"])].sort_values("mttr", ascending=False)
    parts = [f"MTBF 未達 {len(bm)} 台、MTTR 超標 {len(bt)} 台（共 {len(mac)} 台）。"]
    if len(bm):
        parts.append("故障最頻繁：" + "、".join(f"{r.key} {r.name}（{r.faults}次）"
                                           for r in bm.sort_values("faults", ascending=False).head(3).itertuples()) + "。")
    if len(bt):
        parts.append("修復最久：" + "、".join(f"{r.key} {r.name}（MTTR {r.mttr:.1f}h）"
                                         for r in bt.head(3).itertuples()) + "。")
    out.append(("單台", "".join(parts)))

    # 4. 重點事件
    drm = dr[dr["月份"] == mo]
    if len(drm):
        r = drm.sort_values(["停工時數(h)", "維修時數"], ascending=False).iloc[0]
        need = str(r.get("需求", "") or "").strip()
        sq = r.get("seq")
        sq = "" if sq is None or (isinstance(sq, float) and pd.isna(sq)) else f"，seq={int(float(sq))}"
        t = (f"影響最大：{r['機台代碼']}「{need[:24]}」（停機 {float(r['停工時數(h)']):g}h、"
             f"維修 {float(r['維修時數']):g}h{sq}）。")
        rep = drm.groupby("機台代碼").size()
        rep = rep[rep >= 2].sort_values(ascending=False)
        if len(rep):
            t += "重複故障：" + "、".join(f"{c}（{n}次）" for c, n in rep.head(3).items()) + "。"
        out.append(("重點事件", t))

    # 5. 建議
    sug = []
    if f["mttr"] is not None and f["mttr"] > T["mttr"]:
        sug.append("全廠 MTTR 未達標：優先檢討修復最久的機台，確認是否為等料、委外或技術問題，並建立備品清單")
    elif len(bad_mttr):
        sug.append("MTTR 超標的系列：檢討修復時間長的原因（等料、委外、技術），建立常用備品與標準修復步驟")
    if len(bad_av):
        sug.append("可用率未達的系列：檢視停機原因，評估縮短停機的措施（備機、快速換件）")
    if len(drm) and len(drm.groupby("機台代碼").size().pipe(lambda s: s[s >= 2])):
        sug.append("重複故障機台：進行根本原因分析，並納入預防保養項目")
    if len(bm) and not sug:
        sug.append("MTBF 未達的機台：追蹤故障原因，必要時增加點檢頻率")
    if not sug:
        sug.append("本月各項目標皆達成，持續維持現行保養與點檢")
    out.append(("建議", "；".join(sug) + "。"))
    return out
