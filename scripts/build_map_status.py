"""
產生樓層設備位置圖的「儀表板狀態」資料（map_status.json）。
每月更新儀表板資料後執行：python scripts/build_map_status.py
再把 map_status.json 發布到位置圖 artifact（與 index.html 同層）。
"""
import json, os, sys
import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
src = open(os.path.join(ROOT, "app.py"), encoding="utf-8").read()
ns = {"__file__": os.path.join(ROOT, "app.py")}
exec(compile(src.split("# ══════════════════════════════════════════════════════════════════════\ndef main():")[0],
             "app", "exec"), ns)

d = json.load(open(os.path.join(ROOT, "data.json"), encoding="utf-8"))
full = ns["build_full"](pd.DataFrame(d["kpi"]), ns["load_roster"](d), d["hours_base"], d.get("special_off", {}))
full["fm"] = full["機台代碼"].map(ns["fm_key"])
f = full[full["fm"].notna()]
months = sorted(f["月份"].unique(), key=ns["month_key_sort"])
out = {}
for (fm, mo), g in f.groupby(["fm", "月份"]):
    p = ns["pool"](g)
    out.setdefault(fm, {})[mo.replace("月", "")] = [
        p["faults"], p["mtbf"], p["mttr"], round(p["downtime"], 2),
        ns["judge"](p["faults"], p["mtbf"], p["mttr"])]
res = {"upd": d.get("updated", "")[:10], "months": [m.replace("月", "") for m in months], "d": out}
path = os.path.join(ROOT, "map_status.json")
json.dump(res, open(path, "w", encoding="utf-8"), ensure_ascii=False, separators=(",", ":"))
print(f"寫入 {path}：{len(out)} 台、{len(months)} 個月")
