"""
富強醫材 — 設備可靠性分析儀表板
Repo: fortunemed-reliability
"""
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import re, json, base64, requests, io
from datetime import datetime

# ── 常數 ─────────────────────────────────────────────────────────────
GH_REPO   = "babypig600225-maker/fortunemed-reliability"
GH_PATH   = "data.json"
RAW_URL   = f"https://raw.githubusercontent.com/{GH_REPO}/main/{GH_PATH}"
API_URL   = f"https://api.github.com/repos/{GH_REPO}/contents/{GH_PATH}"
THRESHOLD_AVAIL  = 85.0   # 可用率目標 %
THRESHOLD_MTBF   = 100.0  # MTBF 目標 hr
THRESHOLD_MTTR   = 8.0    # MTTR 警戒 hr

MONTH_ORDER = [f"2026-{m}月" for m in
               ["1","2","3","4","5","6","7","8","9","10","11","12"]]

st.set_page_config(page_title="富強醫材｜設備可靠性分析", layout="wide",
                   page_icon="🔧")

# ── 樣式 ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
.metric-card{background:#f8f9fa;border-radius:10px;padding:16px 20px;text-align:center;}
.metric-label{font-size:13px;color:#6c757d;margin-bottom:4px;}
.metric-value{font-size:28px;font-weight:700;color:#212529;}
.metric-sub{font-size:12px;color:#6c757d;margin-top:4px;}
.good{color:#198754 !important;}
.warn{color:#fd7e14 !important;}
.bad {color:#dc3545 !important;}
table{border-collapse:collapse;width:100%;}
th{background:#343a40;color:white;padding:8px 12px;}
td{padding:6px 12px;border-bottom:1px solid #dee2e6;}
tr:hover{background:#f1f3f5;}
</style>
""", unsafe_allow_html=True)

# ── GitHub 讀寫 ───────────────────────────────────────────────────────
@st.cache_data(ttl=60)
def load_data():
    try:
        r = requests.get(RAW_URL, timeout=10)
        if r.status_code == 200:
            return r.json()
    except:
        pass
    return None

def save_data(payload: dict):
    token = st.secrets.get("GITHUB_TOKEN", "")
    if not token:
        return False, "未設定 GITHUB_TOKEN"
    headers = {"Authorization": f"token {token}",
               "Accept": "application/vnd.github.v3+json"}
    content = base64.b64encode(json.dumps(payload, ensure_ascii=False,
                                          indent=2).encode()).decode()
    # 取得現有 sha
    r = requests.get(API_URL, headers=headers, timeout=10)
    sha = r.json().get("sha") if r.status_code == 200 else None
    body = {"message": f"Update data {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "content": content}
    if sha:
        body["sha"] = sha
    r2 = requests.put(API_URL, headers=headers, json=body, timeout=20)
    if r2.status_code in (200, 201):
        return True, "✅ 資料已更新，所有人即時可見！"
    return False, f"❌ 上傳失敗（{r2.status_code}）：{r2.json().get('message','')}"

# ── Excel 解析 ────────────────────────────────────────────────────────
def parse_excel(file_bytes):
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = xl.sheet_names

    # === 時數工作表 ===
    hours_base = {}
    if "時數" in sheets:
        df_h = pd.read_excel(io.BytesIO(file_bytes), sheet_name="時數", header=None)
        # 找月份列（含「月」且第一欄非空）
        for _, row in df_h.iterrows():
            m = str(row.iloc[0])
            if "月" in m and re.match(r"\d+月", m):
                weekdays = int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
                holiday  = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
                key = f"2026-{m}"
                hours_base[key] = weekdays * 11 + holiday * 8

    # 若時數工作表沒讀到，用預設
    defaults = {
        "2026-1月":263,"2026-2月":162,"2026-3月":274,"2026-4月":228,
        "2026-5月":236,"2026-6月":239,"2026-7月":242,"2026-8月":231,
        "2026-9月":220,"2026-10月":217,
    }
    for k, v in defaults.items():
        if k not in hours_base:
            hours_base[k] = v

    # === 報修工作表 ===
    month_sheets = [s for s in sheets if re.match(r"\d{4}-\d+月", s)]
    all_rows = []
    for sh in month_sheets:
        df_raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh, header=None)
        hdr = None
        for i, row in df_raw.iterrows():
            if "seq" in str(row.values).lower():
                hdr = i; break
        if hdr is None:
            continue
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sh, header=hdr)
        df["月份"] = sh
        all_rows.append(df)

    if not all_rows:
        return None, None, "找不到報修工作表"

    df_all = pd.concat(all_rows, ignore_index=True)
    # 只保留設備課
    df_all = df_all[df_all["設備"] == "Y"].copy()

    def extract_code(n):
        if pd.isna(n): return None
        m = re.search(r"M-\d{3}-\d{3}-\d+", str(n))
        return m.group(0) if m else None

    df_all["機台代碼"]  = df_all["設備名稱"].apply(extract_code)
    df_all["維修時數"]  = pd.to_numeric(df_all["維修時數(h)"], errors="coerce").fillna(0).clip(lower=0)
    df_all["停機"]      = df_all["是否停機"].str.upper() == "YES"
    df_all["缺料"]      = df_all["缺料"].astype(str).str.upper().isin(["是","有","YES"])
    df_all["完工超時"]  = df_all["完工-超時"].notna() & (df_all["完工-超時"].astype(str).str.strip() != "")
    df_all["機台名稱"]  = df_all["設備名稱"].apply(
        lambda x: re.sub(r"^M-\d{3}-\d{3}-\d+\s*", "", str(x)) if pd.notna(x) else "")
    df_all["系列"] = df_all["機台代碼"].apply(
        lambda x: x[:7] if isinstance(x, str) else "")

    return df_all, hours_base, None

# ── 計算 KPI ──────────────────────────────────────────────────────────
def calc_kpi(df_all, hours_base, special_off):
    """
    special_off: dict {月份: 特殊停工時數}
    """
    records = []
    for month, grp in df_all.groupby("月份"):
        base   = hours_base.get(month, 220)
        extra  = special_off.get(month, 0)
        avail_h = max(base - extra, 0)

        for code, mg in grp.groupby("機台代碼"):
            if not code: continue
            failures   = len(mg)
            downtime   = mg[mg["停機"]]["維修時數"].sum()
            total_rep  = mg["維修時數"].sum()
            repeat     = 1 if failures >= 2 else 0
            shortage   = mg["缺料"].sum()
            overtime   = mg["完工超時"].sum()

            mttr = round(total_rep / failures, 2) if failures else 0
            op   = max(avail_h - downtime, 0)
            mtbf = round(op / failures, 2) if failures else avail_h
            availability = round(op / avail_h * 100, 1) if avail_h else 100.0

            records.append({
                "月份": month,
                "機台代碼": code,
                "機台名稱": mg["機台名稱"].iloc[0],
                "系列": mg["系列"].iloc[0],
                "故障次數": failures,
                "停機次數": int(mg["停機"].sum()),
                "停工時數(h)": round(downtime, 2),
                "總維修時數(h)": round(total_rep, 2),
                "MTTR(h)": mttr,
                "MTBF(h)": mtbf,
                "可用率(%)": availability,
                "重複故障": repeat,
                "缺料件數": int(shortage),
                "超時完工件數": int(overtime),
                "月可用時數(h)": avail_h,
            })
    return pd.DataFrame(records)

# ── 月別彙總 ──────────────────────────────────────────────────────────
def monthly_summary(df_kpi, df_all, hours_base, special_off):
    rows = []
    for month in sorted(df_kpi["月份"].unique(),
                        key=lambda x: MONTH_ORDER.index(x) if x in MONTH_ORDER else 99):
        mk = df_kpi[df_kpi["月份"] == month]
        ma = df_all[df_all["月份"] == month]
        base   = hours_base.get(month, 220)
        extra  = special_off.get(month, 0)
        avail_h = max(base - extra, 0)
        machines = mk["機台代碼"].nunique()
        total_down = mk["停工時數(h)"].sum()
        rows.append({
            "月份": month,
            "涉及機台數": machines,
            "故障次數": mk["故障次數"].sum(),
            "停工時數(h)": round(total_down, 1),
            "平均MTBF(h)": round(mk["MTBF(h)"].mean(), 1),
            "平均MTTR(h)": round(mk["MTTR(h)"].mean(), 1),
            "平均可用率(%)": round(mk["可用率(%)"].mean(), 1),
            "重複故障台數": mk["重複故障"].sum(),
            "月可用時數(h)": avail_h,
        })
    return pd.DataFrame(rows)

# ══════════════════════════════════════════════════════════════════════
# ── 主程式 ────────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════════════
def main():
    # Header
    col1, col2 = st.columns([3, 1])
    with col1:
        st.title("🔧 富強醫材｜設備可靠性分析")
        st.caption("MTBF · MTTR · 可用率 · 故障分析 · 趨勢")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)

    # ── 管理員上傳區 ──────────────────────────────────────────────────
    stored = load_data()
    with st.sidebar:
        st.header("🔑 管理員更新資料")
        up_file = st.file_uploader("上傳報修 Excel（含時數工作表）",
                                   type=["xlsx"], key="uploader")

        special_off = {}
        if stored and stored.get("special_off"):
            special_off = {k: v for k, v in stored["special_off"].items()}

        with st.expander("📅 特殊停工調整（颱風、旅遊…）"):
            st.caption("填入該月額外未開工的時數，系統自動從可用時數扣除")
            month_opts = [m for m in MONTH_ORDER if "2026" in m]
            for mo in month_opts:
                label = mo.replace("2026-", "")
                val = special_off.get(mo, 0)
                new_val = st.number_input(label, min_value=0, max_value=240,
                                          value=int(val), step=1, key=f"sp_{mo}")
                if new_val:
                    special_off[mo] = new_val
                elif mo in special_off:
                    del special_off[mo]

        if up_file:
            file_bytes = up_file.read()
            with st.spinner("解析中…"):
                df_all_new, hours_base_new, err = parse_excel(file_bytes)
            if err:
                st.error(err)
            else:
                df_kpi_new = calc_kpi(df_all_new, hours_base_new, special_off)
                # 只選實際存在的欄位
                raw_cols = ["月份","機台代碼","機台名稱","系列","維修時數","停機","缺料","完工超時"]
                for col in ["需求","部門","設備名稱"]:
                    if col in df_all_new.columns:
                        raw_cols.append(col)
                payload = {
                    "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "hours_base": hours_base_new,
                    "special_off": special_off,
                    "kpi": df_kpi_new.to_dict(orient="records"),
                    "raw": df_all_new[raw_cols].to_dict(orient="records"),
                }
                ok, msg = save_data(payload)
                st.info(msg)
                if ok:
                    st.cache_data.clear()
                    stored = payload

        elif stored:
            st.success(f"資料更新時間：{stored.get('updated','—')}")
            # 允許只更新特殊停工
            if st.button("更新特殊停工設定"):
                stored["special_off"] = special_off
                ok, msg = save_data(stored)
                st.info(msg)
                if ok:
                    st.cache_data.clear()

    # ── 無資料畫面 ────────────────────────────────────────────────────
    if not stored:
        st.markdown("""
        <div style="text-align:center;padding:80px 0;">
        <div style="font-size:60px;">📊</div>
        <h2 style="color:#6c757d;">尚無資料</h2>
        <p style="color:#adb5bd;">請由左側上傳報修 Excel 檔案</p>
        </div>""", unsafe_allow_html=True)
        return

    # 重建 DataFrame
    df_kpi  = pd.DataFrame(stored["kpi"])
    df_raw  = pd.DataFrame(stored["raw"])
    hours_base  = stored.get("hours_base", {})
    special_off = stored.get("special_off", {})

    months_avail = sorted(df_kpi["月份"].unique(),
                          key=lambda x: MONTH_ORDER.index(x) if x in MONTH_ORDER else 99)
    df_summary = monthly_summary(df_kpi, df_raw, hours_base, special_off)

    # ── 全局篩選器 ────────────────────────────────────────────────────
    st.markdown("---")
    fc1, fc2, fc3 = st.columns(3)
    with fc1:
        sel_months = st.multiselect("月份", months_avail, default=months_avail,
                                    key="f_months")
    with fc2:
        all_series = sorted(df_kpi["系列"].unique())
        sel_series = st.multiselect("機台系列", all_series, default=[], key="f_series")
    with fc3:
        only_stop = st.checkbox("只顯示有停機記錄的機台", key="f_stop")

    # 套用篩選
    dk = df_kpi[df_kpi["月份"].isin(sel_months)].copy()
    dr = df_raw[df_raw["月份"].isin(sel_months)].copy()
    if sel_series:
        dk = dk[dk["系列"].isin(sel_series)]
        dr = dr[dr["系列"].isin(sel_series)]
    if only_stop:
        dk = dk[dk["停機次數"] > 0]

    # ══ 分頁 ═══════════════════════════════════════════════════════════
    tab1, tab2, tab3, tab4 = st.tabs(
        ["📊 月別總覽", "🔧 機台明細", "⚠️ 故障分析", "📈 趨勢分析"])

    # ══ Tab1：月別總覽 ═════════════════════════════════════════════════
    with tab1:
        ds = df_summary[df_summary["月份"].isin(sel_months)]

        # KPI 卡片
        c1,c2,c3,c4,c5 = st.columns(5)
        def kpi_card(col, label, value, unit, good_fn=None, sub=""):
            cls = ""
            if good_fn:
                cls = "good" if good_fn(value) else "bad"
            col.markdown(f"""
            <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value {cls}">{value}{unit}</div>
            <div class="metric-sub">{sub}</div>
            </div>""", unsafe_allow_html=True)

        kpi_card(c1, "總故障次數",   int(ds["故障次數"].sum()), " 次", sub=f"{len(sel_months)}個月累計")
        kpi_card(c2, "總停工時數",   round(ds["停工時數(h)"].sum(),1), " h", sub="停機事件")
        kpi_card(c3, "平均MTBF",     round(ds["平均MTBF(h)"].mean(),1), " h",
                 lambda v: v >= THRESHOLD_MTBF, sub=f"目標≥{THRESHOLD_MTBF}h")
        kpi_card(c4, "平均MTTR",     round(ds["平均MTTR(h)"].mean(),1), " h",
                 lambda v: v <= THRESHOLD_MTTR, sub=f"警戒≤{THRESHOLD_MTTR}h")
        kpi_card(c5, "平均可用率",   round(ds["平均可用率(%)"].mean(),1), "%",
                 lambda v: v >= THRESHOLD_AVAIL, sub=f"目標≥{THRESHOLD_AVAIL}%")

        st.markdown("<br>", unsafe_allow_html=True)

        # 月別摘要表
        st.subheader("月別摘要")
        def style_avail(val):
            if isinstance(val, float):
                if val >= THRESHOLD_AVAIL: return "color:green;font-weight:bold"
                elif val >= 80: return "color:orange"
                return "color:red;font-weight:bold"
            return ""

        st.dataframe(
            ds.style.map(style_avail, subset=["平均可用率(%)"]),
            use_container_width=True, hide_index=True)

        # 月別故障+停工雙軸圖
        st.subheader("月別故障次數 vs 停工時數")
        fig = go.Figure()
        fig.add_bar(x=ds["月份"], y=ds["故障次數"], name="故障次數",
                    marker_color="#4dabf7", yaxis="y")
        fig.add_scatter(x=ds["月份"], y=ds["停工時數(h)"], name="停工時數(h)",
                        mode="lines+markers", marker_color="#f03e3e", yaxis="y2")
        fig.update_layout(
            yaxis=dict(title="故障次數"),
            yaxis2=dict(title="停工時數(h)", overlaying="y", side="right"),
            legend=dict(x=0, y=1.1, orientation="h"),
            height=350, margin=dict(t=20,b=40))
        st.plotly_chart(fig, use_container_width=True)

    # ══ Tab2：機台明細 ═════════════════════════════════════════════════
    with tab2:
        st.subheader("機台明細 — MTBF / MTTR / 可用率")

        # 機台彙總（跨月加總/平均）
        agg = dk.groupby(["機台代碼","機台名稱","系列"]).agg(
            故障次數=("故障次數","sum"),
            停機次數=("停機次數","sum"),
            停工時數=("停工時數(h)","sum"),
            總維修時數=("總維修時數(h)","sum"),
            平均MTBF=("MTBF(h)","mean"),
            平均MTTR=("MTTR(h)","mean"),
            平均可用率=("可用率(%)","mean"),
            重複故障月數=("重複故障","sum"),
        ).reset_index().round(2)
        agg = agg.sort_values("停工時數", ascending=False)

        def color_avail(val):
            if val >= THRESHOLD_AVAIL: return "background-color:#d3f9d8"
            elif val >= 80: return "background-color:#fff3bf"
            return "background-color:#ffe3e3"

        st.dataframe(
            agg.style.map(color_avail, subset=["平均可用率"]),
            use_container_width=True, hide_index=True)

        # 月份明細下鑽
        st.markdown("---")
        sel_m = st.selectbox("查看單台月別明細", ["— 選擇機台 —"] + list(agg["機台代碼"]))
        if sel_m != "— 選擇機台 —":
            dm = dk[dk["機台代碼"] == sel_m].sort_values(
                "月份", key=lambda x: x.map(
                    {m:i for i,m in enumerate(MONTH_ORDER)}))
            st.dataframe(dm[["月份","故障次數","停機次數","停工時數(h)",
                             "MTBF(h)","MTTR(h)","可用率(%)"]],
                         use_container_width=True, hide_index=True)

    # ══ Tab3：故障分析 ═════════════════════════════════════════════════
    with tab3:
        c1, c2 = st.columns(2)

        # Top 10 停工
        with c1:
            st.subheader("⬛ Top 10 停工時數（機台）")
            top10 = dk.groupby(["機台代碼","機台名稱"])["停工時數(h)"].sum()\
                      .reset_index().nlargest(10,"停工時數(h)")
            fig2 = px.bar(top10, x="停工時數(h)", y="機台代碼",
                          orientation="h", text="停工時數(h)",
                          color="停工時數(h)", color_continuous_scale="Reds",
                          hover_data=["機台名稱"])
            fig2.update_layout(height=360, margin=dict(t=10,b=10),
                               coloraxis_showscale=False,
                               yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig2, use_container_width=True)

        # Top 10 故障次數
        with c2:
            st.subheader("🔁 Top 10 故障次數（機台）")
            top10f = dk.groupby(["機台代碼","機台名稱"])["故障次數"].sum()\
                       .reset_index().nlargest(10,"故障次數")
            fig3 = px.bar(top10f, x="故障次數", y="機台代碼",
                          orientation="h", text="故障次數",
                          color="故障次數", color_continuous_scale="Blues",
                          hover_data=["機台名稱"])
            fig3.update_layout(height=360, margin=dict(t=10,b=10),
                               coloraxis_showscale=False,
                               yaxis=dict(autorange="reversed"))
            st.plotly_chart(fig3, use_container_width=True)

        # 重複故障機台
        st.markdown("---")
        st.subheader("⚠️ 重複故障機台（同月 ≥ 2 次）")
        repeat = dk[dk["重複故障"] == 1].groupby(["機台代碼","機台名稱"])\
                   .agg(發生月數=("月份","nunique"),
                        累計故障次數=("故障次數","sum")).reset_index()
        repeat = repeat.sort_values("發生月數", ascending=False)
        st.dataframe(repeat, use_container_width=True, hide_index=True)

        # 故障原因文字分析
        st.markdown("---")
        st.subheader("📋 故障原因分類（需求欄關鍵字）")
        if "需求" in dr.columns:
            keywords = {
                "溫度異常": ["溫度","過熱","溫控"],
                "漏油/漏水": ["漏油","漏水","滲油"],
                "電氣/電線": ["電線","電路","跳電","斷電","短路","感溫線"],
                "機械磨損": ["斷裂","磨損","螺絲","異音","鬆脫"],
                "感測/控制": ["感應","感測","PLC","電磁閥"],
                "無法啟動": ["無法啟動","無法使用","無法運作"],
                "保養/更換": ["保養","更換","清洗"],
                "其他": [],
            }
            cat_counts = {}
            matched = set()
            for cat, kws in keywords.items():
                if cat == "其他": continue
                mask = dr["需求"].astype(str).apply(
                    lambda x: any(k in x for k in kws))
                idxs = dr[mask].index.tolist()
                cat_counts[cat] = len([i for i in idxs if i not in matched])
                matched.update(idxs)
            cat_counts["其他"] = len(dr) - sum(cat_counts.values())

            df_cat = pd.DataFrame(list(cat_counts.items()),
                                  columns=["故障類別","件數"])\
                       .sort_values("件數", ascending=False)
            fig4 = px.pie(df_cat, names="故障類別", values="件數",
                          color_discrete_sequence=px.colors.qualitative.Set2)
            fig4.update_layout(height=350, margin=dict(t=10,b=10))
            st.plotly_chart(fig4, use_container_width=True)

        # 缺料統計
        st.markdown("---")
        st.subheader("📦 缺料分析")
        shortage_m = dk.groupby("月份")["缺料件數"].sum().reset_index()
        shortage_m = shortage_m.sort_values("月份", key=lambda x: x.map(
            {m:i for i,m in enumerate(MONTH_ORDER)}))
        fig5 = px.bar(shortage_m, x="月份", y="缺料件數",
                      text="缺料件數", color_discrete_sequence=["#fab005"])
        fig5.update_layout(height=280, margin=dict(t=10,b=10))
        st.plotly_chart(fig5, use_container_width=True)

    # ══ Tab4：趨勢分析 ════════════════════════════════════════════════
    with tab4:
        ds2 = df_summary[df_summary["月份"].isin(sel_months)].sort_values(
            "月份", key=lambda x: x.map({m:i for i,m in enumerate(MONTH_ORDER)}))

        # MTBF/MTTR 趨勢
        st.subheader("MTBF / MTTR 月別趨勢")
        fig6 = go.Figure()
        fig6.add_scatter(x=ds2["月份"], y=ds2["平均MTBF(h)"],
                         name="平均MTBF(h)", mode="lines+markers",
                         line=dict(color="#339af0", width=2))
        fig6.add_scatter(x=ds2["月份"], y=ds2["平均MTTR(h)"],
                         name="平均MTTR(h)", mode="lines+markers",
                         line=dict(color="#f03e3e", width=2), yaxis="y2")
        fig6.add_hline(y=THRESHOLD_MTBF, line_dash="dash",
                       line_color="#339af0", annotation_text=f"MTBF目標{THRESHOLD_MTBF}h")
        fig6.add_hline(y=THRESHOLD_MTTR, line_dash="dash",
                       line_color="#f03e3e", annotation_text=f"MTTR警戒{THRESHOLD_MTTR}h",
                       yref="y2")
        fig6.update_layout(
            yaxis=dict(title="MTBF(h)"),
            yaxis2=dict(title="MTTR(h)", overlaying="y", side="right"),
            height=350, margin=dict(t=30,b=40),
            legend=dict(x=0, y=1.1, orientation="h"))
        st.plotly_chart(fig6, use_container_width=True)

        # 可用率趨勢
        st.subheader("平均可用率趨勢")
        fig7 = px.line(ds2, x="月份", y="平均可用率(%)",
                       markers=True, color_discrete_sequence=["#2f9e44"])
        fig7.add_hline(y=THRESHOLD_AVAIL, line_dash="dash",
                       line_color="orange",
                       annotation_text=f"目標{THRESHOLD_AVAIL}%")
        fig7.update_layout(height=300, margin=dict(t=10,b=40))
        st.plotly_chart(fig7, use_container_width=True)

        # 系列別可用率熱圖
        st.subheader("機台系列 × 月份 — 可用率熱圖")
        pivot = dk.pivot_table(values="可用率(%)",
                               index="系列", columns="月份", aggfunc="mean").round(1)
        # 排序欄位
        ordered_cols = [c for c in MONTH_ORDER if c in pivot.columns]
        pivot = pivot[ordered_cols]
        fig8 = px.imshow(pivot, text_auto=True, aspect="auto",
                         color_continuous_scale=["#f03e3e","#fff3bf","#2f9e44"],
                         zmin=70, zmax=100)
        fig8.update_layout(height=400, margin=dict(t=10,b=10))
        st.plotly_chart(fig8, use_container_width=True)

    # Footer
    st.markdown("---")
    st.caption(f"富強醫材股份有限公司 設備課｜資料更新：{stored.get('updated','—')}")

if __name__ == "__main__":
    main()
