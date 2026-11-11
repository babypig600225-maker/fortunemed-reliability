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
THRESHOLD_AVAIL  = 85.0
THRESHOLD_MTBF   = 100.0
THRESHOLD_MTTR   = 8.0

MONTH_ORDER = [f"2026-{m}月" for m in
               ["1","2","3","4","5","6","7","8","9","10","11","12"]]

EQ_TYPE_MAP = {"301": "成型類", "302": "加工類", "303": "包裝類"}
EQ_TYPE_ORDER = ["成型類", "加工類", "包裝類"]

# 設備流水碼 → 設備名稱對照表（來源：設備清單 Excel）
SERIES_NAME = {
    "M-301-001": "混料機",
    "M-301-002": "送料機",
    "M-301-003": "抽料機",
    "M-301-004": "加硫機",
    "M-301-005": "押出機",
    "M-301-006": "雷射測定器",
    "M-301-007": "自動外徑控制器",
    "M-301-008": "自動裁切機",
    "M-301-009": "單機型壓鑄成型機",
    "M-301-010": "射出成型機",
    "M-301-011": "液態矽膠射出成型機",
    "M-301-012": "無廢料後射出成型機",
    "M-301-018": "烤箱",
    "M-301-019": "冰水機",
    "M-301-025": "無廢料後射出成型機",
    "M-301-026": "噴塗淋球機",
    "M-301-027": "雷射控制器",
    "M-301-028": "2KM供料機",
    "M-301-030": "模溫機",
    "M-301-031": "塗膜機",
    "M-301-033": "打碎機",
    "M-301-034": "放卷機",
    "M-302-002": "拉力測試機",
    "M-302-004": "試球機",
    "M-302-005": "桌上型鑽床",
    "M-302-008": "磨R角機",
    "M-302-010": "TIP接著機(油壓)",
    "M-302-011": "吸球吸力測試機",
    "M-302-012": "流量機",
    "M-302-014": "網印機",
    "M-302-016": "T3印刷加硫機",
    "M-302-018": "超音波塑膠熔接機",
    "M-302-019": "黏球機",
    "M-302-021": "測漏機",
    "M-302-023": "封箱機",
    "M-302-026": "CV接著機",
    "M-302-027": "CV接著機",
    "M-302-028": "活塞色套機",
    "M-302-029": "自動沖床機",
    "M-302-030": "氣動沖壓機",
    "M-302-031": "C型油壓壓床",
    "M-302-039": "T3印刷加硫機",
    "M-302-041": "點膠機",
    "M-302-042": "磨R角機",
    "M-302-044": "墨杯式移印機",
    "M-302-045": "吸球組裝機",
    "M-302-046": "攪拌機",
    "M-302-047": "C型油壓壓床",
    "M-302-048": "點膠機",
    "M-302-049": "封閉式刮痰管測試機",
    "M-302-053": "引流袋焊接測漏設備",
    "M-302-055": "電烤箱",
    "M-303-001": "消毒袋封口機",
    "M-303-002": "腳踏/手壓封口機",
    "M-303-003": "貼標機",
    "M-303-007": "條碼機",
    "M-303-008": "超音波清洗機",
    "M-303-009": "腳踏/手壓封口機",
    "M-303-010": "啟德包裝機",
    "M-303-011": "打包機",
    "M-303-012": "小型BI培養鍋",
    "M-303-013": "EO負壓滅菌鍋",
    "M-303-017": "燙金機",
    "M-303-018": "收縮機",
    "M-303-020": "預熱室",
    "M-303-021": "連續性封口機",
    "M-303-022": "凝膠充填機",
    "M-303-023": "鎖蓋機",
    "M-303-025": "泡殼機",
    "M-303-028": "密著包裝機",
    "M-303-029": "移動式伸縮包膜機",
}

CHART_COLORS = ['#378ADD','#1D9E75','#BA7517','#D85A30','#7F77DD','#D4537E',
                '#2CA6A4','#E5A234','#5C85D6','#B84DB8']

PLOT_LAYOUT = dict(
    plot_bgcolor='white', paper_bgcolor='white',
    font=dict(family="sans-serif", size=12, color="#1a1a18"),
    margin=dict(t=48, b=40, l=48, r=16),
)
GRID = dict(gridcolor='#eae8e1', linecolor='#eae8e1')

st.set_page_config(
    page_title="富強醫材｜設備可靠性分析",
    page_icon="🔧", layout="wide",
    initial_sidebar_state="collapsed"
)

# ── 樣式 ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* 隱藏 Streamlit 預設 header / footer / menu */
header[data-testid="stHeader"] { display: none !important; }
#MainMenu { visibility: hidden !important; }
footer { visibility: hidden !important; }

/* 整體 */
.main { padding: 0 1.5rem; }
.block-container { padding-top: 1rem !important; max-width:1400px; }

/* 頂部標題 banner */
.page-banner {
    background: linear-gradient(135deg, #1A3A5C 0%, #2E6DA4 100%);
    color: white;
    padding: 18px 28px;
    border-radius: 10px;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.page-banner-title {
    font-size: 22px;
    font-weight: 700;
    letter-spacing: 1px;
    white-space: normal;
    word-break: keep-all;
}
.page-banner-sub {
    font-size: 12px;
    opacity: 0.80;
    margin-top: 4px;
}

/* KPI 卡片 */
div[data-testid="metric-container"] {
    background: #f5f4f0;
    border-radius: 10px;
    padding: 14px 16px !important;
    border-left: 4px solid #2E6DA4;
}
div[data-testid="metric-container"] label {
    font-size: 11px !important;
    color: #888780 !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
}
div[data-testid="metric-container"] [data-testid="stMetricValue"] {
    font-size: 22px !important;
    font-weight: 700 !important;
    color: #1A3A5C !important;
}

/* 設備類型標題 */
.type-header {
    font-size: 16px;
    font-weight: 700;
    padding: 8px 14px;
    border-radius: 6px;
    margin: 20px 0 8px 0;
    color: white;
}

/* 設備流水碼小標題 */
.series-label {
    font-size: 13px;
    font-weight: 600;
    color: #1a1a18;
    margin: 4px 0 8px 0;
}

/* 區塊標題 */
.section-title {
    font-size: 13px;
    font-weight: 600;
    color: #2E6DA4;
    border-bottom: 1px solid #dde3ea;
    padding-bottom: 4px;
    margin: 12px 0 8px 0;
}

/* 徽章 */
.badge {
    display:inline-block; padding:2px 10px; border-radius:4px;
    font-size:12px; font-weight:600;
}
.badge-green { background:#e2efda; color:#1f7a4d; }
.badge-red   { background:#fcebeb; color:#a32d2d; }
.badge-amber { background:#fff2cc; color:#7b4f00; }
.badge-blue  { background:#e6f1fb; color:#185fa5; }

/* 表格 */
.dataframe { font-size:12px !important; }
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

    hours_base = {}
    if "時數" in sheets:
        df_h = pd.read_excel(io.BytesIO(file_bytes), sheet_name="時數", header=None)
        for _, row in df_h.iterrows():
            m = str(row.iloc[0])
            if "月" in m and re.match(r"\d+月", m):
                weekdays = int(row.iloc[1]) if pd.notna(row.iloc[1]) else 0
                holiday  = int(row.iloc[2]) if pd.notna(row.iloc[2]) else 0
                key = f"2026-{m}"
                hours_base[key] = weekdays * 11  # 假日維修部門不上班，不計入

    defaults = {
        "2026-1月":263,"2026-2月":162,"2026-3月":274,"2026-4月":228,
        "2026-5月":236,"2026-6月":239,"2026-7月":242,"2026-8月":231,
        "2026-9月":220,"2026-10月":217,
    }
    for k, v in defaults.items():
        if k not in hours_base:
            hours_base[k] = v

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
    df_all = df_all[df_all["設備"] == "Y"].copy()

    # 機台代碼：取 M-XXX-YYY-D[-穴] 完整代碼（如 M-302-010-1-1）
    def extract_code(n):
        if pd.isna(n): return None
        m = re.search(r"M-\d{3}-\d{3}-\d+(?:-\d+)*", str(n))
        return m.group(0) if m else None

    df_all["機台代碼"] = df_all["設備名稱"].apply(extract_code)
    df_all["維修時數"] = pd.to_numeric(df_all["維修時數(h)"], errors="coerce").fillna(0).clip(lower=0)
    df_all["停機"]    = df_all["是否停機"].str.upper() == "YES"
    df_all["缺料"]    = df_all["缺料"].astype(str).str.upper().isin(["是","有","YES"])
    df_all["完工超時"] = df_all["完工-超時"].notna() & (df_all["完工-超時"].astype(str).str.strip() != "")
    df_all["機台名稱"] = df_all["設備名稱"].apply(
        lambda x: re.sub(r"^M-\d{3}-\d{3}-\d+(?:-\d+)*\s*", "", str(x)) if pd.notna(x) else "")

    # 系列 = M-XXX-YYY（設備類別，如 M-302-010）
    def extract_series(code):
        if not isinstance(code, str): return ""
        m = re.match(r"(M-\d{3}-\d{3})", code)
        return m.group(1) if m else code

    # 設備類型：M-301=成型類、M-302=加工類、M-303=包裝類
    def extract_type(code):
        if not isinstance(code, str): return "其他"
        m = re.match(r"M-(\d{3})", code)
        if not m: return "其他"
        return EQ_TYPE_MAP.get(m.group(1), f"M-{m.group(1)}類")

    df_all["系列"] = df_all["機台代碼"].apply(extract_series)
    df_all["設備類型"] = df_all["機台代碼"].apply(extract_type)

    return df_all, hours_base, None

# ── 計算 KPI ──────────────────────────────────────────────────────────
def calc_kpi(df_all, hours_base, special_off):
    records = []
    for month, grp in df_all.groupby("月份"):
        base    = hours_base.get(month, 220)
        extra   = special_off.get(month, 0)
        avail_h = max(base - extra, 0)

        for code, mg in grp.groupby("機台代碼"):
            if not code: continue
            failures  = len(mg)
            downtime  = mg[mg["停機"]]["維修時數"].sum()
            total_rep = mg["維修時數"].sum()
            repeat    = 1 if failures >= 2 else 0
            shortage  = mg["缺料"].sum()
            overtime  = mg["完工超時"].sum()

            mttr = round(total_rep / failures, 2) if failures else 0
            op   = max(avail_h - downtime, 0)
            mtbf = round(op / failures, 2) if failures else avail_h
            availability = round(op / avail_h * 100, 2) if avail_h else 100.0

            records.append({
                "月份": month, "機台代碼": code,
                "機台名稱": mg["機台名稱"].iloc[0],
                "系列": mg["系列"].iloc[0],
                "設備類型": mg["設備類型"].iloc[0],
                "故障次數": failures,
                "停機次數": int(mg["停機"].sum()),
                "停工時數(h)": round(downtime, 2),
                "總維修時數(h)": round(total_rep, 2),
                "MTTR(h)": mttr, "MTBF(h)": mtbf,
                "可用率(%)": availability,
                "重複故障": repeat,
                "缺料件數": int(shortage),
                "超時完工件數": int(overtime),
                "月可用時數(h)": avail_h,
            })
    return pd.DataFrame(records)

def monthly_summary(df_kpi, hours_base, special_off):
    rows = []
    for month in sorted(df_kpi["月份"].unique(),
                        key=lambda x: MONTH_ORDER.index(x) if x in MONTH_ORDER else 99):
        mk = df_kpi[df_kpi["月份"] == month]
        base    = hours_base.get(month, 220)
        extra   = special_off.get(month, 0)
        avail_h = max(base - extra, 0)
        rows.append({
            "月份": month,
            "涉及機台數": mk["機台代碼"].nunique(),
            "故障次數": int(mk["故障次數"].sum()),
            "停工時數(h)": round(mk["停工時數(h)"].sum(), 2),
            "平均MTBF(h)": round(mk["MTBF(h)"].mean(), 2),
            "平均MTTR(h)": round(mk["MTTR(h)"].mean(), 2),
            "平均可用率(%)": round(mk["可用率(%)"].mean(), 2),
            "重複故障台數": int(mk["重複故障"].sum()),
            "月可用時數(h)": avail_h,
        })
    return pd.DataFrame(rows)

def monthly_summary_by_type(df_kpi, hours_base, special_off):
    """各設備類型 × 月份彙總"""
    rows = []
    for eq_type in EQ_TYPE_ORDER:
        sub = df_kpi[df_kpi["設備類型"] == eq_type]
        if sub.empty:
            continue
        for month in sorted(sub["月份"].unique(),
                            key=lambda x: MONTH_ORDER.index(x) if x in MONTH_ORDER else 99):
            mk = sub[sub["月份"] == month]
            rows.append({
                "設備類型": eq_type,
                "月份": month,
                "涉及機台數": mk["機台代碼"].nunique(),
                "故障次數": int(mk["故障次數"].sum()),
                "停工時數(h)": round(mk["停工時數(h)"].sum(), 2),
                "平均MTBF(h)": round(mk["MTBF(h)"].mean(), 2),
                "平均MTTR(h)": round(mk["MTTR(h)"].mean(), 2),
                "平均可用率(%)": round(mk["可用率(%)"].mean(), 2),
                "重複故障台數": int(mk["重複故障"].sum()),
            })
    return pd.DataFrame(rows)

def avail_badge(val):
    if val >= THRESHOLD_AVAIL:
        return f'<span class="badge badge-green">✓ {val}%</span>'
    elif val >= 80:
        return f'<span class="badge badge-amber">△ {val}%</span>'
    return f'<span class="badge badge-red">✗ {val}%</span>'

# ══════════════════════════════════════════════════════════════════════
def main():
    # ── 頂部標題 banner ──
    st.markdown("""
    <div class="page-banner">
        <div>
            <div class="page-banner-title">🔧 設備可靠性分析儀表板</div>
            <div class="page-banner-sub">富強醫材股份有限公司 設備課｜Equipment Reliability Dashboard</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    stored = load_data()

    # ── 資料更新區（主畫面可折疊）────────────────────────────────────────
    special_off = {}
    if stored and stored.get("special_off"):
        special_off = {k: v for k, v in stored["special_off"].items()}

    with st.expander("🔑 管理員更新資料", expanded=not bool(stored)):
        _tok = st.secrets.get("GITHUB_TOKEN", "")
        if not _tok:
            st.error("⚠️ GITHUB_TOKEN 未設定 — 請至 Streamlit Cloud → App settings → Secrets 加入")
        up_col, sp_col = st.columns([1, 1])
        with up_col:
            up_file = st.file_uploader("上傳報修 Excel（含時數工作表）",
                                       type=["xlsx"], key="uploader")
        with sp_col:
            with st.expander("📅 特殊停工調整（颱風、旅遊…）"):
                st.caption("填入該月額外未開工的時數，系統自動從可用時數扣除")
                for mo in MONTH_ORDER:
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
            elif df_all_new is None or df_all_new.empty:
                st.error("❌ 解析失敗：找不到有效資料，請確認 Excel 格式")
            else:
                series_found = df_all_new["機台代碼"].dropna().apply(
                    lambda c: re.match(r"(M-\d{3}-\d{3})", str(c)).group(1)
                    if re.match(r"(M-\d{3}-\d{3})", str(c)) else None
                ).dropna().unique()
                st.success(f"✓ 解析完成：{len(df_all_new)} 筆記錄，發現系列：{', '.join(sorted(series_found)[:10])}{'…' if len(series_found)>10 else ''}")
                df_kpi_new = calc_kpi(df_all_new, hours_base_new, special_off)
                raw_cols = ["月份","機台代碼","機台名稱","系列","設備類型","維修時數","停機","缺料","完工超時"]
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
                if ok:
                    st.toast("✅ 資料已儲存！頁面更新中…")
                    st.cache_data.clear()
                    st.rerun()
                else:
                    st.error(f"❌ 儲存失敗：{msg}")
        elif stored:
            st.caption(f"資料更新時間：{stored.get('updated','—')}")
            if st.button("更新特殊停工設定"):
                stored["special_off"] = special_off
                ok, msg = save_data(stored)
                if ok:
                    st.toast("✅ 已更新特殊停工設定")
                    st.cache_data.clear()
                else:
                    st.error(msg)

    # ── 無資料畫面 ────────────────────────────────────────────────────
    if not stored:
        st.markdown("""
        <div style="text-align:center;padding:80px 0;">
        <div style="font-size:48px;margin-bottom:12px">📊</div>
        <div style="font-size:18px;font-weight:600;color:#1a1a18;margin-bottom:8px">尚無資料</div>
        <div style="font-size:13px;color:#888780">請展開上方「🔑 管理員更新資料」區域上傳報修 Excel 檔案</div>
        </div>""", unsafe_allow_html=True)
        return

    # 重建 DataFrame
    df_kpi  = pd.DataFrame(stored["kpi"])
    df_raw  = pd.DataFrame(stored["raw"])
    hours_base  = stored.get("hours_base", {})
    special_off = stored.get("special_off", {})

    # 若舊資料無設備類型欄位，補推算
    if "設備類型" not in df_kpi.columns:
        def _type(code):
            if not isinstance(code, str): return "其他"
            m = re.match(r"M-(\d{3})", code)
            return EQ_TYPE_MAP.get(m.group(1), "其他") if m else "其他"
        df_kpi["設備類型"] = df_kpi["機台代碼"].apply(_type)
    if "設備類型" not in df_raw.columns:
        def _type2(code):
            if not isinstance(code, str): return "其他"
            m = re.match(r"M-(\d{3})", code)
            return EQ_TYPE_MAP.get(m.group(1), "其他") if m else "其他"
        df_raw["設備類型"] = df_raw["機台代碼"].apply(_type2)

    months_avail = sorted(df_kpi["月份"].unique(),
                          key=lambda x: MONTH_ORDER.index(x) if x in MONTH_ORDER else 99)
    df_summary = monthly_summary(df_kpi, hours_base, special_off)
    df_summary_type = monthly_summary_by_type(df_kpi, hours_base, special_off)

    # ── 資訊列 ──
    m_range = f"{months_avail[0].replace('2026-','')} – {months_avail[-1].replace('2026-','')}" if months_avail else ""
    info_col, btn_col = st.columns([6, 1])
    with info_col:
        st.caption(f"📅 2026年 {m_range}　｜　資料更新：{stored.get('updated','—')}")
    with btn_col:
        if st.button("↩ 重新上傳"):
            st.cache_data.clear()
            st.rerun()
    st.divider()

    # ── 全局篩選器 ────────────────────────────────────────────────────
    fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 1])
    with fc1:
        month_opts = ["全部月份"] + months_avail
        sel_month_single = st.selectbox("月份", month_opts, key="f_month_single")
        if sel_month_single == "全部月份":
            sel_months = months_avail
        else:
            sel_months = [sel_month_single]
    with fc2:
        all_series = sorted(df_kpi["系列"].unique())
        sel_series = st.multiselect("機台系列", all_series, default=[], key="f_series")
    with fc3:
        only_stop = st.checkbox("只顯示有停機記錄的機台", key="f_stop")
    with fc4:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🔄 重設篩選"):
            st.rerun()

    dk = df_kpi[df_kpi["月份"].isin(sel_months)].copy()
    dr = df_raw[df_raw["月份"].isin(sel_months)].copy()
    if sel_series:
        dk = dk[dk["系列"].isin(sel_series)]
        dr = dr[dr["系列"].isin(sel_series)]
    if only_stop:
        dk = dk[dk["停機次數"] > 0]

    ds = df_summary[df_summary["月份"].isin(sel_months)]

    # ══ 分頁 ═══════════════════════════════════════════════════════════
    tab1, tab2, tab3, tab4 = st.tabs(
        ["📊 月別總覽", "🔧 機台明細", "⚠️ 故障分析", "📈 趨勢分析"])

    # ══ Tab1：月別總覽（大類 → 系列代碼一覽表）════════════════════════════
    with tab1:
        TYPE_COLOR = {"成型類": "#378ADD", "加工類": "#1D9E75", "包裝類": "#BA7517"}

        def style_avail_cell(val):
            if not isinstance(val, (int, float)): return ""
            if val >= THRESHOLD_AVAIL: return "color:#1f7a4d;font-weight:600"
            elif val >= 80: return "color:#7b4f00"
            return "color:#a32d2d;font-weight:600"

        def style_failures(val):
            if not isinstance(val, (int, float)): return ""
            if val == 0: return "color:#888"
            if val >= 5: return "color:#a32d2d;font-weight:600"
            return ""

        for eq_type in EQ_TYPE_ORDER:
            type_color = TYPE_COLOR.get(eq_type, "#378ADD")
            # 該類所有資料
            dk_type = dk[dk["設備類型"] == eq_type]
            if dk_type.empty:
                continue

            # ── 大類標題 ──
            st.markdown(
                f'<div class="type-header" style="background:{type_color}">'
                f'▌ {eq_type}</div>',
                unsafe_allow_html=True)

            # ── 大類 KPI 摘要卡 ──
            k1, k2, k3, k4, k5 = st.columns(5)
            t_machines = dk_type["機台代碼"].nunique()
            t_fail  = int(dk_type["故障次數"].sum())
            t_down  = round(dk_type["停工時數(h)"].sum(), 2)
            t_mtbf  = round(dk_type["MTBF(h)"].mean(), 2)
            t_avail = round(dk_type["可用率(%)"].mean(), 2)
            k1.metric("涉及機台數", f"{t_machines} 台")
            k2.metric("故障次數", f"{t_fail} 次")
            k3.metric("總停工時數", f"{t_down} h")
            k4.metric("平均MTBF", f"{t_mtbf} h",
                      "✓ 達標" if t_mtbf >= THRESHOLD_MTBF else f"目標≥{THRESHOLD_MTBF}h",
                      delta_color="normal" if t_mtbf >= THRESHOLD_MTBF else "inverse")
            k5.metric("平均可用率", f"{t_avail}%",
                      "✓ 達標" if t_avail >= THRESHOLD_AVAIL else f"目標≥{THRESHOLD_AVAIL}%",
                      delta_color="normal" if t_avail >= THRESHOLD_AVAIL else "inverse")

            st.markdown("<br>", unsafe_allow_html=True)

            # ── 系列代碼彙總表（每列 = 一個系列）──
            series_list = sorted(dk_type["系列"].unique())
            tbl_rows = []
            for sc in series_list:
                sk = dk_type[dk_type["系列"] == sc]
                tbl_rows.append({
                    "系列代碼": sc,
                    "設備名稱": SERIES_NAME.get(sc, ""),
                    "台數": sk["機台代碼"].nunique(),
                    "故障次數": int(sk["故障次數"].sum()),
                    "停工時數(h)": round(sk["停工時數(h)"].sum(), 2),
                    "平均MTBF(h)": round(sk["MTBF(h)"].mean(), 2),
                    "平均MTTR(h)": round(sk["MTTR(h)"].mean(), 2),
                    "平均可用率(%)": round(sk["可用率(%)"].mean(), 2),
                    "重複故障台數": int(sk["重複故障"].sum()),
                })
            df_tbl = pd.DataFrame(tbl_rows)

            if not df_tbl.empty:
                styled = (df_tbl.style
                    .map(style_avail_cell, subset=["平均可用率(%)"])
                    .map(style_failures, subset=["故障次數"])
                    .format({
                        "停工時數(h)": "{:.2f}",
                        "平均MTBF(h)": "{:.2f}",
                        "平均MTTR(h)": "{:.2f}",
                        "平均可用率(%)": "{:.2f}",
                    }))
                st.dataframe(styled, use_container_width=True, hide_index=True)

                # ── 停工時數橫條圖（只顯示有故障的系列）──
                df_chart = df_tbl[df_tbl["故障次數"] > 0].copy()
                if not df_chart.empty:
                    df_chart["系列標籤"] = df_chart["系列代碼"] + " " + df_chart["設備名稱"]
                    df_chart = df_chart.sort_values("停工時數(h)", ascending=True)
                    fig_t = go.Figure()
                    fig_t.add_bar(
                        x=df_chart["故障次數"],
                        y=df_chart["系列標籤"],
                        orientation="h",
                        name="故障次數",
                        marker_color=type_color,
                        opacity=0.85,
                        text=df_chart["故障次數"],
                        textposition="outside",
                        xaxis="x",
                    )
                    fig_t.add_scatter(
                        x=df_chart["停工時數(h)"],
                        y=df_chart["系列標籤"],
                        mode="markers",
                        name="停工時數(h)",
                        marker=dict(color="#D85A30", size=10, symbol="diamond"),
                        xaxis="x2",
                    )
                    chart_h = max(200, len(df_chart) * 36 + 60)
                    fig_t.update_layout(**{
                        **PLOT_LAYOUT,
                        "height": chart_h,
                        "xaxis": dict(title="故障次數", side="bottom", **GRID),
                        "xaxis2": dict(title="停工時數(h)", overlaying="x", side="top",
                                       showgrid=False),
                        "yaxis": dict(automargin=True, **GRID),
                        "legend": dict(x=1.02, y=1, orientation="v"),
                        "margin": dict(t=48, b=40, l=180, r=100),
                    })
                    st.plotly_chart(fig_t, use_container_width=True)

            st.markdown("<br>", unsafe_allow_html=True)

    # ══ Tab2：機台明細 ═════════════════════════════════════════════════
    with tab2:
        st.markdown('<div class="section-title">機台明細 — MTBF / MTTR / 可用率</div>',
                    unsafe_allow_html=True)

        agg = dk.groupby(["機台代碼","機台名稱","系列","設備類型"]).agg(
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
            if not isinstance(val, (int, float)): return ""
            if val >= THRESHOLD_AVAIL: return "background-color:#e2efda"
            elif val >= 80: return "background-color:#fff2cc"
            return "background-color:#fcebeb"

        st.dataframe(
            agg.style
               .map(color_avail, subset=["平均可用率"])
               .format({
                   "停工時數":  "{:.2f}",
                   "總維修時數": "{:.2f}",
                   "平均MTBF":  "{:.2f}",
                   "平均MTTR":  "{:.2f}",
                   "平均可用率": "{:.2f}",
               }),
            use_container_width=True, hide_index=True)

        st.markdown("---")
        sel_m = st.selectbox("查看單台月別明細", ["— 選擇機台 —"] + list(agg["機台代碼"]))
        if sel_m != "— 選擇機台 —":
            dm = dk[dk["機台代碼"] == sel_m].sort_values(
                "月份", key=lambda x: x.map({m:i for i,m in enumerate(MONTH_ORDER)}))
            float_cols_dm = ["停工時數(h)", "MTBF(h)", "MTTR(h)", "可用率(%)"]
            st.dataframe(
                dm[["月份","故障次數","停機次數","停工時數(h)","MTBF(h)","MTTR(h)","可用率(%)"]
                  ].style.format({c: "{:.2f}" for c in float_cols_dm}),
                use_container_width=True, hide_index=True)

    # ══ Tab3：故障分析 ═════════════════════════════════════════════════
    with tab3:
        c1, c2 = st.columns(2)

        with c1:
            st.markdown('<div class="section-title">⬛ Top 10 停工時數（機台）</div>',
                        unsafe_allow_html=True)
            top10 = dk.groupby(["機台代碼","機台名稱"])["停工時數(h)"].sum()\
                      .reset_index().nlargest(10,"停工時數(h)")
            top10["停工時數(h)"] = top10["停工時數(h)"].round(2)
            fig2 = go.Figure(go.Bar(
                x=top10["停工時數(h)"], y=top10["機台代碼"],
                orientation="h", text=top10["停工時數(h)"],
                marker_color="#D85A30", opacity=0.85,
                hovertext=top10["機台名稱"]))
            fig2.update_layout(**PLOT_LAYOUT, height=340,
                yaxis=dict(autorange="reversed", gridcolor='#eae8e1'),
                title=dict(text="停工時數 TOP 10", font_size=13, x=0))
            st.plotly_chart(fig2, use_container_width=True)

        with c2:
            st.markdown('<div class="section-title">🔁 Top 10 故障次數（機台）</div>',
                        unsafe_allow_html=True)
            top10f = dk.groupby(["機台代碼","機台名稱"])["故障次數"].sum()\
                       .reset_index().nlargest(10,"故障次數")
            fig3 = go.Figure(go.Bar(
                x=top10f["故障次數"], y=top10f["機台代碼"],
                orientation="h", text=top10f["故障次數"],
                marker_color="#378ADD", opacity=0.85,
                hovertext=top10f["機台名稱"]))
            fig3.update_layout(**PLOT_LAYOUT, height=340,
                yaxis=dict(autorange="reversed", gridcolor='#eae8e1'),
                title=dict(text="故障次數 TOP 10", font_size=13, x=0))
            st.plotly_chart(fig3, use_container_width=True)

        st.markdown("---")
        c3, c4 = st.columns(2)

        with c3:
            st.markdown('<div class="section-title">⚠️ 重複故障機台（同月 ≥ 2 次）</div>',
                        unsafe_allow_html=True)
            repeat = dk[dk["重複故障"] == 1].groupby(["機台代碼","機台名稱"])\
                       .agg(發生月數=("月份","nunique"),
                            累計故障次數=("故障次數","sum")).reset_index()
            st.dataframe(repeat.sort_values("發生月數", ascending=False),
                         use_container_width=True, hide_index=True)

        with c4:
            st.markdown('<div class="section-title">📋 故障原因分類</div>',
                        unsafe_allow_html=True)
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
                    mask = dr["需求"].astype(str).apply(lambda x: any(k in x for k in kws))
                    idxs = dr[mask].index.tolist()
                    cat_counts[cat] = len([i for i in idxs if i not in matched])
                    matched.update(idxs)
                cat_counts["其他"] = len(dr) - sum(cat_counts.values())

                df_cat = pd.DataFrame(list(cat_counts.items()),
                                      columns=["故障類別","件數"])\
                           .sort_values("件數", ascending=False)
                fig4 = px.pie(df_cat, names="故障類別", values="件數",
                              color_discrete_sequence=CHART_COLORS)
                fig4.update_layout(height=320, margin=dict(t=10,b=10),
                                   paper_bgcolor='white')
                st.plotly_chart(fig4, use_container_width=True)
            else:
                st.info("無需求欄位資料")

        # 缺料
        st.markdown("---")
        st.markdown('<div class="section-title">📦 缺料統計（月別）</div>',
                    unsafe_allow_html=True)
        shortage_m = dk.groupby("月份")["缺料件數"].sum().reset_index()
        shortage_m = shortage_m.sort_values("月份", key=lambda x: x.map(
            {m:i for i,m in enumerate(MONTH_ORDER)}))
        fig5 = go.Figure(go.Bar(x=shortage_m["月份"], y=shortage_m["缺料件數"],
                                text=shortage_m["缺料件數"], marker_color="#BA7517", opacity=0.85))
        fig5.update_layout(**PLOT_LAYOUT, height=260,
            title=dict(text="缺料件數", font_size=13, x=0))
        st.plotly_chart(fig5, use_container_width=True)

    # ══ Tab4：趨勢分析 ════════════════════════════════════════════════
    with tab4:
        ds2 = df_summary[df_summary["月份"].isin(sel_months)].sort_values(
            "月份", key=lambda x: x.map({m:i for i,m in enumerate(MONTH_ORDER)}))

        c1, c2 = st.columns(2)

        with c1:
            st.markdown('<div class="section-title">MTBF / MTTR 月別趨勢</div>',
                        unsafe_allow_html=True)
            fig6 = go.Figure()
            fig6.add_scatter(x=ds2["月份"], y=ds2["平均MTBF(h)"],
                             name="平均MTBF(h)", mode="lines+markers",
                             line=dict(color="#378ADD", width=2), marker=dict(size=7))
            fig6.add_scatter(x=ds2["月份"], y=ds2["平均MTTR(h)"],
                             name="平均MTTR(h)", mode="lines+markers",
                             line=dict(color="#D85A30", width=2), marker=dict(size=7),
                             yaxis="y2")
            fig6.add_hline(y=THRESHOLD_MTBF, line_dash="dash", line_color="#378ADD",
                           annotation_text=f"MTBF目標 {THRESHOLD_MTBF}h")
            fig6.update_layout(**PLOT_LAYOUT,
                yaxis=dict(title="MTBF(h)", gridcolor='#eae8e1'),
                yaxis2=dict(title="MTTR(h)", overlaying="y", side="right"),
                legend=dict(x=0, y=1.08, orientation="h"), height=320)
            st.plotly_chart(fig6, use_container_width=True)

        with c2:
            st.markdown('<div class="section-title">平均可用率趨勢</div>',
                        unsafe_allow_html=True)
            fig7 = go.Figure()
            fig7.add_scatter(x=ds2["月份"], y=ds2["平均可用率(%)"],
                             mode="lines+markers",
                             line=dict(color="#1D9E75", width=2), marker=dict(size=7),
                             fill="tozeroy", fillcolor="rgba(29,158,117,0.08)")
            fig7.add_hline(y=THRESHOLD_AVAIL, line_dash="dash", line_color="#BA7517",
                           annotation_text=f"目標 {THRESHOLD_AVAIL}%")
            fig7.update_layout(**PLOT_LAYOUT, height=320,
                yaxis=dict(title="可用率(%)", range=[0,110], gridcolor='#eae8e1'))
            st.plotly_chart(fig7, use_container_width=True)

        # 熱圖 — 分三大類
        st.markdown('<div class="section-title">機台系列 × 月份 — 可用率熱圖</div>',
                    unsafe_allow_html=True)
        TYPE_COLOR_HM = {"成型類": "#378ADD", "加工類": "#1D9E75", "包裝類": "#BA7517"}
        for eq_type in EQ_TYPE_ORDER:
            dk_hm = dk[dk["設備類型"] == eq_type]
            if dk_hm.empty:
                continue
            type_color = TYPE_COLOR_HM.get(eq_type, "#378ADD")
            st.markdown(
                f'<div style="font-size:13px;font-weight:700;color:{type_color};'
                f'margin:16px 0 6px;padding-left:8px;border-left:4px solid {type_color}">'
                f'{eq_type}</div>', unsafe_allow_html=True)
            pivot = dk_hm.pivot_table(values="可用率(%)",
                                      index="系列", columns="月份", aggfunc="mean").round(2)
            ordered_cols = [c for c in MONTH_ORDER if c in pivot.columns]
            pivot = pivot[ordered_cols]
            # 短月份標籤
            pivot.columns = [c.replace("2026-", "") for c in pivot.columns]
            fig_hm = px.imshow(pivot, text_auto=True, aspect="auto",
                               color_continuous_scale=["#fcebeb","#fff2cc","#e2efda"],
                               zmin=70, zmax=100)
            fig_hm.update_layout(
                height=max(180, len(pivot) * 32 + 60),
                margin=dict(t=10, b=10, l=120, r=60),
                paper_bgcolor='white', plot_bgcolor='white',
                coloraxis_showscale=False,
            )
            fig_hm.update_xaxes(side="bottom", tickangle=0)
            st.plotly_chart(fig_hm, use_container_width=True)

    st.markdown("---")
    st.caption("富強醫材股份有限公司 設備課　© 2026")

if __name__ == "__main__":
    main()
