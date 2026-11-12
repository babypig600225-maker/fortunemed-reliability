"""
富強醫材 — 設備可靠性分析儀表板
Repo: fortunemed-reliability
"""
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import re, json, base64, requests, io, os, unicodedata
from equipment import parse_equipment_excel, map_repair_code, ACTIVE_STATUS
from monthly_report import build_month_tables, to_excel, STATUS_STYLE, STATUS_ORDER
import html as _html
from datetime import datetime, date

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

# ── 停工時數計算（與稼動率同口徑：上班日 08:00–20:00） ──────────────────
WIN_START, WIN_END = 8, 20   # 工作視窗：08:00–20:00

HOLIDAYS_2026 = {
    date(2026, 1, 1), date(2026, 2, 16), date(2026, 2, 17), date(2026, 2, 18),
    date(2026, 2, 19), date(2026, 2, 20), date(2026, 2, 27), date(2026, 4, 3),
    date(2026, 4, 6),  date(2026, 5, 1),  date(2026, 6, 19), date(2026, 9, 25),
    date(2026, 9, 28), date(2026, 10, 9), date(2026, 10, 26), date(2026, 12, 25),
}

def is_workday(d: date) -> bool:
    return d.weekday() < 5 and d not in HOLIDAYS_2026

def calc_window_hours(start, end) -> float:
    """計算時間段在上班日 08:00–20:00 內的小時數"""
    if pd.isna(start) or pd.isna(end) or end <= start:
        return 0.0
    total = 0.0
    day = pd.Timestamp(start).normalize()
    while day <= end:
        if is_workday(day.date()):
            win_s = day + pd.Timedelta(hours=WIN_START)
            win_e = day + pd.Timedelta(hours=WIN_END)
            s = max(start, win_s)
            e = min(end, win_e)
            if e > s:
                total += (e - s).total_seconds() / 3600
        day += pd.Timedelta(days=1)
    return total

# 設備流水碼 → 設備名稱對照表（來源：設備清單 Excel）
SERIES_NAME = {
    # 成型類 M-301
    "M-301-001": "混料機",
    "M-301-002": "送料機",
    "M-301-003": "抽料機",
    "M-301-004": "加硫機",
    "M-301-005": "押出機",
    "M-301-006": "雷射外徑測定器",
    "M-301-007": "自動外徑控制器",
    "M-301-008": "自動裁切機",
    "M-301-009": "單機型壓鑄成型機",
    "M-301-010": "自動射出成型機",
    "M-301-011": "液態矽膠射出成型機",
    "M-301-012": "無廢料射出成型機",
    "M-301-018": "烤箱",
    "M-301-019": "冰水機",
    "M-301-025": "後射式矽膠射出成型機",
    "M-301-026": "旋轉噴塗機",
    "M-301-027": "雷射外徑控制器",
    "M-301-028": "2KM兩液型供料系統",
    "M-301-030": "模溫機",
    "M-301-031": "塗佈機",
    "M-301-033": "粉碎機",
    "M-301-034": "放卷機",
    "M-301-035": "直壓式預成型機",
    # 加工類 M-302
    "M-302-002": "拉力測試機",
    "M-302-003": "淋膠機",
    "M-302-004": "充氣試球機",
    "M-302-005": "迷你型桌上型鑽床",
    "M-302-008": "磨R角機",
    "M-302-010": "TIP接著機",
    "M-302-011": "吸球吸力測試機",
    "M-302-012": "引流測試流量機",
    "M-302-013": "氣體測試治具",
    "M-302-014": "網印機",
    "M-302-016": "平面T3加熱器",
    "M-302-018": "超音波塑膠熔接機",
    "M-302-019": "黏球機",
    "M-302-021": "測漏機",
    "M-302-023": "側面驅動式封箱機",
    "M-302-026": "T管接著機",
    "M-302-027": "CV接著機",
    "M-302-028": "活塞色套組裝機",
    "M-302-029": "自動沖床機",
    "M-302-030": "氣動沖壓機",
    "M-302-031": "裁斷機",
    "M-302-039": "T3印刷加硫機",
    "M-302-041": "數位點膠機",
    "M-302-042": "研磨機",
    "M-302-044": "墨杯式移印機",
    "M-302-045": "吸球組裝機",
    "M-302-046": "攪拌機",
    "M-302-047": "C型油壓壓床",
    "M-302-048": "自動點膠機(XY Table)",
    "M-302-049": "封閉式刮痰管測試設備",
    "M-302-052": "呼吸器組裝機",
    "M-302-053": "引流袋焊接測漏設備",
    "M-302-055": "專業級電烤箱",
    # 包裝類 M-303
    "M-303-001": "封口機",
    "M-303-002": "手壓封口機",
    "M-303-003": "貼標機",
    "M-303-007": "條碼機",
    "M-303-008": "超音波清洗機",
    "M-303-009": "腳踏封口機",
    "M-303-010": "包裝機",
    "M-303-011": "打包機",
    "M-303-012": "培養鍋/箱",
    "M-303-013": "EO負壓滅菌鍋",
    "M-303-017": "燙金機",
    "M-303-018": "收縮機",
    "M-303-020": "預熱室",
    "M-303-021": "連續式封口機",
    "M-303-022": "充填機",
    "M-303-023": "鎖蓋機",
    "M-303-025": "泡殼機",
    "M-303-027": "油墨噴印機",
    "M-303-028": "密著包裝機",
    "M-303-029": "伸縮膜包裝機",
    "M-303-030": "打印機",
    # 其他
    "M-304-001": "手壓台",
    "M-306-027": "4k顯微影像觀察系統",
}

CHART_COLORS = ['#2563c9','#2e9d57','#f0a020','#d64541','#7F77DD','#D4537E',
                '#2CA6A4','#E5A234','#5C85D6','#B84DB8']

PLOT_LAYOUT = dict(
    plot_bgcolor='#ffffff', paper_bgcolor='#ffffff',
    font=dict(family='"Noto Sans TC","Microsoft JhengHei","PingFang TC",system-ui,sans-serif',
              size=12, color="#6b7684"),
    title_font=dict(size=14, color="#1d2530"),
    hoverlabel=dict(bgcolor="#ffffff", bordercolor="#e4e6ea",
                    font=dict(size=12, color="#1d2530")),
    margin=dict(t=48, b=40, l=48, r=16),
)
GRID = dict(gridcolor='#e4e6ea', linecolor='#e4e6ea')

st.set_page_config(
    page_title="富強醫材｜設備可靠性分析",
    page_icon="🔧", layout="wide",
    initial_sidebar_state="collapsed"
)

# ── 樣式 ─────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;600;700&display=swap');

/* CSS 設計變數（與稼動率儀表板一致） */
:root {
    --bg: #f7f7f5; --card: #ffffff; --fg: #1d2530; --muted: #6b7684;
    --line: #e4e6ea; --soft: #f1f3f6; --accent: #2563c9; --accent-soft: #e8f0fc;
    --c-low: #d64541; --c-mid: #f0a020; --c-ok: #2e9d57;
}

/* 隱藏 Streamlit 預設 header / footer / menu */
header[data-testid="stHeader"] { display: none !important; }
#MainMenu { visibility: hidden !important; }
footer { visibility: hidden !important; }

/* 整體字體與底色 */
html, body, .main, [data-testid="stAppViewContainer"] {
    font-family: "Noto Sans TC", "Microsoft JhengHei", "PingFang TC", system-ui, sans-serif !important;
    background: var(--bg) !important;
    color: var(--fg) !important;
}
.main { padding: 0 1.5rem; }
/* 與視覺化總覽相同寬度：內容 1180px，左右 18px */
.block-container {
    padding-top: 1rem !important; padding-left: 18px !important; padding-right: 18px !important;
    max-width: 1216px !important;
}
/* 基準字級 14px（與 HTML body 一致） */
[data-testid="stMarkdownContainer"] p, [data-testid="stMarkdownContainer"] li,
[data-testid="stText"], [data-testid="stWidgetLabel"] p {
    font-size: 14px !important; line-height: 1.55 !important;
}

/* 頂部標題 banner（簡潔白底風格） */
.page-banner {
    background: var(--card);
    border: 1px solid var(--line);
    border-radius: 12px;
    padding: 16px 24px;
    margin-bottom: 16px;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.page-banner-title {
    font-size: 22px;
    font-weight: 700;
    color: var(--fg);
    letter-spacing: 0.5px;
    white-space: normal;
    word-break: keep-all;
}
.page-banner-sub {
    font-size: 13px;
    color: var(--muted);
    margin-top: 4px;
}

/* KPI 卡片（白底 + 細邊框，與稼動率一致） */
div[data-testid="metric-container"], div[data-testid="stMetric"] {
    background: var(--card) !important;
    border-radius: 12px !important;
    padding: 14px 18px !important;
    border: 1px solid var(--line) !important;
}
div[data-testid="metric-container"] label, div[data-testid="stMetric"] label,
[data-testid="stMetricLabel"] p {
    font-size: 13px !important;
    color: var(--muted) !important;
    font-weight: 400 !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
}
[data-testid="stMetricValue"], [data-testid="stMetricValue"] *,
[data-testid="stMetricValue"] [data-testid="stMarkdownContainer"] p {
    font-size: 1.85rem !important;
    font-weight: 600 !important;
    color: var(--fg) !important;
    font-variant-numeric: tabular-nums;
    line-height: 1.25 !important;
}

/* Streamlit Tab 樣式（accent 底線） */
[data-testid="stTabs"] [role="tablist"] {
    border-bottom: 1px solid var(--line);
    gap: 4px;
}
[data-testid="stTabs"] [role="tab"] {
    font-family: "Noto Sans TC", "Microsoft JhengHei", sans-serif !important;
    font-size: 14px !important;
    color: var(--muted) !important;
    padding: 9px 14px !important;
    border-radius: 0 !important;
    border-bottom: 2px solid transparent !important;
    background: none !important;
}
[data-testid="stTabs"] [role="tab"][aria-selected="true"] {
    color: var(--accent) !important;
    border-bottom-color: var(--accent) !important;
    font-weight: 600 !important;
}

/* 設備類型標題 */
.type-header {
    font-size: 15.5px;
    font-weight: 700;
    padding: 12px 16px;
    border-radius: 12px;
    margin: 18px 0 10px 0;
    color: var(--fg);
    background: var(--card) !important;
    border: 1px solid var(--line);
    display: flex; align-items: center; gap: 10px;
}
.type-header i { width: 12px; height: 12px; border-radius: 50%; display: inline-block; flex: none; }

/* 設備流水碼小標題 */
.series-label {
    font-size: 15px;
    font-weight: 600;
    color: var(--fg);
    margin: 4px 0 8px 0;
}

/* 區塊標題 */
.section-title {
    font-size: 15.5px;
    font-weight: 600;
    color: var(--fg);
    margin: 14px 0 8px 0;
}

/* 徽章 */
.badge {
    display:inline-block; padding:3px 12px; border-radius:6px;
    font-size:13px; font-weight:600;
}
.badge-green { background:#e2efda; color:#1f7a4d; }
.badge-red   { background:#fcebeb; color:#a32d2d; }
.badge-amber { background:#fff2cc; color:#7b4f00; }
.badge-blue  { background:var(--accent-soft); color:var(--accent); }

/* 表格 */
.dataframe { font-size:14px !important; }

/* Selectbox / multiselect label */
[data-testid="stSelectbox"] label,
[data-testid="stMultiSelect"] label {
    font-size: 13px !important;
    color: var(--muted) !important;
}

/* expander 卡片化 */
[data-testid="stExpander"] {
    background: var(--card) !important;
    border: 1px solid var(--line) !important;
    border-radius: 12px !important;
}

/* 頂層主分頁（視覺化總覽 / 詳細分析）：加大字體、加底線 */
[data-testid="stTabs"]:first-of-type > div > [role="tablist"] > [role="tab"] {
    font-size: 15px !important;
    font-weight: 500 !important;
    padding: 10px 20px !important;
}
[data-testid="stTabs"]:first-of-type > div > [role="tablist"] > [role="tab"][aria-selected="true"] {
    font-size: 15px !important;
    font-weight: 700 !important;
}

/* dataframe：去除雜亂背景，表頭與 HTML 一致 */
[data-testid="stDataFrame"] > div {
    border-radius: 10px !important;
    border: 1px solid var(--line) !important;
    overflow: hidden;
}
[data-testid="stDataFrame"] th {
    background: var(--soft) !important;
    color: var(--muted) !important;
    font-size: 13px !important;
    font-weight: 600 !important;
    padding: 8px 10px !important;
}
[data-testid="stDataFrame"] td {
    font-size: 13px !important;
    padding: 7px 10px !important;
    color: var(--fg) !important;
}

/* selectbox / multiselect 控件：白底圓角邊框 */
[data-testid="stSelectbox"] > div > div,
[data-testid="stMultiSelect"] > div > div {
    background: var(--card) !important;
    border: 1px solid var(--line) !important;
    border-radius: 8px !important;
    font-size: 13px !important;
}

/* checkbox label */
[data-testid="stCheckbox"] span {
    font-size: 13px !important;
    color: var(--fg) !important;
}

/* button（重設篩選等） */
[data-testid="stButton"] > button {
    background: var(--card) !important;
    border: 1px solid var(--line) !important;
    border-radius: 8px !important;
    color: var(--fg) !important;
    font-size: 13px !important;
    font-family: "Noto Sans TC", "Microsoft JhengHei", sans-serif !important;
}
[data-testid="stButton"] > button:hover {
    border-color: var(--accent) !important;
    color: var(--accent) !important;
}

/* caption / 小字 */
[data-testid="stCaptionContainer"] p {
    font-size: 13px !important;
    color: var(--muted) !important;
}

/* divider */
hr {
    border-color: var(--line) !important;
    margin: 12px 0 !important;
}

/* 月度分析表 */
.ma-wrap { overflow-x: auto; margin: 0 0 8px; }
.ma-wrap { border: 1px solid var(--line); border-radius: 12px; background: var(--card); }
.ma-table { width: 100%; border-collapse: collapse; font-size: 13px; table-layout: auto; }
.ma-table th { background: var(--soft); color: var(--muted); font-weight: 600; padding: 8px 6px;
    border-bottom: 1px solid var(--line); text-align: center; line-height: 1.35; font-size: 12.5px; }
.ma-table td { padding: 8px 6px; border-bottom: 1px solid var(--line); text-align: center;
    vertical-align: top; font-variant-numeric: tabular-nums; }
.ma-table td.l { text-align: left; }
.ma-table td.txt { width: 40%; min-width: 280px; line-height: 1.55; color: var(--fg);
    word-break: break-all; }
.ma-table td.hot { color: #d64541; font-weight: 700; }
.ma-table tr:last-child td { border-bottom: 0; }
.st-pill { display: inline-block; padding: 2px 10px; border-radius: 20px; font-size: 12px;
    font-weight: 600; white-space: nowrap; }
.ma-watch { margin-top: 4px; font-size: 12px; color: var(--muted); }
.ma-legend { font-size: 13px; color: var(--fg); line-height: 2.1; margin: 4px 0 10px; }
.ma-rule { font-size: 12px; color: var(--muted); }
.ma-sub { font-size: 12.5px; font-weight: 400; color: var(--muted); margin-left: 4px; }

/* plotly chart 容器圓角 */
[data-testid="stPlotlyChart"] > div {
    border-radius: 12px !important;
    border: 1px solid var(--line) !important;
    overflow: hidden;
    background: var(--card) !important;
}
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
        m = re.search(r"M-\d{3}-\d{3}-\d+(?:-\d+)*",
                      unicodedata.normalize("NFKC", str(n)))
        return m.group(0) if m else None

    df_all["機台代碼"] = df_all["設備名稱"].apply(extract_code)
    df_all["維修時數"] = pd.to_numeric(df_all["維修時數(h)"], errors="coerce").fillna(0).clip(lower=0)
    df_all["停機"]    = df_all["是否停機"].str.upper() == "YES"
    # 時間戳（有則用於精確計算停工時數，與稼動率同口徑）
    for col in ("建立時間", "完工時間"):
        if col in df_all.columns:
            df_all[col] = pd.to_datetime(df_all[col], errors="coerce")
        else:
            df_all[col] = pd.NaT
    df_all["缺料"]    = df_all["缺料"].astype(str).str.upper().isin(["是","有","YES"])
    df_all["完工超時"] = df_all["完工-超時"].notna() & (df_all["完工-超時"].astype(str).str.strip() != "")
    # 單筆停工時數：只計「是否停機＝YES」；有時間戳用 08–20 上班時段計算，否則用維修時數
    def _rec_down(r):
        if not r["停機"]:
            return 0.0
        if pd.notna(r["建立時間"]) and pd.notna(r["完工時間"]):
            return round(calc_window_hours(r["建立時間"], r["完工時間"]), 2)
        return round(float(r["維修時數"]), 2)
    df_all["停工時數(h)"] = df_all.apply(_rec_down, axis=1)
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
            # 停工時數：有時間戳 → 計算上班日 08:00-20:00 內有效停工時數（與稼動率同口徑）
            #           無時間戳 → 退用填寫的維修時數
            stopped = mg[mg["停機"]]
            has_ts  = stopped["建立時間"].notna() & stopped["完工時間"].notna()
            downtime = (
                stopped[has_ts].apply(
                    lambda r: calc_window_hours(r["建立時間"], r["完工時間"]), axis=1
                ).sum()
                + stopped[~has_ts]["維修時數"].sum()
            )
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

# ── 在用機台名冊 × 月份（含無故障機台）與合計法 KPI ─────────────────────
def month_key_sort(m):
    return MONTH_ORDER.index(m) if m in MONTH_ORDER else 99

def load_roster(stored):
    """在用機台名冊：優先用管理員上傳的設備清單，否則用 repo 內 equipment.json"""
    if stored and stored.get("equipment"):
        return stored["equipment"]
    here = os.path.dirname(os.path.abspath(__file__))
    fp = os.path.join(here, "equipment.json")
    if os.path.exists(fp):
        with open(fp, encoding="utf-8") as f:
            return json.load(f)
    return []

def _month_end(mo):
    m = re.match(r"(\d{4})-(\d+)月", mo)
    if not m:
        return None
    return pd.Timestamp(int(m.group(1)), int(m.group(2)), 1) + pd.offsets.MonthEnd(0)

def build_full(df_kpi, roster, hours_base, special_off):
    """每月 × 在用機台一列；無報修的機台故障數 = 0。
    停用/報廢/設計中/請購中等不在名冊的機台不計入。"""
    num = ["故障次數", "停機次數", "停工時數(h)", "總維修時數(h)", "缺料件數", "超時完工件數"]
    months = sorted(df_kpi["月份"].unique(), key=month_key_sort)
    if not roster:                      # 無名冊：退回只用有報修的機台
        k = df_kpi.copy()
        k["使用狀態"] = "—"
        k["月可用時數(h)"] = k["月份"].map(
            lambda mo: max(hours_base.get(mo, 220) - special_off.get(mo, 0), 0))
    else:
        codes = {e["code"] for e in roster}
        kk = df_kpi.copy()
        kk["機台代碼"] = kk["機台代碼"].apply(lambda c: map_repair_code(c, codes))
        kk = kk[kk["機台代碼"].notna()]
        kg = kk.groupby(["月份", "機台代碼"])[num].sum().reset_index()
        rows = []
        for mo in months:
            me = _month_end(mo)
            ah = max(hours_base.get(mo, 220) - special_off.get(mo, 0), 0)
            for e in roster:
                ind = pd.to_datetime(e.get("in_date"), errors="coerce")
                if me is not None and pd.notna(ind) and ind > me:
                    continue            # 該月尚未入廠
                rows.append({"月份": mo, "機台代碼": e["code"], "機台名稱": e.get("name", ""),
                             "系列": e["series"],
                             "設備類型": EQ_TYPE_MAP.get(e["code"][2:5], "其他"),
                             "使用狀態": e.get("status", ""), "月可用時數(h)": ah})
            # 名冊有子代碼、報修用上層代碼的情形 → 補列
            extra = set(kg.loc[kg["月份"] == mo, "機台代碼"]) - {r["機台代碼"] for r in rows if r["月份"] == mo}
            for c in extra:
                rows.append({"月份": mo, "機台代碼": c,
                             "機台名稱": SERIES_NAME.get(c[:9], ""), "系列": c[:9],
                             "設備類型": EQ_TYPE_MAP.get(c[2:5], "其他"),
                             "使用狀態": "使用中", "月可用時數(h)": ah})
        k = pd.DataFrame(rows).merge(kg, on=["月份", "機台代碼"], how="left")
    for c in num:
        k[c] = k[c].fillna(0)
    for c in ["故障次數", "停機次數", "缺料件數", "超時完工件數"]:
        k[c] = k[c].astype(int)
    k["運作時數(h)"] = (k["月可用時數(h)"] - k["停工時數(h)"]).clip(lower=0).round(2)
    f = k["故障次數"]
    k["MTBF(h)"] = (k["運作時數(h)"] / f.where(f > 0)).round(2)
    k["MTTR(h)"] = (k["總維修時數(h)"] / f.where(f > 0)).round(2)
    k["可用率(%)"] = (k["運作時數(h)"] / k["月可用時數(h)"].where(k["月可用時數(h)"] > 0) * 100).round(2)
    k["重複故障"] = (f >= 2).astype(int)
    return k.reset_index(drop=True)

def pool(df):
    """合計法：MTBF = Σ(月時數−停工) ÷ Σ故障；MTTR = Σ維修時數 ÷ Σ故障；可用率 = Σ運作 ÷ Σ月時數"""
    f   = int(df["故障次數"].sum()) if len(df) else 0
    rep = float(df["總維修時數(h)"].sum()) if len(df) else 0.0
    ah  = float(df["月可用時數(h)"].sum()) if len(df) else 0.0
    op  = float(df["運作時數(h)"].sum()) if len(df) else 0.0
    return {
        "faults": f,
        "stops": int(df["停機次數"].sum()) if len(df) else 0,
        "repair": round(rep, 2),
        "downtime": round(float(df["停工時數(h)"].sum()), 2) if len(df) else 0.0,
        "hours": round(ah, 2), "op": round(op, 2),
        "n": int(df["機台代碼"].nunique()) if len(df) else 0,
        "n_fault": int(df.loc[df["故障次數"] > 0, "機台代碼"].nunique()) if len(df) else 0,
        "mtbf": round(op / f, 2) if f else None,
        "mttr": round(rep / f, 2) if f else None,
        "avail": round(op / ah * 100, 2) if ah else None,
    }

def judge(faults, mtbf, mttr):
    """狀態判定"""
    if not faults:
        return "表現卓越"
    if mtbf is not None and mtbf < 50:
        return "需重點改善"
    if (mtbf is not None and mtbf < THRESHOLD_MTBF) or (mttr is not None and mttr > THRESHOLD_MTTR):
        return "需改善"
    return "正常"

def _summary_row(mk):
    p = pool(mk)
    return {
        "在用機台數": p["n"], "涉及機台數": p["n_fault"],
        "故障次數": p["faults"], "停工時數(h)": p["downtime"],
        "平均MTBF(h)": p["mtbf"], "平均MTTR(h)": p["mttr"],
        "平均可用率(%)": p["avail"],
        "重複故障台數": int(mk["重複故障"].sum()),
    }

def monthly_summary(df_kpi, hours_base, special_off):
    rows = []
    for month in sorted(df_kpi["月份"].unique(), key=month_key_sort):
        mk = df_kpi[df_kpi["月份"] == month]
        rows.append({"月份": month, **_summary_row(mk),
                     "月可用時數(h)": max(hours_base.get(month, 220) - special_off.get(month, 0), 0)})
    return pd.DataFrame(rows)

def monthly_summary_by_type(df_kpi, hours_base, special_off):
    """各設備類型 × 月份彙總（合計法）"""
    rows = []
    for eq_type in EQ_TYPE_ORDER:
        sub = df_kpi[df_kpi["設備類型"] == eq_type]
        for month in sorted(sub["月份"].unique(), key=month_key_sort):
            rows.append({"設備類型": eq_type, "月份": month,
                         **_summary_row(sub[sub["月份"] == month])})
    return pd.DataFrame(rows)

def build_dashboard_json(df_kpi, hours_base, special_off, updated=""):
    """Build JSON payload for dashboard_reliability.html template（合計法）。"""
    months_sorted = sorted(df_kpi["月份"].unique(), key=month_key_sort)
    short_months = [m.replace("2026-","") for m in months_sorted]
    types_all = ["全部"] + EQ_TYPE_ORDER
    r1 = lambda v: None if v is None else round(v, 1)

    def kpi_for(subset, month_key):
        df = subset if month_key == "全年" else subset[subset["月份"] == "2026-" + month_key]
        if df.empty:
            return {}
        p = pool(df)
        per_m = df.groupby("機台代碼")[["運作時數(h)", "月可用時數(h)"]].sum()
        low = int(((per_m["運作時數(h)"] / per_m["月可用時數(h)"] * 100) < THRESHOLD_AVAIL).sum())
        return {"avail": r1(p["avail"]), "mtbf": r1(p["mtbf"]), "mttr": r1(p["mttr"]),
                "downtime": r1(p["downtime"]), "faults": p["faults"],
                "machines": p["n_fault"], "active": p["n"], "low_avail": low}

    kpis, trend = {}, {}
    for t in types_all:
        sub = df_kpi if t == "全部" else df_kpi[df_kpi["設備類型"] == t]
        kpis[t] = {"全年": kpi_for(sub, "全年")}
        rows = []
        for sm in short_months:
            kd = kpi_for(sub, sm)
            if kd:
                kpis[t][sm] = kd
                rows.append({"m": sm, "avail": kd["avail"], "mtbf": kd["mtbf"], "mttr": kd["mttr"],
                             "downtime": kd["downtime"], "faults": kd["faults"]})
        trend[t] = rows

    # 系列 KPI（合計法，含無故障機台）
    series_kpi = {}
    for sc, grp in df_kpi.groupby("系列"):
        d = {}
        for key, g in [("全年", grp)] + [(sm, grp[grp["月份"] == "2026-" + sm]) for sm in short_months]:
            if g.empty:
                continue
            p = pool(g)
            d[key] = {"mtbf": r1(p["mtbf"]), "mttr": r1(p["mttr"]), "avail": r1(p["avail"]),
                      "faults": p["faults"], "n": p["n"], "n_fault": p["n_fault"]}
        series_kpi[sc] = d

    # 機台：只列有故障的機台；全年值 = 合計法（含無故障月份）
    machines = []
    fault_codes = df_kpi.loc[df_kpi["故障次數"] > 0, "機台代碼"].unique()
    for code, grp in df_kpi[df_kpi["機台代碼"].isin(fault_codes)].groupby("機台代碼"):
        p = pool(grp)
        monthly = []
        for sm in short_months:
            dm = grp[(grp["月份"] == "2026-" + sm) & (grp["故障次數"] > 0)]
            if dm.empty:
                continue
            row = dm.iloc[0]
            monthly.append({"m": sm, "avail": r1(row["可用率(%)"]), "mtbf": r1(row["MTBF(h)"]),
                            "mttr": r1(row["MTTR(h)"]), "faults": int(row["故障次數"]),
                            "downtime": r1(row["停工時數(h)"])})
        machines.append({
            "code": code, "name": grp["機台名稱"].iloc[0],
            "type": grp["設備類型"].iloc[0], "series": grp["系列"].iloc[0],
            "avail": r1(p["avail"]), "mtbf": r1(p["mtbf"]), "mttr": r1(p["mttr"]),
            "faults": p["faults"], "downtime": r1(p["downtime"]), "monthly": monthly,
        })

    # Series display names (from Excel mapping)
    series_names = {
        "M-301-001": "混料機", "M-301-002": "送料機", "M-301-003": "抽料機",
        "M-301-004": "加硫機", "M-301-005": "押出機", "M-301-006": "雷射外徑測定器",
        "M-301-007": "自動外徑控制器", "M-301-008": "自動裁切機",
        "M-301-009": "單機型壓鑄成型機", "M-301-010": "自動射出成型機",
        "M-301-011": "液態矽膠射出成型機", "M-301-012": "無廢料射出成型機",
        "M-301-018": "烤箱", "M-301-019": "冰水機",
        "M-301-025": "後射式矽膠射出成型機", "M-301-026": "旋轉噴塗機",
        "M-301-027": "雷射外徑控制器", "M-301-028": "2KM兩液型供料系統",
        "M-301-030": "模溫機", "M-301-031": "塗佈機", "M-301-033": "粉碎機",
        "M-301-034": "放卷機", "M-301-035": "直壓式預成型機",
        "M-302-002": "拉力測試機", "M-302-003": "淋膠機", "M-302-004": "充氣試球機",
        "M-302-005": "迷你型桌上型鑽床", "M-302-008": "磨R角機",
        "M-302-010": "TIP接著機", "M-302-011": "吸球吸力測試機",
        "M-302-012": "引流測試流量機", "M-302-013": "氣體測試治具",
        "M-302-014": "網印機", "M-302-016": "平面T3加熱器",
        "M-302-018": "超音波塑膠熔接機", "M-302-019": "黏球機",
        "M-302-021": "測漏機", "M-302-023": "側面驅動式封箱機",
        "M-302-026": "T管接著機", "M-302-027": "CV接著機",
        "M-302-028": "活塞色套組裝機", "M-302-029": "自動沖床機",
        "M-302-030": "氣動沖壓機", "M-302-031": "裁斷機",
        "M-302-039": "T3印刷加硫機", "M-302-041": "數位點膠機",
        "M-302-042": "研磨機", "M-302-044": "墨杯式移印機",
        "M-302-045": "吸球組裝機", "M-302-046": "攪拌機",
        "M-302-047": "C型油壓壓床", "M-302-048": "自動點膠機(XY Table)",
        "M-302-049": "封閉式刮痰管測試設備", "M-302-052": "呼吸器組裝機",
        "M-302-053": "引流袋焊接測漏設備", "M-302-055": "專業級電烤箱",
        "M-303-001": "封口機", "M-303-002": "手壓封口機", "M-303-003": "貼標機",
        "M-303-007": "條碼機", "M-303-008": "超音波清洗機",
        "M-303-009": "腳踏封口機", "M-303-010": "包裝機", "M-303-011": "打包機",
        "M-303-012": "培養鍋/箱", "M-303-013": "EO負壓滅菌鍋",
        "M-303-017": "燙金機", "M-303-018": "收縮機", "M-303-020": "預熱室",
        "M-303-021": "連續式封口機", "M-303-022": "充填機", "M-303-023": "鎖蓋機",
        "M-303-025": "泡殼機", "M-303-027": "油墨噴印機",
        "M-303-028": "密著包裝機", "M-303-029": "伸縮膜包裝機",
        "M-303-030": "打印機", "M-304-001": "手壓台",
        "M-306-027": "4k顯微影像觀察系統",
    }

    return {
        "updated": updated,
        "months": short_months,
        "types": types_all,
        "kpis": kpis,
        "trend": trend,
        "machines": machines,
        "series_kpi": series_kpi,
        "series_names": series_names,
    }


def render_html_dashboard(data_json: dict):
    """Load dashboard_reliability.html, inject JSON data, render via components.html()."""
    here = os.path.dirname(os.path.abspath(__file__))
    tpl_path = os.path.join(here, "dashboard_reliability.html")
    if not os.path.exists(tpl_path):
        st.error("找不到 dashboard_reliability.html 模板，請確認檔案存在。")
        return
    with open(tpl_path, encoding="utf-8") as f:
        html = f.read()
    json_str = json.dumps(data_json, ensure_ascii=False)
    html = html.replace("/*__DATA__*/null", json_str, 1)
    components.html(html, height=2800, scrolling=True)


def avail_badge(val):
    if val >= THRESHOLD_AVAIL:
        return f'<span class="badge badge-green">✓ {val}%</span>'
    elif val >= 80:
        return f'<span class="badge badge-amber">△ {val}%</span>'
    return f'<span class="badge badge-red">✗ {val}%</span>'

def _status_pill(v):
    bg, fg = STATUS_STYLE.get(v, ("#f1f3f6", "#6b7684"))
    return (f'<span class="st-pill" style="background:{bg};color:{fg}">{_html.escape(str(v))}</span>')

def _num(v, unit=""):
    if isinstance(v, str):
        return _html.escape(v)
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "—"
    return f"{v:,.2f}".rstrip("0").rstrip(".") + unit

def render_monthly_analysis(df_kpi, df_raw, months_avail, excluded):
    st.markdown('<div class="section-title">月度可靠性分析</div>', unsafe_allow_html=True)
    st.caption("系列＝合計法：MTBF＝(在用台數×月標準時數 − 停工合計)÷故障合計，MTTR＝維修工時合計÷故障合計；"
               "每張報修單算 1 次故障，停工只計「是否停機＝YES」。"
               "只計使用中／測試中／驗證中的 M-301／302／303 機台。")
    c1, c2, c3 = st.columns([2, 2, 3])
    month = c1.selectbox("分析月份", list(reversed(months_avail)), key="ma_month")
    typ = c2.selectbox("設備類型", ["全部"] + EQ_TYPE_ORDER, key="ma_type")
    show = c3.radio("顯示系列", ["全部", "只看有故障", "只看需改善"], horizontal=True, key="ma_show")

    dk = df_kpi if typ == "全部" else df_kpi[df_kpi["設備類型"] == typ]
    dr = df_raw if typ == "全部" else df_raw[df_raw["設備類型"] == typ]
    df_s, df_m = build_month_tables(month, dk, dr, pool, judge, SERIES_NAME)
    if df_s.empty:
        st.info("此月份無資料")
        return

    # ── KPI 列 ──
    p = pool(dk[dk["月份"] == month])
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("月標準時數", f"{df_s['月標準時數'].iloc[0]:g} h")
    k2.metric("在用／故障機台", f"{p['n']}／{p['n_fault']} 台")
    k3.metric("故障次數", f"{p['faults']} 次")
    k4.metric("MTBF（合計）", f"{p['mtbf']} h" if p["mtbf"] is not None else "無故障",
              "✓ 達標" if (p["mtbf"] or 1e9) >= THRESHOLD_MTBF else f"目標≥{THRESHOLD_MTBF:g}h",
              delta_color="normal" if (p["mtbf"] or 1e9) >= THRESHOLD_MTBF else "inverse")
    k5.metric("MTTR（合計）", f"{p['mttr']} h" if p["mttr"] is not None else "—",
              "✓ 達標" if (p["mttr"] or 0) <= THRESHOLD_MTTR else f"目標≤{THRESHOLD_MTTR:g}h",
              delta_color="normal" if (p["mttr"] or 0) <= THRESHOLD_MTTR else "inverse")

    cnt_s = df_s["狀態判定"].value_counts()
    cnt_m = df_m["狀態判定"].value_counts()
    order = ["需重點改善", "需改善", "正常", "表現卓越"]
    st.markdown(
        '<div class="ma-legend">系列：' + "".join(f"{_status_pill(o)} {int(cnt_s.get(o,0))}　" for o in order)
        + '<br>機台：' + "".join(f"{_status_pill(o)} {int(cnt_m.get(o,0))}　" for o in order)
        + '<br><span class="ma-rule">判定：無故障＝表現卓越；MTBF＜50h＝需重點改善；'
          f'MTBF＜{THRESHOLD_MTBF:g}h 或 MTTR＞{THRESHOLD_MTTR:g}h＝需改善；其餘＝正常</span></div>',
        unsafe_allow_html=True)

    st.download_button(
        "⬇ 下載本月分析 Excel",
        data=to_excel(df_s, df_m, month),
        file_name=f"設備可靠性月度分析_{month}{'' if typ=='全部' else '_'+typ}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="ma_dl")

    # ── 系列表（合計）──
    view = df_s.copy()
    if show == "只看有故障":
        view = view[view["故障總次數"] > 0]
    elif show == "只看需改善":
        view = view[view["狀態判定"].isin(["需重點改善", "需改善"])]
    view = view.assign(_o=view["狀態判定"].map(STATUS_ORDER)).sort_values(["設備類型", "設備系列"])
    for t in [x for x in EQ_TYPE_ORDER if x in set(view["設備類型"])]:
        vt = view[view["設備類型"] == t]
        st.markdown(f'<div class="type-header"><i style="background:'
                    f'{ {"成型類":"#2563c9","加工類":"#2e9d57","包裝類":"#f0a020"}.get(t,"#2563c9") }'
                    f'"></i>{t}　<span class="ma-sub">{len(vt)} 個系列</span></div>',
                    unsafe_allow_html=True)
        head = ["設備名稱", "設備系列", "在用／<br>故障台數", "故障<br>總次數", "總維修<br>工時(H)",
                "總維修<br>停工時間(H)", "MTBF<br>(小時/次)", "MTTR<br>(小時/次)", "狀態判定", "分析與建議"]
        rows = []
        for _, r in vt.iterrows():
            red = ' class="hot"' if r["故障總次數"] >= 5 else ""
            watch = (f'<div class="ma-watch">需關注：{_html.escape(r["需關注機台"])}</div>'
                     if r["需關注機台"] else "")
            rows.append(
                "<tr>"
                f"<td class='l'>{_html.escape(str(r['設備名稱']))}</td>"
                f"<td>{r['設備系列']}</td>"
                f"<td>{r['在用台數']}／{r['故障台數']}</td>"
                f"<td{red}>{r['故障總次數']}</td>"
                f"<td>{_num(r['總維修工時(H)'])}</td>"
                f"<td>{_num(r['總維修停工時間(H)'])}</td>"
                f"<td>{_num(r['MTBF(小時/次)'])}</td>"
                f"<td>{_num(r['MTTR(小時/次)'])}</td>"
                f"<td>{_status_pill(r['狀態判定'])}</td>"
                f"<td class='l txt'>{_html.escape(r['分析與建議'])}{watch}</td>"
                "</tr>")
        st.markdown('<div class="ma-wrap"><table class="ma-table"><thead><tr>'
                    + "".join(f"<th>{h}</th>" for h in head) + "</tr></thead><tbody>"
                    + "".join(rows) + "</tbody></table></div>", unsafe_allow_html=True)

    # ── 機台表（個別）──
    st.markdown('<div class="section-title">機台明細（個別）</div>', unsafe_allow_html=True)
    only_f = st.checkbox("只顯示本月有故障的機台", value=True, key="ma_only_f")
    dm = df_m[df_m["故障次數"] > 0] if only_f else df_m
    dm = dm.assign(_o=dm["狀態判定"].map(STATUS_ORDER)).sort_values(["_o", "設備代號"]).drop(columns="_o")
    def _st_color(v):
        bg, fg = STATUS_STYLE.get(v, ("", ""))
        return f"background-color:{bg};color:{fg};font-weight:600" if bg else ""
    st.dataframe(
        dm[["設備代號", "設備名稱", "設備系列", "故障次數", "總維修工時(H)", "總維修停工時間(H)",
            "MTBF(小時/次)", "MTTR(小時/次)", "狀態判定", "主要故障"]]
          .style.map(_st_color, subset=["狀態判定"])
          .format({"總維修工時(H)": "{:.2f}", "總維修停工時間(H)": "{:.2f}", "MTTR(小時/次)": "{:.2f}",
                   "MTBF(小時/次)": lambda v: v if isinstance(v, str) else f"{v:.2f}"}),
        use_container_width=True, hide_index=True,
        column_config={"主要故障": st.column_config.TextColumn(width="large")})

    # ── 未列入計算的報修單 ──
    ex = excluded[excluded["月份"] == month] if len(excluded) else excluded
    if len(ex):
        with st.expander(f"未列入計算的報修單（{len(ex)} 筆：機台不在在用名冊，如停用、報廢或非 M-301/302/303）"):
            cols = [c for c in ["原始代碼", "機台名稱", "需求", "維修時數", "停機"] if c in ex.columns]
            st.dataframe(ex[cols], use_container_width=True, hide_index=True)

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
            eq_file = st.file_uploader(
                "上傳設備清單 Excel（選用：更新在用機台名冊）", type=["xlsx"], key="eq_uploader",
                help="只取 M-301 / M-302 / M-303 工作表；使用狀態為 使用中、測試中、驗證中 才列入")
            _roster_now = load_roster(stored)
            st.caption(f"目前在用機台名冊：{len(_roster_now)} 台"
                       + ("（已上傳的設備清單）" if stored and stored.get("equipment") else "（預設設備清單）"))
            if eq_file:
                try:
                    eq_new = parse_equipment_excel(eq_file.getvalue())
                except Exception as ex:
                    eq_new = []
                    st.error(f"設備清單解析失敗：{ex}")
                if eq_new:
                    st.info(f"解析到在用機台 {len(eq_new)} 台（{'、'.join(ACTIVE_STATUS)}）")
                    if stored and st.button("儲存設備清單"):
                        stored["equipment"] = eq_new
                        ok, msg = save_data(stored)
                        if ok:
                            st.toast("✅ 設備清單已更新")
                            st.cache_data.clear()
                            st.rerun()
                        else:
                            st.error(msg)
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
                for col in ["需求","部門","設備名稱","seq","停工時數(h)"]:
                    if col in df_all_new.columns:
                        raw_cols.append(col)
                payload = {
                    "updated": datetime.now().strftime("%Y-%m-%d %H:%M"),
                    "hours_base": hours_base_new,
                    "special_off": special_off,
                    "kpi": df_kpi_new.to_dict(orient="records"),
                    "raw": df_all_new[raw_cols].to_dict(orient="records"),
                }
                if stored and stored.get("equipment"):
                    payload["equipment"] = stored["equipment"]
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
        <div style="font-size:18px;font-weight:600;color:#1d2530;margin-bottom:8px">尚無資料</div>
        <div style="font-size:13px;color:#6b7684">請展開上方「🔑 管理員更新資料」區域上傳報修 Excel 檔案</div>
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

    # ── 在用機台名冊：補入無故障機台、排除停用/報廢等 ──
    roster = load_roster(stored)
    if roster:
        _codes = {e["code"] for e in roster}
        _name = {e["code"]: e.get("name", "") for e in roster}
        df_raw["原始代碼"] = df_raw["機台代碼"]
        df_raw["機台代碼"] = df_raw["機台代碼"].apply(lambda c: map_repair_code(c, _codes))
        excluded = df_raw[df_raw["機台代碼"].isna() & df_raw["原始代碼"].notna()]
        df_raw = df_raw[df_raw["機台代碼"].notna()].copy()
    else:
        excluded = df_raw.iloc[0:0]
    if "停工時數(h)" not in df_raw.columns:      # 舊資料：無單筆停工 → 停機單用維修時數
        df_raw["停工時數(h)"] = df_raw["維修時數"].where(df_raw["停機"].astype(bool), 0.0)
    df_kpi = build_full(df_kpi, roster, hours_base, special_off)

    months_avail = sorted(df_kpi["月份"].unique(),
                          key=lambda x: MONTH_ORDER.index(x) if x in MONTH_ORDER else 99)
    df_summary = monthly_summary(df_kpi, hours_base, special_off)
    df_summary_type = monthly_summary_by_type(df_kpi, hours_base, special_off)

    # ── HTML 自訂儀表板（與稼動率相同風格）──────────────────────────────
    dash_data = build_dashboard_json(df_kpi, hours_base, special_off,
                                     updated=stored.get("updated", ""))

    # ══ 頂層分頁：視覺化總覽 ／ 詳細分析 ══════════════════════════════════
    main_tab1, main_tab2 = st.tabs(["📊 視覺化總覽", "📋 詳細分析"])

    with main_tab1:
        render_html_dashboard(dash_data)

    with main_tab2:
        # ── 資訊列 ──
        m_range = f"{months_avail[0].replace('2026-','')} – {months_avail[-1].replace('2026-','')}" if months_avail else ""
        st.caption(f"📅 2026年 {m_range}　｜　資料更新：{stored.get('updated','—')}")
        st.divider()

        # ── 全局篩選器 ────────────────────────────────────────────────────
        fc1, fc2, fc3, fc4 = st.columns([2, 2, 2, 1])
        with fc1:
            # 月份：最新月份排最前（方便選最近一個月）
            months_desc = list(reversed(months_avail))
            month_opts = ["全部月份"] + months_desc
            sel_month_single = st.selectbox("月份", month_opts, key="f_month_single")
            if sel_month_single == "全部月份":
                sel_months = months_avail
            else:
                sel_months = [sel_month_single]
        with fc2:
            all_series = sorted(df_kpi["系列"].unique())
            # 顯示標籤加上中文名稱，如「M-301-005 押出機」
            series_labels = [f"{s}　{SERIES_NAME[s]}" if s in SERIES_NAME else s for s in all_series]
            label_to_series = dict(zip(series_labels, all_series))
            sel_labels = st.multiselect("機台系列", series_labels, default=[], key="f_series")
            sel_series = [label_to_series[lb] for lb in sel_labels]
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
        tab5, tab1, tab2, tab3, tab4 = st.tabs(
            ["📝 月度分析", "📊 月別總覽", "🔧 機台明細", "⚠️ 故障分析", "📈 趨勢分析"])

        # ══ Tab1：月別總覽（大類 → 系列代碼一覽表）════════════════════════════
        with tab1:
            TYPE_COLOR = {"成型類": "#2563c9", "加工類": "#2e9d57", "包裝類": "#f0a020"}

            def style_avail_cell(val):
                if not isinstance(val, (int, float)): return ""
                if val >= THRESHOLD_AVAIL: return "color:#1f7a4d;font-weight:600"
                elif val >= 80: return "color:#7b4f00"
                return "color:#a32d2d;font-weight:600"

            def style_failures(val):
                if not isinstance(val, (int, float)): return ""
                if val == 0: return "color:#6b7684"
                if val >= 5: return "color:#a32d2d;font-weight:600"
                return ""

            for eq_type in EQ_TYPE_ORDER:
                type_color = TYPE_COLOR.get(eq_type, "#2563c9")
                # 該類所有資料
                dk_type = dk[dk["設備類型"] == eq_type]
                if dk_type.empty:
                    continue

                # ── 大類標題 ──
                st.markdown(
                    f'<div class="type-header"><i style="background:{type_color}"></i>'
                    f'{eq_type}</div>',
                    unsafe_allow_html=True)

                # ── 大類 KPI 摘要卡 ──
                k1, k2, k3, k4, k5 = st.columns(5)
                tp = pool(dk_type)
                t_fail  = tp["faults"]
                t_down  = tp["downtime"]
                t_mtbf  = tp["mtbf"] if tp["mtbf"] is not None else 0
                t_avail = tp["avail"] if tp["avail"] is not None else 0
                k1.metric("在用／故障機台", f"{tp['n']}／{tp['n_fault']} 台")
                k2.metric("故障次數", f"{t_fail} 次")
                k3.metric("總停工時數", f"{t_down} h")
                k4.metric("MTBF（合計）", f"{tp['mtbf']} h" if tp["mtbf"] is not None else "無故障",
                          "✓ 達標" if t_mtbf >= THRESHOLD_MTBF or not t_fail else f"目標≥{THRESHOLD_MTBF}h",
                          delta_color="normal" if t_mtbf >= THRESHOLD_MTBF or not t_fail else "inverse")
                k5.metric("可用率（合計）", f"{t_avail}%",
                          "✓ 達標" if t_avail >= THRESHOLD_AVAIL else f"目標≥{THRESHOLD_AVAIL}%",
                          delta_color="normal" if t_avail >= THRESHOLD_AVAIL else "inverse")

                st.markdown("<br>", unsafe_allow_html=True)

                # ── 系列代碼彙總表（每列 = 一個系列）──
                series_list = sorted(dk_type["系列"].unique())
                tbl_rows = []
                for sc in series_list:
                    sk = dk_type[dk_type["系列"] == sc]
                    sp = pool(sk)
                    tbl_rows.append({
                        "系列代碼": sc,
                        "設備名稱": SERIES_NAME.get(sc, ""),
                        "在用台數": sp["n"],
                        "故障台數": sp["n_fault"],
                        "故障次數": sp["faults"],
                        "停工時數(h)": sp["downtime"],
                        "平均MTBF(h)": sp["mtbf"],
                        "平均MTTR(h)": sp["mttr"],
                        "平均可用率(%)": sp["avail"],
                        "重複故障台數": int(sk["重複故障"].sum()),
                    })
                df_tbl = pd.DataFrame(tbl_rows)
                for _c in ["平均MTBF(h)", "平均MTTR(h)", "平均可用率(%)"]:
                    if _c in df_tbl.columns:
                        df_tbl[_c] = pd.to_numeric(df_tbl[_c], errors="coerce")

                if not df_tbl.empty:
                    styled = (df_tbl.style
                        .map(style_avail_cell, subset=["平均可用率(%)"])
                        .map(style_failures, subset=["故障次數"])
                        .format({
                            "停工時數(h)": "{:.2f}",
                            "平均MTBF(h)": "{:.2f}",
                            "平均MTTR(h)": "{:.2f}",
                            "平均可用率(%)": "{:.2f}",
                        }, na_rep="無故障"))
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
                            marker=dict(color="#d64541", size=10, symbol="diamond"),
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

            show_zero = st.checkbox("顯示無故障的在用機台", value=False, key="t2_zero")
            _rows = []
            for (code, name, ser, typ), g in dk.groupby(["機台代碼","機台名稱","系列","設備類型"]):
                pp = pool(g)
                if not show_zero and pp["faults"] == 0:
                    continue
                _rows.append({"機台代碼": code, "機台名稱": name, "系列": ser, "設備類型": typ,
                              "故障次數": pp["faults"], "停機次數": pp["stops"],
                              "停工時數": pp["downtime"], "總維修時數": pp["repair"],
                              "平均MTBF": pp["mtbf"], "平均MTTR": pp["mttr"],
                              "平均可用率": pp["avail"],
                              "重複故障月數": int(g["重複故障"].sum())})
            agg = pd.DataFrame(_rows, columns=["機台代碼","機台名稱","系列","設備類型","故障次數",
                                               "停機次數","停工時數","總維修時數","平均MTBF",
                                               "平均MTTR","平均可用率","重複故障月數"])
            for _c in ["平均MTBF", "平均MTTR", "平均可用率"]:
                agg[_c] = pd.to_numeric(agg[_c], errors="coerce")
            agg = agg.sort_values(["停工時數","故障次數"], ascending=False)
            st.caption("MTBF／MTTR／可用率為選定期間合計：MTBF＝Σ(月時數−停工)÷Σ故障；無故障月份的運作時數也計入。")

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
                   }, na_rep="無故障"),
                use_container_width=True, hide_index=True)

            st.markdown("---")
            sel_m = st.selectbox("查看單台月別明細", ["— 選擇機台 —"] + list(agg["機台代碼"]))
            if sel_m != "— 選擇機台 —":
                dm = dk[dk["機台代碼"] == sel_m].sort_values(
                    "月份", key=lambda x: x.map({m:i for i,m in enumerate(MONTH_ORDER)}))
                float_cols_dm = ["停工時數(h)", "MTBF(h)", "MTTR(h)", "可用率(%)"]
                st.dataframe(
                    dm[["月份","故障次數","停機次數","停工時數(h)","MTBF(h)","MTTR(h)","可用率(%)"]
                      ].style.format({c: "{:.2f}" for c in float_cols_dm}, na_rep="無故障"),
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
                top10 = top10[top10["停工時數(h)"] > 0]
                fig2 = go.Figure(go.Bar(
                    x=top10["停工時數(h)"], y=top10["機台代碼"],
                    orientation="h", text=top10["停工時數(h)"],
                    marker_color="#d64541", opacity=0.85, textangle=0, textposition="auto",
                    hovertext=top10["機台名稱"]))
                fig2.update_layout(**PLOT_LAYOUT, height=340,
                    yaxis=dict(autorange="reversed", gridcolor='#e4e6ea'),
                    title=dict(text="停工時數 TOP 10", font_size=13, x=0))
                st.plotly_chart(fig2, use_container_width=True)

            with c2:
                st.markdown('<div class="section-title">🔁 Top 10 故障次數（機台）</div>',
                            unsafe_allow_html=True)
                top10f = dk.groupby(["機台代碼","機台名稱"])["故障次數"].sum()\
                           .reset_index().nlargest(10,"故障次數")
                top10f = top10f[top10f["故障次數"] > 0]
                fig3 = go.Figure(go.Bar(
                    x=top10f["故障次數"], y=top10f["機台代碼"],
                    orientation="h", text=top10f["故障次數"],
                    marker_color="#2563c9", opacity=0.85, textangle=0, textposition="auto",
                    customdata=top10f["機台名稱"],
                    hovertemplate="%{y}<br>%{customdata}<br>故障 %{x} 次<br><i>點擊查看故障原因</i><extra></extra>"))
                fig3.update_layout(**PLOT_LAYOUT, height=340,
                    yaxis=dict(autorange="reversed", gridcolor='#e4e6ea'),
                    title=dict(text="故障次數 TOP 10（點長條看故障原因）", font_size=13, x=0),
                    clickmode="event+select")
                ev = st.plotly_chart(fig3, use_container_width=True,
                                     on_select="rerun", selection_mode="points",
                                     key="top10_fault_chart")

                # 點選後帶出故障原因
                pts = (ev.get("selection", {}) or {}).get("points", []) if ev else []
                if pts:
                    sel_code = pts[0].get("y")
                    sel_name = SERIES_NAME.get(
                        re.sub(r"^(M-\d{3}-\d{3}).*$", r"\1", str(sel_code)), "")
                    recs = dr[dr["機台代碼"] == sel_code].copy()
                    st.markdown(f"**🔍 {sel_code}　{sel_name}　故障原因（{len(recs)} 筆）**")
                    if recs.empty or "需求" not in recs.columns:
                        st.info("無故障原因資料")
                    else:
                        recs["_o"] = recs["月份"].map(
                            {m: i for i, m in enumerate(MONTH_ORDER)})
                        show_cols = [c for c in ["月份", "需求", "維修時數", "停機"]
                                     if c in recs.columns]
                        out = recs.sort_values("_o")[show_cols].rename(
                            columns={"需求": "故障原因", "維修時數": "維修時數(h)"})
                        out["故障原因"] = out["故障原因"].fillna("無資料").replace("", "無資料")
                        if "停機" in out.columns:
                            out["停機"] = out["停機"].map(
                                lambda v: "是" if v in (1, True, "1", "是", "Y") else "否")
                        fmt = {"維修時數(h)": "{:.2f}"} if "維修時數(h)" in out.columns else {}
                        st.dataframe(out.style.format(fmt, na_rep="無資料"),
                                     use_container_width=True, hide_index=True)
                else:
                    st.caption("👆 點選上方長條即可查看該機台的故障原因")

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
                                    text=shortage_m["缺料件數"], marker_color="#f0a020", opacity=0.85))
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
                                 line=dict(color="#2563c9", width=2), marker=dict(size=7))
                fig6.add_scatter(x=ds2["月份"], y=ds2["平均MTTR(h)"],
                                 name="平均MTTR(h)", mode="lines+markers",
                                 line=dict(color="#f0a020", width=2), marker=dict(size=7),
                                 yaxis="y2")
                fig6.add_hline(y=THRESHOLD_MTBF, line_dash="dash", line_color="#2563c9",
                               annotation_text=f"MTBF目標 {THRESHOLD_MTBF}h")
                fig6.update_layout(**PLOT_LAYOUT,
                    yaxis=dict(title="MTBF(h)", gridcolor='#e4e6ea'),
                    yaxis2=dict(title="MTTR(h)", overlaying="y", side="right"),
                    legend=dict(x=0, y=1.08, orientation="h"), height=320)
                st.plotly_chart(fig6, use_container_width=True)

            with c2:
                st.markdown('<div class="section-title">平均可用率趨勢</div>',
                            unsafe_allow_html=True)
                fig7 = go.Figure()
                fig7.add_scatter(x=ds2["月份"], y=ds2["平均可用率(%)"],
                                 mode="lines+markers",
                                 line=dict(color="#2e9d57", width=2), marker=dict(size=7),
                                 fill="tozeroy", fillcolor="rgba(29,158,117,0.08)")
                fig7.add_hline(y=THRESHOLD_AVAIL, line_dash="dash", line_color="#f0a020",
                               annotation_text=f"目標 {THRESHOLD_AVAIL}%")
                fig7.update_layout(**PLOT_LAYOUT, height=320,
                    yaxis=dict(title="可用率(%)", range=[0,110], gridcolor='#e4e6ea'))
                st.plotly_chart(fig7, use_container_width=True)

            # 熱圖 — 分三大類
            st.markdown('<div class="section-title">機台系列 × 月份 — 可用率熱圖</div>',
                        unsafe_allow_html=True)
            # 色階說明
            st.markdown(
                '<div style="display:flex;gap:16px;margin:4px 0 12px;font-size:13px;">'
                '<span style="background:#fcebeb;padding:3px 10px;border-radius:4px;'
                'color:#a32d2d;font-weight:600;">🟥 ＜80%　警示</span>'
                '<span style="background:#fff2cc;padding:3px 10px;border-radius:4px;'
                'color:#7b4f00;font-weight:600;">🟨 80–94%　偏低</span>'
                '<span style="background:#e2efda;padding:3px 10px;border-radius:4px;'
                'color:#1f7a4d;font-weight:600;">🟩 ≥95%　良好</span>'
                '<span style="color:#6b7684;padding:3px 0;">（格內數字為平均可用率 %）</span>'
                '</div>',
                unsafe_allow_html=True)
            TYPE_COLOR_HM = {"成型類": "#2563c9", "加工類": "#2e9d57", "包裝類": "#f0a020"}
            for eq_type in EQ_TYPE_ORDER:
                dk_hm = dk[dk["設備類型"] == eq_type]
                if dk_hm.empty:
                    continue
                type_color = TYPE_COLOR_HM.get(eq_type, "#2563c9")
                st.markdown(
                    f'<div class="type-header"><i style="background:{type_color}"></i>'
                    f'{eq_type}</div>', unsafe_allow_html=True)
                _g = dk_hm.groupby(["系列","月份"])[["運作時數(h)","月可用時數(h)"]].sum()
                _g["可用率"] = _g["運作時數(h)"] / _g["月可用時數(h)"] * 100
                pivot = _g["可用率"].unstack("月份").round(2)
                ordered_cols = [c for c in MONTH_ORDER if c in pivot.columns]
                pivot = pivot[ordered_cols]
                # 短月份標籤
                pivot.columns = [c.replace("2026-", "") for c in pivot.columns]
                fig_hm = px.imshow(pivot, text_auto=".1f", aspect="auto",
                                   color_continuous_scale=["#fcebeb","#fff2cc","#e2efda"],
                                   zmin=70, zmax=100)
                fig_hm.update_layout(
                    height=max(200, len(pivot) * 38 + 60),
                    margin=dict(t=10, b=10, l=130, r=60),
                    paper_bgcolor='white', plot_bgcolor='white',
                    coloraxis_showscale=False,
                    font=dict(size=13),
                )
                fig_hm.update_xaxes(side="bottom", tickangle=0, tickfont=dict(size=13))
                fig_hm.update_yaxes(tickfont=dict(size=13))
                st.plotly_chart(fig_hm, use_container_width=True)

        # ══ Tab5：月度分析（系列合計＋機台個別）══════════════════════════
        with tab5:
            render_monthly_analysis(df_kpi, df_raw, months_avail, excluded)

        st.markdown("---")
        st.caption("富強醫材股份有限公司 設備課　© 2026")

if __name__ == "__main__":
    main()
