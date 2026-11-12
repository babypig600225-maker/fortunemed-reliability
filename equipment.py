"""
設備清單（在用機台名冊）解析與報修代碼對應
- 只取 M-301 / M-302 / M-303 工作表
- 使用狀態只計：使用中、測試中、驗證中
"""
import io, re, unicodedata
import pandas as pd

ACTIVE_STATUS = ("使用中", "測試中", "驗證中")
EQ_SHEET_PREFIX = ("M-301", "M-302", "M-303")
CODE_RE = re.compile(r"^(M-\d{3}-\d{3}-\d+(?:-\d+)*)\s*(?:\((\d+)\s*[~～-]\s*(\d+)\))?$")


def _parse_date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    s = str(v).strip().replace(".", "-").replace("/", "-")
    d = pd.to_datetime(s, errors="coerce")
    return None if pd.isna(d) else d.strftime("%Y-%m-%d")


def parse_equipment_excel(file_bytes):
    """回傳在用機台 list[dict]：code, name, status, series, in_date"""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))
    out = []
    for sh in xl.sheet_names:
        if not sh.startswith(EQ_SHEET_PREFIX):
            continue
        df = pd.read_excel(xl, sh, header=None)
        # 找表頭列（含「設備代號」）
        hdr = next((i for i, r in df.iterrows()
                    if any(str(x).strip() == "設備代號" for x in r.values)), None)
        if hdr is None:
            continue
        cols = [str(x).strip() for x in df.iloc[hdr].values]
        df = df.iloc[hdr + 1:]
        df.columns = cols
        if "使用狀態" not in df.columns:
            continue
        for _, r in df.iterrows():
            raw = unicodedata.normalize("NFKC", str(r.get("設備代號", ""))).strip()
            status = str(r.get("使用狀態", "")).strip()
            if status not in ACTIVE_STATUS:
                continue
            m = CODE_RE.match(raw)
            if not m:
                continue
            base, a, b = m.group(1), m.group(2), m.group(3)
            codes = [f"{base}-{i}" for i in range(int(a), int(b) + 1)] if a else [base]
            name = str(r.get("設備名稱", "") or "").strip()
            for c in codes:
                out.append({
                    "code": c, "name": name, "status": status,
                    "series": re.match(r"(M-\d{3}-\d{3})", c).group(1),
                    "in_date": _parse_date(r.get("入廠日期")),
                })
    # 去重（保留第一筆）
    seen, uniq = set(), []
    for e in out:
        if e["code"] not in seen:
            seen.add(e["code"]); uniq.append(e)
    return uniq


def map_repair_code(code, roster_codes):
    """報修代碼 → 名冊代碼；不在名冊（停用/報廢等）回傳 None。
    規則：完全相同 → 本身；名冊有其上層代碼 → 上層；名冊有其子代碼 → 本身。"""
    if not isinstance(code, str) or not code:
        return None
    code = unicodedata.normalize("NFKC", code).strip()
    if code in roster_codes:
        return code
    parts = code.split("-")
    while len(parts) > 4:              # M-XXX-YYY-N 至少 4 段
        parts = parts[:-1]
        p = "-".join(parts)
        if p in roster_codes:
            return p
    pref = code + "-"
    if any(c.startswith(pref) for c in roster_codes):
        return code
    return None
