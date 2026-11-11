# 富強醫材 設備可靠性分析儀表板

上傳報修 Excel（含「時數」工作表）即可顯示完整分析（外觀與 HTML 版完全相同）。

## 部署到 Streamlit Cloud

1. 建立 GitHub Repo：`fortunemed-reliability`
2. 上傳此資料夾所有檔案
3. 至 streamlit.io/cloud 部署，Main file path 填 `app.py`
4. Settings → Secrets 填入：
   ```
   GITHUB_TOKEN = "你的token"
   ```

## 本機執行

```
pip install -r requirements.txt
streamlit run app.py
```

## 功能頁面

| 頁面 | 說明 |
|------|------|
| 📊 月別總覽 | 5大KPI卡片＋月別故障/停工圖 |
| 🔧 機台明細 | 每台MTBF/MTTR/可用率，可下鑽月份 |
| ⚠️ 故障分析 | Top10停工、重複故障、原因分類、缺料 |
| 📈 趨勢分析 | MTBF/MTTR折線、可用率趨勢、熱圖 |

## 計算說明

- **平日**：11小時（8h + 3h加班）
- **補班假日**：8小時
- **MTTR** = 總維修時數 ÷ 故障次數
- **MTBF** = （月可用時數 − 停工時數）÷ 故障次數
- **可用率** = （月可用時數 − 停工時數）÷ 月可用時數 × 100%
