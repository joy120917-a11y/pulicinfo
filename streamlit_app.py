"""
公開資訊觀測站 精華版查詢 — Streamlit 版
執行方式：streamlit run mops_streamlit.py
"""

import io
import warnings

import pandas as pd
import requests
import streamlit as st
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore")

# ── 常數 ──────────────────────────────────────────────────────
BASE_URL  = "https://mopsov.twse.com.tw"
QUERY_URL = f"{BASE_URL}/mops/web/ajax_t146sb05"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer":      f"{BASE_URL}/mops/web/t146sb05",
    "Content-Type": "application/x-www-form-urlencoded",
    "Origin":       BASE_URL,
}

FORM_BASE = {
    "step":      "1",
    "firstin":   "true",
    "off":       "1",
    "keyword4":  "",
    "code1":     "",
    "TYPEK2":    "",
    "checkbtn":  "",
    "queryName": "co_id",
    "inpuType":  "co_id",
    "TYPEK":     "all",
}

# ── 查詢函式 ──────────────────────────────────────────────────
def fetch_mops(co_id: str) -> list[pd.DataFrame]:
    """POST 至 MOPS，解析所有 HTML 表格，回傳 list[DataFrame]。"""
    payload = {**FORM_BASE, "co_id": co_id.strip()}
    session = requests.Session()
    session.get(
        f"{BASE_URL}/mops/web/t146sb05",
        headers=HEADERS, timeout=15, verify=False,
    )
    resp = session.post(
        QUERY_URL, data=payload,
        headers=HEADERS, timeout=20, verify=False,
    )
    resp.raise_for_status()
    resp.encoding = "utf-8"

    soup = BeautifulSoup(resp.text, "html.parser")
    dfs = []
    for tbl in soup.find_all("table"):
        rows = tbl.find_all("tr")
        if len(rows) < 2:
            continue
        data = [
            [c.get_text(strip=True) for c in r.find_all(["th", "td"])]
            for r in rows
        ]
        max_col = max(len(r) for r in data)
        data = [r + [""] * (max_col - len(r)) for r in data]
        header, *body = data
        if not any(header):
            continue
        df = pd.DataFrame(body, columns=header)
        df = (
            df.loc[:, df.columns != ""]
            .dropna(how="all")
            .reset_index(drop=True)
        )
        if not df.empty:
            dfs.append(df)
    return dfs


# ── Excel 匯出輔助 ────────────────────────────────────────────
def to_excel_bytes(dfs: list[pd.DataFrame]) -> bytes:
    """將多張 DataFrame 寫入 Excel，回傳 bytes。"""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        for i, df in enumerate(dfs, 1):
            df.to_excel(writer, sheet_name=f"表格{i}", index=False)
    return buf.getvalue()


# ── 頁面設定 ──────────────────────────────────────────────────
st.set_page_config(
    page_title="公開資訊觀測站查詢",
    page_icon="📊",
    layout="wide",
)

# ── 自訂 CSS ──────────────────────────────────────────────────
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@400;500;700&display=swap');

    html, body, [class*="css"] {
        font-family: 'Noto Sans TC', sans-serif;
    }

    /* 頂部 Banner */
    .banner {
        background: linear-gradient(135deg, #1565C0, #0D47A1);
        padding: 24px 32px;
        border-radius: 14px;
        margin-bottom: 24px;
    }
    .banner h1 { color: white; margin: 0; font-size: 1.8rem; }
    .banner p  { color: #BBDEFB; margin: 6px 0 0; font-size: 0.95rem; }

    /* 查詢卡片 */
    .query-card {
        background: #F8FBFF;
        border: 1px solid #BBDEFB;
        border-radius: 10px;
        padding: 20px 24px;
        margin-bottom: 20px;
    }

    /* 表格標題 */
    .table-title {
        color: #1565C0;
        font-size: 1rem;
        font-weight: 700;
        margin: 16px 0 6px;
        border-left: 4px solid #1565C0;
        padding-left: 8px;
    }

    /* 成功 / 警告 訊息 */
    .msg-success { color: #2E7D32; font-weight: 600; }
    .msg-warn    { color: #E65100; font-weight: 600; }

    /* dataframe 表頭顏色覆寫 */
    thead tr th {
        background-color: #1565C0 !important;
        color: white !important;
    }

    /* 按鈕美化 */
    div.stButton > button {
        border-radius: 8px;
        height: 2.5rem;
        font-size: 0.95rem;
        font-weight: 600;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ── 頁面 Banner ───────────────────────────────────────────────
st.markdown(
    """
    <div class="banner">
      <h1>📊 公開資訊觀測站 精華版查詢</h1>
      <p>Taiwan Stock Exchange MOPS — Company Information Query</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# ── Session State 初始化 ───────────────────────────────────────
if "dfs" not in st.session_state:
    st.session_state.dfs = []
if "last_co_id" not in st.session_state:
    st.session_state.last_co_id = ""

# ── 查詢區塊 ──────────────────────────────────────────────────
with st.container():
    col_input, col_btn, col_export = st.columns([3, 1, 1.4], gap="small")

    with col_input:
        co_id = st.text_input(
            "股票代號",
            value="2330",
            placeholder="輸入股票代號或公司簡稱，例如：2330 或 台積電",
            label_visibility="collapsed",
        )

    with col_btn:
        query_clicked = st.button("🔍  查詢", use_container_width=True, type="primary")

    with col_export:
        export_clicked = st.button("📥  匯出 Excel", use_container_width=True)

st.markdown("<hr style='border:1px solid #e0e0e0;margin:4px 0 16px;'>", unsafe_allow_html=True)

# ── 執行查詢 ──────────────────────────────────────────────────
if query_clicked:
    if not co_id.strip():
        st.warning("⚠️ 請輸入股票代號或公司簡稱")
    else:
        with st.spinner(f"🔍 正在查詢「{co_id}」，請稍候…"):
            try:
                dfs = fetch_mops(co_id)
                st.session_state.dfs = dfs
                st.session_state.last_co_id = co_id.strip()
            except Exception as e:
                st.error(f"❌ 查詢失敗：{e}")
                st.session_state.dfs = []

# ── 顯示結果 ──────────────────────────────────────────────────
dfs = st.session_state.dfs
last_co_id = st.session_state.last_co_id

if dfs:
    st.success(f"✅ 「{last_co_id}」查詢完成，共取得 {len(dfs)} 張表格")

    for i, df in enumerate(dfs, 1):
        st.markdown(f"<div class='table-title'>表格 {i}</div>", unsafe_allow_html=True)

        # 使用 st.dataframe 顯示（可排序、可捲動）
        st.dataframe(
            df,
            use_container_width=True,
            hide_index=True,
        )

elif last_co_id:
    # 查詢過但沒資料
    st.warning(f"⚠️「{last_co_id}」查無資料，請確認代號是否正確")

# ── 匯出 Excel ────────────────────────────────────────────────
if export_clicked:
    if not dfs:
        st.warning("⚠️ 尚無查詢結果可匯出，請先執行查詢")
    else:
        excel_bytes = to_excel_bytes(dfs)
        filename = f"mops_{last_co_id}.xlsx"
        st.download_button(
            label="⬇️ 點此下載 Excel",
            data=excel_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.success(f"📥 Excel 已準備好：{filename}")

# ── 頁尾 ──────────────────────────────────────────────────────
st.markdown(
    """
    <div style="text-align:center;color:#9E9E9E;font-size:0.8rem;margin-top:40px;">
        資料來源：<a href="https://mops.twse.com.tw" target="_blank"
        style="color:#1565C0;">公開資訊觀測站 (MOPS)</a>
        ｜本工具僅供參考，投資請自行判斷
    </div>
    """,
    unsafe_allow_html=True,
)
