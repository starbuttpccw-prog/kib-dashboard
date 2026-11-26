import streamlit as st
import pandas as pd
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

st.set_page_config(
    page_title="KIB 統一管理",
    page_icon="🏢",
    layout="centered",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    .stButton > button { width: 100%; padding: 0.8rem; font-size: 1.05rem; border-radius: 12px; }
    .block-container { padding-top: 1rem; padding-bottom: 2rem; }
    div[data-testid="metric-container"] {
        background: #f8f9fa;
        border-radius: 12px;
        padding: 12px;
    }
</style>
""", unsafe_allow_html=True)

st.title("🏢 KIB 統一管理")
st.caption("祈保工程 · 跟進表 / 報價 / 發票")

# ========== Google Drive ==========
def get_credentials():
    if "gcp_service_account" in st.secrets:
        return Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
    return None

creds = get_credentials()
if creds is None:
    st.error("未設定 Google Drive 連接")
    st.stop()

@st.cache_data(ttl=180)
def load_followup():
    try:
        service = build("drive", "v3", credentials=creds)
        file_id = "1o430tei9i5SrDqvvhe3TpUDx7icGRmXf"
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        # 自動搵標題行
        preview = pd.read_excel(fh, sheet_name=0, header=None, nrows=20)
        fh.seek(0)
        header_row = 0
        for i, row in preview.iterrows():
            text = " ".join(str(v) for v in row if pd.notna(v))
            if any(k in text for k in ["QC", "OS", "地盤", "金額", "狀態", "工作"]):
                header_row = i
                break

        df = pd.read_excel(fh, sheet_name=0, header=header_row)
        df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"讀取失敗：{e}")
        return pd.DataFrame()

df = load_followup()

# ========== 頁面 ==========
page = st.sidebar.radio("功能", ["總覽", "跟進表", "開報價", "開發票"])

if page == "總覽":
    st.header("📊 總覽")
    if df.empty:
        st.warning("讀唔到跟進表")
    else:
        # 嘗試搵狀態欄
        status_col = None
        for c in df.columns:
            if any(k in str(c) for k in ["狀態", "status", "階段"]):
                status_col = c
                break

        col1, col2 = st.columns(2)
        col1.metric("總項目", len(df))
        if status_col:
            waiting = df[df[status_col].astype(str).str.contains("03|等", na=False)].shape[0]
            invoiced = df[df[status_col].astype(str).str.contains("04|已開", na=False)].shape[0]
            col2.metric("等 OS (03)", waiting)
            st.metric("已開未收 (04)", invoiced)
        else:
            st.info("未搵到狀態欄，顯示原始數據")

        st.subheader("最近項目")
        st.dataframe(df.head(15), use_container_width=True, height=300)

elif page == "跟進表":
    st.header("🔍 跟進表")
    if df.empty:
        st.warning("無數據")
    else:
        search = st.text_input("搜尋 QC / OS / 地盤 / 工作", placeholder="輸入關鍵字")
        if search:
            try:
                search_str = search.lower().strip()
                mask = df.astype(str).apply(
                    lambda col: col.str.lower().str.contains(search_str, na=False)
                ).any(axis=1)
                result = df[mask]
                st.success(f"找到 {len(result)} 筆")
                st.dataframe(result, use_container_width=True, height=400)
            except Exception as e:
                st.error(f"搜尋出錯：{e}")
        else:
            st.dataframe(df.head(30), use_container_width=True, height=400)

elif page == "開報價":
    st.header("📄 開新報價")
    with st.form("new_qc"):
        site = st.text_input("地盤")
        work = st.text_input("工作內容")
        amount = st.number_input("金額 (HK$)", min_value=0, step=100)
        person = st.selectbox("負責人", ["Carmen", "Peter", "Brandon", "Neo", "Paul", "Tammy", "Jason", "Sean"])
        submitted = st.form_submit_button("準備開報價", use_container_width=True)
        if submitted and site and work:
            st.success(f"已記錄：{site} - {work} (${amount:,})")
            st.info("下一步會寫入 QC260001 並呼叫 Excel Print Bot 出 PDF")

elif page == "開發票":
    st.header("🧾 開發票")
    os_no = st.text_input("OS 編號", placeholder="例如 OS26005161")
    if st.button("查詢並準備開單", use_container_width=True):
        if os_no and not df.empty:
            try:
                # 更安全嘅搜尋方式
                search_str = os_no.lower().strip()
                mask = df.astype(str).apply(
                    lambda col: col.str.lower().str.contains(search_str, na=False)
                ).any(axis=1)
                found = df[mask]
                if not found.empty:
                    st.success(f"找到 {len(found)} 筆相關記錄")
                    st.dataframe(found, use_container_width=True)
                    st.info("正式版會分配 KB 編號、出 PDF、更新跟進表 03→04")
                else:
                    st.warning("跟進表搵唔到呢個 OS")
            except Exception as e:
                st.error(f"搜尋出錯：{e}")
        else:
            st.warning("請輸入 OS 編號")

st.sidebar.divider()
st.sidebar.caption(f"更新時間：{datetime.now().strftime('%H:%M')}")
st.sidebar.caption("手機可直接開此頁面使用")
