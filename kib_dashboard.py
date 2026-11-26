import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
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
    .stButton > button { width: 100%; padding: 0.75rem; font-size: 1.05rem; border-radius: 12px; }
    .block-container { padding-top: 1.5rem; }
</style>
""", unsafe_allow_html=True)

st.title("🏢 KIB 統一管理")
st.caption("祈保工程 · 連接 Google Drive")

# ========== Google Drive 連接 ==========
def get_credentials():
    if "gcp_service_account" in st.secrets:
        return Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive.readonly"]
        )
    return None

creds = get_credentials()

if creds is None:
    st.warning("尚未設定 Google Drive 連接")
    st.stop()

st.success("已連接 Google Drive")

# Sidebar
page = st.sidebar.radio("前往", ["總覽", "跟進表搜尋", "開報價", "開發票", "Bot 狀態"])

# ========== 讀取 Excel 檔案 ==========
@st.cache_data(ttl=300)
def load_followup_table():
    """下載並讀取跟進表 Excel"""
    try:
        service = build("drive", "v3", credentials=creds)
        # 請將 file_id 換成你真實跟進表 ID
        file_id = "1o430tei9i5SrDqvvhe3TpUDx7icGRmXf"
        
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        
        fh.seek(0)
        # 讀取 Excel（預設讀第一個 sheet，或指定 sheet_name）
        df = pd.read_excel(fh, sheet_name=0)
        return df
    except Exception as e:
        st.error(f"讀取跟進表失敗：{e}")
        return pd.DataFrame()

if page == "總覽":
    st.header("📊 總覽")
    df = load_followup_table()
    if not df.empty:
        st.metric("跟進表行數", len(df))
        st.dataframe(df.head(20), use_container_width=True)
    else:
        st.info("暫時讀唔到跟進表，請檢查 file_id 同分享權限")

elif page == "跟進表搜尋":
    st.header("🔍 跟進表搜尋")
    df = load_followup_table()
    search = st.text_input("搜尋 QC / KB / OS / 地盤")
    if not df.empty and search:
        filtered = df[df.apply(lambda r: search.lower() in str(r).lower(), axis=1)]
        st.success(f"找到 {len(filtered)} 筆")
        st.dataframe(filtered, use_container_width=True)
    elif not df.empty:
        st.dataframe(df.head(50), use_container_width=True)

elif page == "開報價":
    st.header("📄 開新報價")
    st.info("此功能之後會寫入 QC260001.xlsx")
    with st.form("quote"):
        site = st.text_input("地盤")
        work = st.text_input("工作內容")
        amount = st.number_input("金額", min_value=0)
        if st.form_submit_button("建立", use_container_width=True):
            st.success("已記錄（正式版會寫入 Drive）")

elif page == "開發票":
    st.header("🧾 開發票")
    os_no = st.text_input("OS 編號")
    if st.button("準備開單", use_container_width=True):
        st.info("正式版會從跟進表分配 KB 並更新")

elif page == "Bot 狀態":
    st.header("🤖 Bot 狀態")
    st.info("Bot 觸發功能之後會整合")

st.sidebar.caption(f"更新：{datetime.now().strftime('%H:%M')}")
