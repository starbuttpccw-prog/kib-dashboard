import streamlit as st
import pandas as pd
import io
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="KIB Dashboard", page_icon="📊")

st.title("📊 KIB 統一管理")

# 連接 Google Drive
@st.cache_resource
def get_service():
    try:
        info = dict(st.secrets["gcp_service_account"])
        creds = Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive"]
        )
        return build("drive", "v3", credentials=creds)
    except Exception as e:
        st.error(f"連接失敗：{e}")
        return None

service = get_service()

if service:
    st.success("✅ 已連接 Google Drive")
else:
    st.error("❌ 連接失敗，請檢查 secrets")
    st.stop()

# 讀取跟進表
@st.cache_data(ttl=300)
def load_data(file_id):
    try:
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        
        # 讀取第一個 sheet，全部轉字串
        df = pd.read_excel(fh, header=None)
        df = df.astype(str).replace("nan", "")
        return df
    except Exception as e:
        st.error(f"讀取失敗：{e}")
        return pd.DataFrame()

FILE_ID = "1o430tei9i5SrDqvvhe3TpUDx7icGRmXf"

with st.spinner("讀取跟進表..."):
    df = load_data(FILE_ID)

if df.empty:
    st.warning("讀唔到數據")
else:
    st.metric("總行數", len(df))
    st.write("前 10 行：")
    # 只顯示少量，避免 Arrow 問題
    st.dataframe(df.head(10), width="stretch", height=250)

st.info("App 正常運作。之後可以逐步加功能。")
