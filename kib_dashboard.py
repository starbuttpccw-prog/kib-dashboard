import streamlit as st
import io
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

st.set_page_config(page_title="KIB Dashboard", page_icon="📊")

st.title("📊 KIB 統一管理")

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
    st.error("❌ 連接失敗")
    st.stop()

st.write("---")
st.write("App 正常運作中。")
st.write("跟進表讀取功能暫時停用，避免錯誤。")
st.info("如果見到呢個頁面，代表 app 已經成功開到。")
