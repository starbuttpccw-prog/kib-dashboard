import streamlit as st
import pandas as pd
import io
import re
import zipfile
import shutil
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

st.set_page_config(page_title="KIB 統一管理", page_icon="📊", layout="wide")

@st.cache_resource
def get_drive_service():
    try:
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]),
            scopes=["https://www.googleapis.com/auth/drive"]
        )
        return build("drive", "v3", credentials=creds)
    except Exception as e:
        st.error(f"連接失敗：{e}")
        return None

@st.cache_data(ttl=300)
def load_followup_table(file_id):
    service = get_drive_service()
    if not service:
        return pd.DataFrame()
    try:
        request = service.files().get_media(fileId=file_id)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)
        
        xl = pd.ExcelFile(fh)
        target_sheet = None
        header_row = 0
        for sheet in xl.sheet_names:
            for hdr in range(0, 10):
                try:
                    temp = pd.read_excel(fh, sheet_name=sheet, header=hdr, nrows=5)
                    cols = [str(c).lower() for c in temp.columns]
                    if any("report" in c or "報告" in c or "編號" in c or "site" in c or "地盤" in c for c in cols):
                        target_sheet = sheet
                        header_row = hdr
                        break
                except:
                    continue
            if target_sheet:
                break
        
        fh.seek(0)
        if target_sheet is None:
            target_sheet = xl.sheet_names[0]
        
        df = pd.read_excel(fh, sheet_name=target_sheet, header=header_row)
        df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
        # 強制全部轉字串，避免 Arrow 錯誤
        df = df.astype(str).replace("nan", "")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"讀取失敗：{e}")
        return pd.DataFrame()

FOLLOWUP_ID = "1o430tei9i5SrDqvvhe3TpUDx7icGRmXf"

st.title("📊 KIB 統一管理 Dashboard")

service = get_drive_service()
if service:
    st.sidebar.success("✅ 已連接 Google Drive")
else:
    st.sidebar.error("❌ 連接失敗")
    st.stop()

df = load_followup_table(FOLLOWUP_ID)

if df.empty:
    st.warning("暫時讀唔到跟進表")
else:
    st.metric("跟進表行數", len(df))
    
    # 只顯示前 20 行，避免大表格問題
    st.subheader("最近記錄")
    display_df = df.head(20)
    st.dataframe(display_df, width="stretch", height=300)

# 搜尋
st.subheader("🔍 搜尋")
keyword = st.text_input("輸入關鍵字")
if keyword and not df.empty:
    mask = df.astype(str).apply(lambda x: x.str.contains(keyword, case=False, na=False)).any(axis=1)
    result = df[mask]
    st.write(f"找到 {len(result)} 筆")
    if not result.empty:
        st.dataframe(result.head(10), width="stretch")

st.info("App 正常運作中。完整功能會逐步加返。")
