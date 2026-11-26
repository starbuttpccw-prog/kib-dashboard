import streamlit as st
import pandas as pd
import io
import re
import tempfile
import os
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import openpyxl

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
        target_sheet = xl.sheet_names[0]
        header_row = 0
        
        for sheet in xl.sheet_names:
            for hdr in range(0, 8):
                try:
                    temp = pd.read_excel(fh, sheet_name=sheet, header=hdr, nrows=3)
                    cols = [str(c).lower() for c in temp.columns]
                    if any(k in " ".join(cols) for k in ["report", "報告", "編號", "site", "地盤", "os"]):
                        target_sheet = sheet
                        header_row = hdr
                        break
                except:
                    continue
            if target_sheet != xl.sheet_names[0]:
                break
        
        fh.seek(0)
        df = pd.read_excel(fh, sheet_name=target_sheet, header=header_row)
        df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
        df = df.astype(str).replace("nan", "").replace("None", "")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"讀取跟進表失敗：{e}")
        return pd.DataFrame()

FOLLOWUP_ID = "1o430tei9i5SrDqvvhe3TpUDx7icGRmXf"
KB_FILE_ID = "17xlaPUUkWS0tE_L0aC0Sj4vF-_QuJ-Ql"

service = get_drive_service()
if not service:
    st.stop()

st.title("📊 KIB 統一管理 Dashboard")
st.sidebar.success("✅ 已連接 Google Drive")

df = load_followup_table(FOLLOWUP_ID)

page = st.sidebar.radio("功能", ["總覽", "跟進表", "開發票"])

if page == "總覽":
    st.header("📊 總覽")
    if not df.empty:
        st.metric("跟進表項目", len(df))
        st.dataframe(df.head(15), width="stretch", height=300)
    else:
        st.warning("讀唔到跟進表")

elif page == "跟進表":
    st.header("🔍 跟進表")
    search = st.text_input("搜尋關鍵字")
    if not df.empty:
        if search:
            mask = df.astype(str).apply(
                lambda col: col.str.lower().str.contains(search.lower(), na=False)
            ).any(axis=1)
            result = df[mask]
            st.write(f"找到 {len(result)} 筆")
            st.dataframe(result.head(20), width="stretch")
        else:
            st.dataframe(df.head(20), width="stretch")

elif page == "開發票":
    st.header("🧾 開發票")
    st.warning("實驗功能：會修改 KB260001.xlsx")
    
    os_no = st.text_input("OS 編號", placeholder="OS2600xxxx")
    
    if st.button("搜尋並建立發票", type="primary"):
        if not os_no or df.empty:
            st.error("請輸入 OS 編號")
        else:
            mask = df.astype(str).apply(
                lambda col: col.str.lower().str.contains(os_no.lower(), na=False)
            ).any(axis=1)
            found = df[mask]
            
            if found.empty:
                st.warning("跟進表搵唔到")
            else:
                st.success(f"找到 {len(found)} 筆")
                st.dataframe(found.head(5), width="stretch")
                
                row = found.iloc[0]
                cols = list(found.columns)
                
                site = work = qc_no = ""
                amount = 0.0
                
                # 嘗試提取
                for i, col in enumerate(cols):
                    val = str(row[col]).strip()
                    if val.startswith("QC") and len(val) > 5:
                        qc_no = val
                    if val.replace(".", "").replace(",", "").isdigit():
                        try:
                            num = float(val.replace(",", ""))
                            if num > 100:
                                amount = num
                        except:
                            pass
                    if " - " in val and len(val) > 10:
                        parts = val.split(" - ", 1)
                        site = parts[0].strip()
                        work = parts[1].strip()
                
                st.write(f"**地盤：** {site or '未搵到'}")
                st.write(f"**工作：** {work or '未搵到'}")
                st.write(f"**金額：** {amount}")
                st.write(f"**QC：** {qc_no or '未搵到'}")
                
                if amount > 0:
                    st.info("發票建立功能暫時停用（避免複雜錯誤）。資料已提取，之後可以加返。")
                else:
                    st.warning("搵唔到金額")

st.caption("KIB Dashboard • 安全版")
