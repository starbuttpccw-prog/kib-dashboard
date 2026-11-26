import streamlit as st
import pandas as pd
from datetime import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import io
import openpyxl
from openpyxl.utils import get_column_letter
import tempfile
import os

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
</style>
""", unsafe_allow_html=True)

st.title("🏢 KIB 統一管理")
st.caption("祈保工程 · 跟進表 / 報價 / 發票")

# ========== Config ==========
FOLLOWUP_FILE_ID = "1o430tei9i5SrDqvvhe3TpUDx7icGRmXf"
KB_FILE_ID = "17xlaPUUkWS0tE_L0aC0Sj4vF-_QuJ-Ql"  # 請更新成最新 KB260001 file_id

def get_credentials():
    if "gcp_service_account" in st.secrets:
        return Credentials.from_service_account_info(
            st.secrets["gcp_service_account"],
            scopes=["https://www.googleapis.com/auth/drive"]
        )
    return None

creds = get_credentials()
if creds is None:
    st.error("未設定 Google Drive 連接")
    st.stop()

service = build("drive", "v3", credentials=creds)

# ========== Load Follow-up ==========
@st.cache_data(ttl=180)
def load_followup():
    try:
        request = service.files().get_media(fileId=FOLLOWUP_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        xl = pd.ExcelFile(fh)
        sheet_names = xl.sheet_names
        fh.seek(0)

        target_sheet = next((n for n in sheet_names if "03" in str(n) or "等" in str(n)), sheet_names[0] if sheet_names else None)
        st.sidebar.caption(f"Sheets: {', '.join(sheet_names)}")
        st.sidebar.caption(f"讀取: {target_sheet}")

        preview = pd.read_excel(fh, sheet_name=target_sheet, header=None, nrows=15)
        fh.seek(0)
        header_row = 0
        for i, row in preview.iterrows():
            text = " ".join(str(v) for v in row if pd.notna(v))
            if any(k in text for k in ["QC", "OS", "地盤", "金額"]):
                header_row = i
                break

        df = pd.read_excel(fh, sheet_name=target_sheet, header=header_row)
        df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"讀取跟進表失敗：{e}")
        return pd.DataFrame()

df = load_followup()

# ========== Invoice Creation ==========
def create_invoice_sheet(os_no, site, work, amount, qc_no=""):
    """下載 KB260001，抄 sheet，填資料，上傳"""
    try:
        # 下載
        request = service.files().get_media(fileId=KB_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        wb = openpyxl.load_workbook(fh)
        
        # 搵最新嘅正式發票 sheet（KB26xxxx 格式，排除 Summary）
        template = None
        candidates = [name for name in wb.sheetnames if name.startswith("KB26") and name != "Summary" and len(name) <= 10]
        if candidates:
            template = sorted(candidates)[-1]  # 最新嘅
        else:
            template = wb.sheetnames[-2] if len(wb.sheetnames) > 1 else wb.sheetnames[-1]

        st.info(f"使用 template: {template}")

        # 新 sheet 名
        new_name = f"KB{datetime.now().strftime('%y%m%d%H%M')}"
        if new_name in wb.sheetnames:
            new_name += "a"
        
        # 抄 sheet
        source = wb[template]
        target = wb.copy_worksheet(source)
        target.title = new_name

        # 填資料（根據實際結構）
        # Invoice No
        for row in target.iter_rows(min_row=1, max_row=15, max_col=10):
            for cell in row:
                if cell.value and "Invoice No" in str(cell.value):
                    # 右邊一格填新編號
                    target.cell(row=cell.row, column=cell.column + 1).value = new_name
                if cell.value and "Your Order No" in str(cell.value):
                    target.cell(row=cell.row, column=cell.column + 1).value = os_no
                if cell.value and "Quotation No" in str(cell.value):
                    target.cell(row=cell.row, column=cell.column + 1).value = qc_no
                if cell.value and "Project" in str(cell.value):
                    target.cell(row=cell.row, column=cell.column + 1).value = f"{os_no} - {site}"

        # 簡單填描述同金額（第19行附近）
        target['B19'] = work
        target['G19'] = amount

        # 儲存到臨時檔
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            tmp_path = tmp.name

        # 上傳覆蓋
        media = MediaFileUpload(tmp_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        service.files().update(fileId=KB_FILE_ID, media_body=media).execute()
        
        os.unlink(tmp_path)
        return True, new_name

    except Exception as e:
        return False, str(e)

# ========== Pages ==========
page = st.sidebar.radio("功能", ["總覽", "跟進表", "開報價", "開發票"])

if page == "總覽":
    st.header("📊 總覽")
    if not df.empty:
        st.metric("跟進表項目", len(df))
        st.dataframe(df.head(10), use_container_width=True)
    else:
        st.warning("讀唔到跟進表")

elif page == "跟進表":
    st.header("🔍 跟進表")
    search = st.text_input("搜尋")
    if not df.empty:
        if search:
            mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(search.lower(), na=False)).any(axis=1)
            st.dataframe(df[mask], use_container_width=True)
        else:
            st.dataframe(df.head(20), use_container_width=True)

elif page == "開報價":
    st.header("📄 開報價")
    st.info("開發票功能優先開發中，報價稍後加入")

elif page == "開發票":
    st.header("🧾 開發票")
    st.warning("實驗功能：會直接修改 KB260001.xlsx")
    
    with st.form("invoice_form"):
        os_no = st.text_input("OS 編號", placeholder="OS2600xxxx")
        site = st.text_input("地盤")
        work = st.text_input("工作內容")
        amount = st.number_input("金額", min_value=0.0, step=100.0)
        qc_no = st.text_input("報價編號（可選）", placeholder="QC260xxx")
        submitted = st.form_submit_button("建立發票 Sheet", type="primary")
        
        if submitted:
            if os_no and work and amount > 0:
                with st.spinner("正在建立發票..."):
                    success, result = create_invoice_sheet(os_no, site, work, amount, qc_no)
                if success:
                    st.success(f"已建立新 sheet：{result}")
                    st.info("請到 Google Drive 打開 KB260001.xlsx 檢查，之後用 Excel Print Bot 出 PDF")
                else:
                    st.error(f"失敗：{result}")
            else:
                st.error("請填 OS、工作內容同金額")

st.sidebar.caption(f"更新：{datetime.now().strftime('%H:%M')}")
st.sidebar.caption("同事可用此連結")
