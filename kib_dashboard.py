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
def create_invoice_sheet(os_no, site, work, amount, qc_no="", qty=1):
    """下載 KB260001，抄 sheet，填資料，上傳"""
    try:
        request = service.files().get_media(fileId=KB_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        wb = openpyxl.load_workbook(fh)
        
        # 搵最新正式 KB26xxxx sheet
        import re
        candidates = [name for name in wb.sheetnames if re.match(r'^KB26\d{4}$', name)]
        if not candidates:
            return False, "搵唔到有效 template"
        
        template = sorted(candidates)[-1]
        last_num = int(template[2:])
        new_name = f"KB{last_num + 1}"
        
        st.info(f"Template: {template} → 新編號: {new_name}")

        # 抄 sheet
        source = wb[template]
        target = wb.copy_worksheet(source)
        target.title = new_name

        # 計算 unit rate
        unit_rate = round(amount / qty, 2) if qty else amount

        # 填資料
        for row in target.iter_rows(min_row=1, max_row=20, max_col=10):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if "Invoice No" in cell.value:
                        target.cell(row=cell.row, column=cell.column + 1).value = new_name
                    if "Your Order No" in cell.value:
                        target.cell(row=cell.row, column=cell.column + 1).value = os_no
                    if "Quotation No" in cell.value:
                        target.cell(row=cell.row, column=cell.column + 1).value = qc_no
                    if "Project" in cell.value:
                        target.cell(row=cell.row, column=cell.column + 1).value = f"{os_no} - {site}"

        # 填 item
        target['B19'] = work
        target['D19'] = qty
        target['F19'] = unit_rate
        target['G19'] = amount

        # 確保 Summary 永遠喺最後
        if 'Summary' in wb.sheetnames:
            current_idx = wb.sheetnames.index('Summary')
            offset = len(wb.sheetnames) - 1 - current_idx
            if offset > 0:
                wb.move_sheet('Summary', offset=offset)

        # 儲存上傳
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            tmp_path = tmp.name


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
    
    os_no = st.text_input("OS 編號", placeholder="OS2600xxxx")
    
    if st.button("搜尋並建立發票", type="primary", use_container_width=True):
        if not os_no:
            st.error("請輸入 OS 編號")
        elif df.empty:
            st.error("讀唔到跟進表")
        else:
            # 從跟進表搜尋
            search_str = os_no.lower().strip()
            mask = df.astype(str).apply(
                lambda col: col.str.lower().str.contains(search_str, na=False)
            ).any(axis=1)
            found = df[mask]
            
            if found.empty:
                st.warning("跟進表搵唔到呢個 OS")
            else:
                st.success(f"找到 {len(found)} 筆記錄")
                st.dataframe(found, use_container_width=True)
                
                # 嘗試提取資料（取第一筆）
                row = found.iloc[0]
                
                # 嘗試搵欄位
                site = ""
                work = ""
                amount = 0
                qc_no = ""
                
                for col in found.columns:
                    val = str(row[col]) if pd.notna(row[col]) else ""
                    if any(k in col.lower() for k in ["地盤", "site", "project"]):
                        site = val
                    if any(k in col.lower() for k in ["工作", "描述", "desc", "content"]):
                        work = val
                    if any(k in col.lower() for k in ["金額", "amount", "價錢"]):
                        try:
                            amount = float(val.replace(",", "").replace("$", ""))
                        except:
                            pass
                    if any(k in col.lower() for k in ["qc", "報價"]):
                        qc_no = val
                
                st.write(f"**地盤：** {site or '未搵到'}")
                st.write(f"**工作：** {work or '未搵到'}")
                st.write(f"**金額：** {amount}")
                st.write(f"**QC：** {qc_no or '未搵到'}")
                
                if amount > 0 and work:
                    with st.spinner("正在建立發票..."):
                        success, result = create_invoice_sheet(os_no, site, work, amount, qc_no)
                    if success:
                        st.success(f"已建立：{result}")
                    else:
                        st.error(f"失敗：{result}")
                else:
                    st.warning("資料唔完整，請手動檢查跟進表")

st.sidebar.caption(f"更新：{datetime.now().strftime('%H:%M')}")
st.sidebar.caption("同事可用此連結")
