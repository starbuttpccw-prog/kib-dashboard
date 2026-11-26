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
    """用 zip 層面抄 sheet，保留 logo + 印"""
    try:
        import zipfile
        import shutil
        import re
        
        # 下載
        request = service.files().get_media(fileId=KB_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        # 先搵 template 同新編號
        wb_temp = openpyxl.load_workbook(fh)
        
        # 優先從 Summary 讀取最新編號
        new_name = None
        if 'Summary' in wb_temp.sheetnames:
            summary = wb_temp['Summary']
            # 搵 Summary 入面最大嘅 KB26xxxx
            max_num = 0
            for row in summary.iter_rows(min_row=1, max_row=50, max_col=5, values_only=True):
                for cell in row:
                    if cell and isinstance(cell, str) and re.match(r'^KB26\d{4}$', cell):
                        num = int(cell[2:])
                        if num > max_num:
                            max_num = num
            if max_num > 0:
                new_name = f"KB{max_num + 1}"
        
        # 後備：從 sheet 名搵
        if not new_name:
            candidates = [name for name in wb_temp.sheetnames if re.match(r'^KB26\d{4}$', name)]
            if candidates:
                template = sorted(candidates)[-1]
                last_num = int(template[2:])
                new_name = f"KB{last_num + 1}"
            else:
                return False, "搵唔到有效編號"
        
        # 搵一個有 logo 嘅 template（最新正式 sheet）
        candidates = [name for name in wb_temp.sheetnames if re.match(r'^KB26\d{4}$', name)]
        template = sorted(candidates)[-1] if candidates else wb_temp.sheetnames[-2]
        
        new_sheet_id = int(new_name[2:])
        st.info(f"Template: {template} → 新編號: {new_name}")

        unit_rate = round(amount / qty, 2) if qty else amount

        # Zip 層面抄
        fh.seek(0)
        with tempfile.TemporaryDirectory() as tmpdir:
            with zipfile.ZipFile(fh, 'r') as z:
                z.extractall(tmpdir)
            
            # 搵 template sheet path
            wb_xml = os.path.join(tmpdir, 'xl', 'workbook.xml')
            with open(wb_xml, 'r', encoding='utf-8') as f:
                wb_content = f.read()
            
            rid_match = re.search(rf'name="{re.escape(template)}"[^>]*r:id="([^"]+)"', wb_content)
            if not rid_match:
                return False, "搵唔到 sheet rid"
            rid = rid_match.group(1)
            
            rels_path = os.path.join(tmpdir, 'xl', '_rels', 'workbook.xml.rels')
            with open(rels_path, 'r', encoding='utf-8') as f:
                rels_content = f.read()
            
            target_match = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels_content)
            if not target_match:
                return False, "搵唔到 sheet path"
            sheet_rel_path = target_match.group(1)
            
            # 抄 sheet XML
            src_sheet = os.path.join(tmpdir, 'xl', sheet_rel_path)
            new_sheet_filename = f'sheet{new_sheet_id}.xml'
            dst_sheet = os.path.join(tmpdir, 'xl', 'worksheets', new_sheet_filename)
            shutil.copy2(src_sheet, dst_sheet)
            
            # 抄 drawing rels（如果有）
            src_rels = os.path.join(tmpdir, 'xl', 'worksheets', '_rels', os.path.basename(sheet_rel_path) + '.rels')
            if os.path.exists(src_rels):
                dst_rels = os.path.join(tmpdir, 'xl', 'worksheets', '_rels', new_sheet_filename + '.rels')
                os.makedirs(os.path.dirname(dst_rels), exist_ok=True)
                shutil.copy2(src_rels, dst_rels)
            
            # 更新 workbook.xml
            new_rid = f'rId{new_sheet_id}'
            sheet_entry = f'<sheet name="{new_name}" sheetId="{new_sheet_id}" r:id="{new_rid}"/>'
            wb_content = wb_content.replace('</sheets>', sheet_entry + '</sheets>')
            with open(wb_xml, 'w', encoding='utf-8') as f:
                f.write(wb_content)
            
            # 更新 rels
            new_rel = f'<Relationship Id="{new_rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{new_sheet_filename}"/>'
            rels_content = rels_content.replace('</Relationships>', new_rel + '</Relationships>')
            with open(rels_path, 'w', encoding='utf-8') as f:
                f.write(rels_content)
            
            # 重新打包
            result_path = os.path.join(tmpdir, 'result.xlsx')
            with zipfile.ZipFile(result_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for root, dirs, files in os.walk(tmpdir):
                    for file in files:
                        if file == 'result.xlsx':
                            continue
                        full_path = os.path.join(root, file)
                        arcname = os.path.relpath(full_path, tmpdir)
                        zout.write(full_path, arcname)
            
            # 用 openpyxl 填資料
            wb = openpyxl.load_workbook(result_path)
            target = wb[new_name]
            
            for r in target.iter_rows(min_row=1, max_row=20, max_col=10):
                for cell in r:
                    if cell.value and isinstance(cell.value, str):
                        if "Invoice No" in cell.value:
                            target.cell(row=cell.row, column=cell.column + 1).value = new_name
                        if "Your Order No" in cell.value:
                            target.cell(row=cell.row, column=cell.column + 1).value = os_no
                        if "Quotation No" in cell.value:
                            target.cell(row=cell.row, column=cell.column + 1).value = qc_no
                        if "Project" in cell.value:
                            target.cell(row=cell.row, column=cell.column + 1).value = f"{os_no} - {site}"
            
            target['B19'] = work
            target['D19'] = qty
            target['F19'] = unit_rate
            target['G19'] = amount
            
            # Summary 去最後
            if 'Summary' in wb.sheetnames:
                idx = wb.sheetnames.index('Summary')
                if idx < len(wb.sheetnames) - 1:
                    wb.move_sheet('Summary', offset=len(wb.sheetnames) - 1 - idx)
            
            wb.save(result_path)
            
            # 上傳
            media = MediaFileUpload(result_path, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            service.files().update(fileId=KB_FILE_ID, media_body=media).execute()
            
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
                
                # 取第一筆
                row = found.iloc[0]
                
                # 根據實際結構提取（Unnamed 欄位）
                site = ""
                work = ""
                amount = 0.0
                qc_no = ""
                
                # 嘗試用位置提取
                cols = list(found.columns)
                try:
                    # Unnamed: 4 通常係 QC
                    if len(cols) > 4 and pd.notna(row[cols[4]]):
                        val = str(row[cols[4]])
                        if val.upper().startswith("QC"):
                            qc_no = val
                    # Unnamed: 6 通常係 地盤 - 工作
                    if len(cols) > 6 and pd.notna(row[cols[6]]):
                        val = str(row[cols[6]])
                        if " - " in val:
                            parts = val.split(" - ", 1)
                            site = parts[0].strip()
                            work = parts[1].strip()
                        else:
                            work = val
                    # Unnamed: 8 係純工作
                    if len(cols) > 8 and pd.notna(row[cols[8]]) and not work:
                        work = str(row[cols[8]])
                    # Unnamed: 9 係金額
                    if len(cols) > 9 and pd.notna(row[cols[9]]):
                        try:
                            amount = float(str(row[cols[9]]).replace(",", ""))
                        except:
                            pass
                except Exception as e:
                    st.warning(f"提取出錯：{e}")
                
                # 後備：掃描所有值
                if not qc_no or amount == 0:
                    for col in cols:
                        val = row[col]
                        if pd.isna(val):
                            continue
                        val_str = str(val).strip()
                        if not qc_no and val_str.upper().startswith("QC"):
                            qc_no = val_str
                        if amount == 0 and val_str.replace(".", "").replace(",", "").isdigit():
                            try:
                                num = float(val_str.replace(",", ""))
                                if num > 100:
                                    amount = num
                            except:
                                pass
                        if not work and len(val_str) > 15 and "battery" in val_str.lower():
                            work = val_str

                
                st.write(f"**地盤：** {site or '未搵到'}")
                st.write(f"**工作：** {work or '未搵到'}")
                st.write(f"**金額：** {amount}")
                st.write(f"**QC：** {qc_no or '未搵到'}")
                
                # 顯示原始行方便確認
                with st.expander("原始數據"):
                    st.write(row.to_dict())
                
                # 只要有金額就允許建立（用 OS 做描述後備）
                if amount > 0:
                    final_work = work or f"Work for {os_no}"
                    final_site = site or os_no
                    with st.spinner("正在建立發票..."):
                        success, result = create_invoice_sheet(os_no, final_site, final_work, amount, qc_no)
                    if success:
                        st.success(f"已建立：{result}")
                        st.info("請到 Drive 檢查，之後用 Excel Print Bot 出 PDF")
                    else:
                        st.error(f"失敗：{result}")
                else:
                    st.warning("搵唔到金額，請手動檢查跟進表")

st.sidebar.caption(f"更新：{datetime.now().strftime('%H:%M')}")
st.sidebar.caption("同事可用此連結")
