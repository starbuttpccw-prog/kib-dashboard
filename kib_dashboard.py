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
        
        # 試多個 sheet 同 header
        xl = pd.ExcelFile(fh)
        best_df = None
        best_score = 0
        
        for sheet in xl.sheet_names:
            for hdr in range(0, 10):
                try:
                    fh.seek(0)
                    try:
                        fh.seek(0)
                        # 安全讀取最多 2000 行
                        temp = pd.read_excel(fh, sheet_name=sheet, header=hdr, nrows=2000)
                    except Exception:
                        continue
                    if temp.empty:
                        continue
                    # 計分：有幾多欄包含有用字
                    score = 0
                    cols_str = " ".join([str(c).lower() for c in temp.columns])
                    if any(k in cols_str for k in ["report", "報告", "編號", "site", "地盤", "os", "qc", "金額"]):
                        score += 5
                    # 睇數據入面有幾多 OS / QC
                    for col in temp.columns:
                        vals = temp[col].astype(str).str.upper()
                        score += vals.str.contains("OS26|QC26", na=False).sum()
                    if score > best_score:
                        best_score = score
                        best_df = temp.copy()
                        best_df.attrs['sheet'] = sheet
                        best_df.attrs['header'] = hdr
                except:
                    continue
        
        if best_df is None or best_df.empty:
            # 後備：讀第一個 sheet
            fh.seek(0)
            best_df = pd.read_excel(fh, header=None)
        
        best_df = best_df.dropna(axis=1, how="all").dropna(axis=0, how="all")
        best_df = best_df.astype(str).replace("nan", "").replace("None", "")
        best_df.columns = [str(c).strip() for c in best_df.columns]
        return best_df
    except Exception as e:
        st.error(f"讀取跟進表失敗：{e}")
        return pd.DataFrame()

FOLLOWUP_ID = "1o430tei9i5SrDqvvhe3TpUDx7icGRmXf"
KB_FILE_ID = "17xlaPUUkWS0tE_L0aC0Sj4vF-_QuJ-Ql"

def create_invoice_sheet(os_no, site, work, amount, qc_no="", qty=1):
    """輕量 zip 抄 sheet，盡量保留 logo + 印"""
    try:
        import zipfile
        import shutil
        
        request = service.files().get_media(fileId=KB_FILE_ID)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        fh.seek(0)

        # 先搵編號
        wb_temp = openpyxl.load_workbook(fh, read_only=True)
        candidates = [n for n in wb_temp.sheetnames if re.match(r'^KB26\d{4}$', n)]
        wb_temp.close()
        if not candidates:
            return False, "搵唔到 template"
        
        template = sorted(candidates)[-1]
        last_num = int(template[2:])
        new_name = f"KB{last_num + 1}"
        new_id = last_num + 1
        
        unit_rate = round(amount / qty, 2) if qty else amount

        fh.seek(0)
        with tempfile.TemporaryDirectory() as tmpdir:
            # 解壓
            with zipfile.ZipFile(fh) as z:
                z.extractall(tmpdir)
            
            # 搵 sheet 檔
            wb_xml_path = os.path.join(tmpdir, "xl", "workbook.xml")
            with open(wb_xml_path, encoding="utf-8") as f:
                wb_xml = f.read()
            
            rid_m = re.search(rf'name="{re.escape(template)}"[^>]*r:id="(rId\d+)"', wb_xml)
            if not rid_m:
                return False, "搵唔到 rid"
            rid = rid_m.group(1)
            
            rels_path = os.path.join(tmpdir, "xl", "_rels", "workbook.xml.rels")
            with open(rels_path, encoding="utf-8") as f:
                rels = f.read()
            
            tgt_m = re.search(rf'Id="{rid}"[^>]*Target="([^"]+)"', rels)
            if not tgt_m:
                return False, "搵唔到 target"
            sheet_path = tgt_m.group(1)
            
            # 抄 sheet
            src = os.path.join(tmpdir, "xl", sheet_path)
            new_sheet = f"sheet{new_id}.xml"
            dst = os.path.join(tmpdir, "xl", "worksheets", new_sheet)
            shutil.copy2(src, dst)
            
            # 抄 rels（如果有）
            src_rels = os.path.join(tmpdir, "xl", "worksheets", "_rels", os.path.basename(sheet_path) + ".rels")
            if os.path.exists(src_rels):
                dst_rels = os.path.join(tmpdir, "xl", "worksheets", "_rels", new_sheet + ".rels")
                os.makedirs(os.path.dirname(dst_rels), exist_ok=True)
                shutil.copy2(src_rels, dst_rels)
            
            # 更新 workbook.xml
            new_entry = f'<sheet name="{new_name}" sheetId="{new_id}" r:id="rId{new_id}"/>'
            wb_xml = wb_xml.replace("</sheets>", new_entry + "</sheets>")
            with open(wb_xml_path, "w", encoding="utf-8") as f:
                f.write(wb_xml)
            
            # 更新 rels
            new_rel = f'<Relationship Id="rId{new_id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/{new_sheet}"/>'
            rels = rels.replace("</Relationships>", new_rel + "</Relationships>")
            with open(rels_path, "w", encoding="utf-8") as f:
                f.write(rels)
            
            # 打包
            out_path = os.path.join(tmpdir, "out.xlsx")
            with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for root, _, files in os.walk(tmpdir):
                    for fn in files:
                        if fn == "out.xlsx":
                            continue
                        full = os.path.join(root, fn)
                        zout.write(full, os.path.relpath(full, tmpdir))
            
            # 填資料
            wb = openpyxl.load_workbook(out_path)
            if new_name not in wb.sheetnames:
                return False, "新 sheet 加入失敗"
            tgt = wb[new_name]
            
            for row in tgt.iter_rows(min_row=1, max_row=12, max_col=6):
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        v = cell.value
                        if "Invoice No" in v:
                            tgt.cell(cell.row, cell.column + 1).value = new_name
                        elif "Order No" in v:
                            tgt.cell(cell.row, cell.column + 1).value = os_no
                        elif "Quotation" in v:
                            tgt.cell(cell.row, cell.column + 1).value = qc_no
                        elif "Project" in v:
                            tgt.cell(cell.row, cell.column + 1).value = f"{os_no} - {site}"
            
            try:
                tgt["B19"] = work
                tgt["D19"] = qty
                tgt["F19"] = unit_rate
                tgt["G19"] = amount
            except Exception:
                pass
            
            if "Summary" in wb.sheetnames:
                idx = wb.sheetnames.index("Summary")
                if idx < len(wb.sheetnames) - 1:
                    wb.move_sheet("Summary", offset=len(wb.sheetnames) - 1 - idx)
            
            wb.save(out_path)
            
            media = MediaFileUpload(out_path, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            service.files().update(fileId=KB_FILE_ID, media_body=media).execute()
            
            return True, new_name
            
    except Exception as e:
        return False, f"錯誤：{e}"

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
        st.write(f"顯示前 50 行（共 {len(df)} 行）")
        st.dataframe(df.head(50), width="stretch", height=500)
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
            st.dataframe(result.head(50), width="stretch", height=400)
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
                lambda col: col.str.contains(os_no, case=False, na=False, regex=False)
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
                    final_work = work or f"Work for {os_no}"
                    final_site = site or os_no
                    with st.spinner("正在建立發票..."):
                        success, result = create_invoice_sheet(os_no, final_site, final_work, amount, qc_no)
                    if success:
                        st.success(f"已建立：{result}")
                        st.info("請到 Google Drive 打開 KB260001.xlsx 檢查")
                    else:
                        st.error(f"失敗：{result}")
                else:
                    st.warning("搵唔到金額")

st.caption("KIB Dashboard • 安全版")
