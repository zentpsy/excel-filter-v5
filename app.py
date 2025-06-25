import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
from io import BytesIO
import re

st.set_page_config(page_title="Excel Filter App - Google Sheets", layout="wide")
st.title("📊 ข้อมูล - งบประมาณ ปี 2561-2568 จาก Google Sheets")

# --- เชื่อม Google Sheets ด้วย Service Account จาก Secrets ---
creds_info = st.secrets["gcp_service_account"]  # ต้องมีใน secrets.toml
credentials = Credentials.from_service_account_info(creds_info, scopes=["https://www.googleapis.com/auth/spreadsheets"])

gc = gspread.authorize(credentials)

# --- เปิด Google Sheet และ Worksheet ---
SPREADSHEET_ID = "ใส่ Spreadsheet ID ตรงนี้"  # เช่น https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit...
WORKSHEET_NAME = "Sheet1"  # ชื่อชีตที่ต้องการดึงข้อมูล

sheet = gc.open_by_key(SPREADSHEET_ID).worksheet(WORKSHEET_NAME)

# ดึงข้อมูลมาเป็น dict แล้วแปลงเป็น DataFrame
data = sheet.get_all_records()
df = pd.DataFrame(data)

# ตรวจสอบคอลัมน์
required_columns = ["ลำดับ", "โครงการ", "รูปแบบงบประมาณ", "ปีงบประมาณ", "หน่วยงาน",
                    "สถานที่", "หมู่ที่", "ตำบล", "อำเภอ", "จังหวัด"]
if not all(col in df.columns for col in required_columns):
    st.error("ไฟล์ Google Sheets ไม่มีคอลัมน์ที่ต้องการ หรือชื่อคอลัมน์ไม่ถูกต้อง")
    st.stop()

# ฟังก์ชันช่วย sort ปี และหน่วยงาน
def extract_number(s):
    match = re.search(r"\d+", str(s))
    return int(match.group()) if match else float('inf')

# ฟังก์ชันดึง options สำหรับ dropdown + เพิ่ม "ทั้งหมด"
def get_options(df, col_name):
    opts = df[col_name].dropna().unique().tolist()
    if col_name == "ปีงบประมาณ":
        opts = sorted([str(x) for x in opts])
    elif col_name == "หน่วยงาน":
        opts = sorted(opts, key=extract_number)
    else:
        opts.sort()
    return ["ทั้งหมด"] + opts

# สร้างตัวแปรกรองสำหรับ dropdown option ให้สัมพันธ์กัน
filtered_for_options = df.copy()

# --- สร้าง Dropdown ตัวกรอง ---

col1, col2 = st.columns(2)
col3, col4 = st.columns(2)

with col1:
    budget_options = get_options(filtered_for_options, "รูปแบบงบประมาณ")
    selected_budget = st.selectbox("💰 รูปแบบงบประมาณ", budget_options, key="budget_select")
    if selected_budget != "ทั้งหมด":
        filtered_for_options = filtered_for_options[filtered_for_options["รูปแบบงบประมาณ"] == selected_budget]

with col2:
    year_options = get_options(filtered_for_options, "ปีงบประมาณ")
    selected_year = st.selectbox("📅 ปีงบประมาณ", year_options, key="year_select")
    if selected_year != "ทั้งหมด":
        filtered_for_options = filtered_for_options[filtered_for_options["ปีงบประมาณ"].astype(str) == selected_year]

with col3:
    project_options = get_options(filtered_for_options, "โครงการ")
    selected_project = st.selectbox("📌 โครงการ", project_options, key="project_select")
    if selected_project != "ทั้งหมด":
        filtered_for_options = filtered_for_options[filtered_for_options["โครงการ"] == selected_project]

with col4:
    department_options = get_options(filtered_for_options, "หน่วยงาน")
    # ตั้ง default ให้ multiselect
    default_departments = st.session_state.get("dept_select", ["ทั้งหมด"])
    valid_defaults = [d for d in default_departments if d in department_options]
    if not valid_defaults:
        valid_defaults = ["ทั้งหมด"]
    selected_departments = st.multiselect("🏢 หน่วยงาน", department_options, default=valid_defaults, key="dept_select")
    if "ทั้งหมด" not in selected_departments:
        filtered_for_options = filtered_for_options[filtered_for_options["หน่วยงาน"].isin(selected_departments)]

# --- กรองข้อมูลจริงตามตัวเลือก ---

filtered_df = df.copy()

if selected_budget != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["รูปแบบงบประมาณ"] == selected_budget]

if selected_year != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["ปีงบประมาณ"].astype(str) == selected_year]

if selected_project != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["โครงการ"] == selected_project]

if "ทั้งหมด" not in selected_departments:
    filtered_df = filtered_df[filtered_df["หน่วยงาน"].isin(selected_departments)]

# --- แสดงผลจำนวนข้อมูล พร้อมตกแต่งข้อความสีฟ้า ---
if not filtered_df.empty:
    st.markdown(
        f"<div style='font-size:24px; color:#3178c6; background-color:#d0e7ff; padding:10px; border-radius:6px;'>"
        f"📈 พบข้อมูลทั้งหมด {len(filtered_df)} รายการ</div>",
        unsafe_allow_html=True
    )
else:
    st.warning("⚠️ ไม่พบข้อมูลที่ตรงกับเงื่อนไขที่เลือก")

# --- แสดงตารางข้อมูล ---
st.markdown("### 📄 ตารางข้อมูล")
st.dataframe(filtered_df, use_container_width=True)

# --- ฟังก์ชันแปลง DataFrame เป็น Excel bytes ---
def to_excel_bytes(df_to_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_export.to_excel(writer, index=False)
    return output.getvalue()

# --- ปุ่มดาวน์โหลด Excel ---
if not filtered_df.empty:
    st.download_button(
        label="📥 ดาวน์โหลดข้อมูลที่กรองเป็น Excel",
        data=to_excel_bytes(filtered_df),
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
