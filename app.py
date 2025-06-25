import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="Excel Filter App", layout="wide")
st.title("📊 ข้อมูล - งบประมาณ ปี 2561-2568")

# ---------- โหลดข้อมูลจาก Google Sheet ----------
sheet_id = "1Pjf0A4-M9NTxkK8Cj0AMCMiLmazfQNqq7zRb3Lnw2G8"
sheet_name = "Sheet1"  # เปลี่ยนชื่อชีตตรงนี้ถ้าชื่อไม่ใช่ Sheet1
csv_url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv&sheet={sheet_name}"

try:
    df = pd.read_csv(csv_url)
except Exception as e:
    st.error(f"❌ โหลดข้อมูลจาก Google Sheet ไม่สำเร็จ: {e}")
    st.stop()

# ---------- ตรวจสอบคอลัมน์ที่จำเป็น ----------
required_columns = ["ลำดับ", "โครงการ", "รูปแบบงบประมาณ", "ปีงบประมาณ", "หน่วยงาน",
                    "สถานที่", "หมู่ที่", "ตำบล", "อำเภอ", "จังหวัด"]
if not all(col in df.columns for col in required_columns):
    st.error("⚠️ ไฟล์ Google Sheet ไม่มีคอลัมน์ที่ต้องการ หรือชื่อคอลัมน์ไม่ถูกต้อง")
    st.stop()

# ---------- ฟังก์ชันช่วย ----------
def extract_number(s):
    match = re.search(r"\d+", str(s))
    return int(match.group()) if match else float('inf')

def get_options(dataframe, column_name):
    options = dataframe[column_name].dropna().unique().tolist()
    if column_name == "ปีงบประมาณ":
        options = sorted([str(y) for y in options])
    elif column_name == "หน่วยงาน":
        options = sorted(options, key=extract_number)
    else:
        options.sort()
    return ["ทั้งหมด"] + options

# ---------- ค่าที่เลือกเริ่มต้น ----------
selected_budget = "ทั้งหมด"
selected_year = "ทั้งหมด"
selected_project = "ทั้งหมด"
selected_departments = ["ทั้งหมด"]

# ---------- ส่วนของตัวกรอง ----------
st.markdown("### 🔍 เลือกตัวกรองข้อมูล")
col1, col2 = st.columns(2)
col3, col4 = st.columns(2)

filtered_df_for_options = df.copy()

with col1:
    budget_options = get_options(filtered_df_for_options, "รูปแบบงบประมาณ")
    selected_budget = st.selectbox("💰 รูปแบบงบประมาณ", budget_options, key="budget_select")
    if selected_budget != "ทั้งหมด":
        filtered_df_for_options = filtered_df_for_options[filtered_df_for_options["รูปแบบงบประมาณ"] == selected_budget]

with col2:
    year_options = get_options(filtered_df_for_options, "ปีงบประมาณ")
    selected_year = st.selectbox("📅 ปีงบประมาณ", year_options, key="year_select")
    if selected_year != "ทั้งหมด":
        filtered_df_for_options = filtered_df_for_options[filtered_df_for_options["ปีงบประมาณ"].astype(str) == selected_year]

with col3:
    project_options = get_options(filtered_df_for_options, "โครงการ")
    selected_project = st.selectbox("📌 โครงการ", project_options, key="project_select")
    if selected_project != "ทั้งหมด":
        filtered_df_for_options = filtered_df_for_options[filtered_df_for_options["โครงการ"] == selected_project]

with col4:
    department_options = get_options(filtered_df_for_options, "หน่วยงาน")
    current_selected_departments = st.session_state.get("dept_select", ["ทั้งหมด"])
    valid_defaults = [d for d in current_selected_departments if d in department_options]
    if not valid_defaults:
        valid_defaults = ["ทั้งหมด"]
    selected_departments = st.multiselect("🏢 หน่วยงาน", department_options, default=valid_defaults, key="dept_select")
    if "ทั้งหมด" not in selected_departments:
        filtered_df_for_options = filtered_df_for_options[filtered_df_for_options["หน่วยงาน"].isin(selected_departments)]

# ---------- กรองข้อมูลจริง ----------
filtered_df = df.copy()
if selected_budget != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["รูปแบบงบประมาณ"] == selected_budget]
if selected_year != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["ปีงบประมาณ"].astype(str) == selected_year]
if selected_project != "ทั้งหมด":
    filtered_df = filtered_df[filtered_df["โครงการ"] == selected_project]
if "ทั้งหมด" not in selected_departments:
    filtered_df = filtered_df[filtered_df["หน่วยงาน"].isin(selected_departments)]

# ---------- แสดงผลลัพธ์ ----------
if not filtered_df.empty:
    st.markdown(
        f"""
        <div style='
            font-size: 24px;
            color: #084298;
            background-color: #cfe2ff;
            padding: 12px;
            border-radius: 8px;
            border-left: 6px solid #084298;
            margin-bottom: 16px;
        '>
            📈 พบข้อมูลทั้งหมด {len(filtered_df)} รายการ
        </div>
        """,
        unsafe_allow_html=True
    )
else:
    st.markdown(
        """
        <div style='
            font-size: 24px;
            color: #8a6d3b;
            background-color: #fcf8e3;
            padding: 12px;
            border-radius: 8px;
            border-left: 6px solid #8a6d3b;
            margin-bottom: 16px;
        '>
            ⚠️ ไม่พบข้อมูลที่ตรงกับเงื่อนไขที่เลือก
        </div>
        """,
        unsafe_allow_html=True
    )

# ---------- ตารางข้อมูล ----------
st.markdown("### 📄 ตารางข้อมูล")
st.dataframe(filtered_df, use_container_width=True)

# ---------- ดาวน์โหลดเป็น Excel ----------
def to_excel_bytes(df_to_export):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_to_export.to_excel(writer, index=False)
    return output.getvalue()

if not filtered_df.empty:
    st.download_button(
        label="📥 ดาวน์โหลดข้อมูลที่กรองเป็น Excel",
        data=to_excel_bytes(filtered_df),
        file_name="filtered_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
