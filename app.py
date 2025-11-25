import streamlit as st
import pandas as pd
import io

# ==========================================
# 1. ตั้งค่าหน้าเว็บ (ต้องอยู่บรรทัดแรกสุด)
# ==========================================
st.set_page_config(page_title="Excel Utility App", layout="wide")

# ==========================================
# 2. สร้างเมนูเลือกฟังก์ชัน (Sidebar)
# ==========================================
st.sidebar.title("เมนูเลือกเครื่องมือ")
app_mode = st.sidebar.radio(
    "เลือกฟังก์ชันที่ต้องการใช้งาน:",
    ["📊 กรองข้อมูล (Multi-Filter)", "🛠️ แก้ภาษาต่างด้าว (Fix Encoding)"]
)

# ==========================================
# ฟังก์ชันช่วย: แก้ภาษาต่างด้าว
# ==========================================
def fix_thai_encoding(text):
    if not isinstance(text, str):
        return text
    try:
        # แปลง Latin-1 กลับเป็น Bytes แล้วถอดรหัสด้วย TIS-620 (cp874)
        return text.encode('latin-1').decode('cp874')
    except (UnicodeEncodeError, UnicodeDecodeError):
        return text

# ==========================================
# ส่วนที่ 1: แอปกรองข้อมูล (โค้ดเดิมของคุณ)
# ==========================================
if app_mode == "📊 กรองข้อมูล (Multi-Filter)":
    st.title("📊 Excel Multi-Filter Web App")

    # 🔗 เพิ่มลิงก์
    st.markdown(
        '<a href="http://192.168.0.50/modules/mis/report_plandate.php" target="_blank" style="text-decoration: none;">🔗 เปิดรายงาน Plan Date แบบออนไลน์</a>',
        unsafe_allow_html=True
    )

    # อัปโหลดไฟล์ Excel
    uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xls หรือ .xlsx) เพื่อกรอง", type=["xls", "xlsx"], key="filter_uploader")

    if uploaded_file:
        try:
            # ✅ โหลดไฟล์ (ข้าม 3 บรรทัดแรกตามโค้ดเดิม)
            df = pd.read_excel(uploaded_file, header=3)

            # ✅ กรองข้อมูล (ตามเงื่อนไข Onsite / จังหวัด)
            exclude_status = ["Complete", "Incomplete", "MIS Complete", "MIS Incomplete"]
            selected_provinces = [
                "ชัยนาท", "นครสวรรค์", "ตาก", "เชียงใหม่", "ลำปาง", "เชียงราย", "กำแพงเพชร", "พิษณุโลก",
                "น่าน", "ลำพูน", "พิจิตร", "อุตรดิตถ์", "สุโขทัย", "เพชรบูรณ์", "พะเยา", "แม่ฮ่องสอน",
                "แพร่", "อุทัยธานี"
            ]

            # ตรวจสอบว่าคอลัมน์ที่ 8 และ 16 มีอยู่จริงหรือไม่เพื่อป้องกัน Error
            if df.shape[1] > 16:
                df = df[~df.iloc[:, 8].isin(exclude_status) & df.iloc[:, 16].isin(selected_provinces)]
            else:
                st.warning("⚠️ ไฟล์ที่อัปโหลดมีจำนวนคอลัมน์ไม่เพียงพอสำหรับการกรองอัตโนมัติ (ข้ามการกรองจังหวัด/สถานะ)")

            # ✅ ส่วนกรองหลายคอลัมน์
            st.subheader("🔎 ตั้งค่าการกรองเพิ่มเติม")
            filtered_df = df.copy()

            # รายชื่อคอลัมน์ที่ต้องการกรอง (ต้องมีชื่อตรงกับในไฟล์จริง)
            target_columns = ['Project', 'Plan Date', 'Status', 'Province']
            
            # ตรวจสอบว่าคอลัมน์เหล่านี้มีอยู่ในไฟล์จริงไหม
            available_columns = [col for col in target_columns if col in df.columns]

            if not available_columns:
                st.info(f"ไม่พบคอลัมน์ {target_columns} ในไฟล์นี้ แสดงตัวกรองเฉพาะคอลัมน์ที่มีอยู่...")
                available_columns = df.columns.tolist()

            for column in available_columns:
                # สร้าง Container ให้ดูสะอาดตา
                with st.expander(f"ตัวกรอง: {column}", expanded=False):
                    if pd.api.types.is_numeric_dtype(df[column]):
                        min_val = float(df[column].min())
                        max_val = float(df[column].max())
                        # ตรวจสอบว่าค่าไม่เป็น NaN
                        if pd.notna(min_val) and pd.notna(max_val):
                            selected_range = st.slider(f"ช่วงตัวเลข ({column})", min_val, max_val, (min_val, max_val))
                            filtered_df = filtered_df[filtered_df[column].between(*selected_range)]
                    
                    elif pd.api.types.is_datetime64_any_dtype(df[column]):
                        date_input = st.date_input(f"ช่วงวันที่ ({column})", [])
                        if len(date_input) == 2:
                            start_date, end_date = pd.to_datetime(date_input[0]), pd.to_datetime(date_input[1])
                            filtered_df = filtered_df[(df[column] >= start_date) & (df[column] <= end_date)]
                    
                    else:
                        # สำหรับข้อความ (Text)
                        unique_vals = df[column].dropna().astype(str).unique()
                        selected_vals = st.multiselect(f"เลือกค่า ({column})", sorted(unique_vals))
                        if selected_vals:
                            filtered_df = filtered_df[df[column].astype(str).isin(selected_vals)]

            # ✅ แสดงผล
            st.subheader(f"📋 ข้อมูลที่กรองแล้ว ({len(filtered_df)} รายการ)")
            st.dataframe(filtered_df, use_container_width=True)

            # 🔽 ดาวน์โหลดเป็น Excel
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                filtered_df.to_excel(writer, index=False, sheet_name='FilteredData')
            
            st.download_button(
                label="📥 ดาวน์โหลดผลลัพธ์ (Excel)",
                data=output.getvalue(),
                file_name="filtered_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาดในการประมวลผล: {e}")

# ==========================================
# ส่วนที่ 2: แอปแก้ภาษาต่างด้าว (โค้ดใหม่)
# ==========================================
elif app_mode == "🛠️ แก้ภาษาต่างด้าว (Fix Encoding)":
    st.title("🛠️ โปรแกรมแก้ภาษาต่างด้าว")
    st.info("ใช้สำหรับไฟล์ Excel/CSV ที่เปิดแล้วเป็นภาษาอ่านไม่ออก (เช่น `§Ò¹·ÕèÊè§`)")

    uploaded_file_fix = st.file_uploader("อัปโหลดไฟล์ที่ภาษาเพี้ยน (.csv, .xls, .xlsx)", type=['csv', 'xlsx', 'xls'], key="fix_uploader")

    if uploaded_file_fix is not None:
        try:
            df_fix = None
            
            # ตรวจสอบประเภทไฟล์
            if uploaded_file_fix.name.endswith('.csv'):
                # อ่านแบบ Latin-1 เพื่อรับค่า Mojibake เข้ามา
                df_fix = pd.read_csv(uploaded_file_fix, encoding='latin-1')
            else:
                df_fix = pd.read_excel(uploaded_file_fix)

            st.write("### 1. ตัวอย่างข้อมูลก่อนแก้")
            st.dataframe(df_fix.head())

            # แปลงภาษาทั้งตาราง
            df_fixed = df_fix.applymap(fix_thai_encoding)

            st.write("### 2. ตัวอย่างข้อมูลหลังแก้ (ภาษาไทย)")
            st.dataframe(df_fixed.head())

            # แปลงกลับเป็นไฟล์ Excel
            output_fix = io.BytesIO()
            with pd.ExcelWriter(output_fix, engine='xlsxwriter') as writer:
                df_fixed.to_excel(writer, index=False, sheet_name='Corrected_Data')
            
            st.success("✅ แปลงภาษาเรียบร้อย!")
            
            st.download_button(
                label="📥 ดาวน์โหลดไฟล์ที่แก้แล้ว",
                data=output_fix.getvalue(),
                file_name=f"fixed_{uploaded_file_fix.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")
            st.warning("หากเป็นไฟล์ CSV ลอง Save As เป็น UTF-8 ก่อนนำเข้า หรือตรวจสอบรูปแบบไฟล์ต้นฉบับ")
