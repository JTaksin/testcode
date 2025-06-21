import streamlit as st
import pandas as pd

# ตั้งค่าหน้าของ Streamlit
st.set_page_config(page_title="Excel Multi-Filter App", layout="wide")
st.title("📊 Excel Multi-Filter Web App")

# อัปโหลดไฟล์ Excel
uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xls หรือ .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        # อ่าน Excel โดยใช้แถวที่ 4 เป็นหัวตาราง
        df = pd.read_excel(uploaded_file, header=3)

        # เงื่อนไข: ไม่เอาค่าเหล่านี้ใน column 9
        exclude_status = ["Complete", "Incomplete", "MIS Complete", "MIS Incomplete"]

        # เงื่อนไข: จังหวัดที่ต้องการใน column 17
        selected_provinces = [
        "ชัยนาท", "นครสวรรค์", "ตาก", "เชียงใหม่", "ลำปาง", "เชียงราย", "กำแพงเพชร", "พิษณุโลก",
        "น่าน", "ลำพูน", "พิจิตร", "อุตรดิตถ์", "สุโขทัย", "เพชรบูรณ์", "พะเยา", "แม่ฮ่องสอน",
        "แพร่", "อุทัยธานี"
        ]

# กรอง: column 9 (index 8) ไม่อยู่ใน exclude_status
# และ column 17 (index 16) ต้องอยู่ใน selected_provinces
df = df[~df.iloc[:, 8].isin(exclude_status) & df.iloc[:, 16].isin(selected_provinces)]


        
        st.success("✅ อัปโหลดไฟล์เรียบร้อยแล้ว!")
        st.subheader("ข้อมูลทั้งหมด:")
        st.dataframe(df, use_container_width=True)

        st.subheader("🔎 ตั้งค่าการกรองหลายคอลัมน์")

        filtered_df = df.copy()

        # เลือกเฉพาะ column 4, 7, 9, 16 โดยอิงจาก index
        filter_columns = [df.columns[i] for i in [3, 6, 8, 16]]

        for column in filter_columns:
            with st.expander(f"กรองคอลัมน์: {column}"):
                if pd.api.types.is_numeric_dtype(df[column]):
                    min_val = float(df[column].min())
                    max_val = float(df[column].max())
                    selected_range = st.slider(f"{column} - เลือกช่วงตัวเลข", min_val, max_val, (min_val, max_val))
                    filtered_df = filtered_df[filtered_df[column].between(*selected_range)]

                elif pd.api.types.is_datetime64_any_dtype(df[column]):
                    date_range = st.date_input(f"{column} - เลือกช่วงวันที่", [])
                    if len(date_range) == 2:
                        start_date, end_date = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
                        filtered_df = filtered_df[(df[column] >= start_date) & (df[column] <= end_date)]

                else:
                    unique_vals = df[column].dropna().astype(str).unique()
                    selected_vals = st.multiselect(f"{column} - เลือกค่าที่ต้องการ", sorted(unique_vals))
                    if selected_vals:
                        filtered_df = filtered_df[df[column].astype(str).isin(selected_vals)]

        st.subheader("📋 ข้อมูลหลังกรองทั้งหมด:")
        st.dataframe(filtered_df, use_container_width=True)

        # Optional: Download
        csv = filtered_df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            label="📥 ดาวน์โหลดข้อมูลกรองเป็น CSV",
            data=csv,
            file_name="filtered_multi_data.csv",
            mime="text/csv"
        )

    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
