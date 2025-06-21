import streamlit as st
import pandas as pd

# ตั้งค่าหน้าของ Streamlit
st.set_page_config(page_title="Excel Multi-Filter App", layout="wide")
st.title("📊 Excel Multi-Filter Web App")

# อัปโหลดไฟล์ Excel
uploaded_file = st.file_uploader("อัปโหลดไฟล์ Excel (.xls หรือ .xlsx)", type=["xls", "xlsx"])

if uploaded_file:
    try:
        # อ่านไฟล์ Excel ทั้ง .xls และ .xlsx โดย pandas จะเลือก engine อัตโนมัติ
        df = pd.read_excel(uploaded_file, header=3)
        
        st.success("✅ อัปโหลดไฟล์เรียบร้อยแล้ว!")
        st.subheader("ข้อมูลทั้งหมด:")
        st.dataframe(df, use_container_width=True)

        st.subheader("🔎 ตั้งค่าการกรองหลายคอลัมน์")

        filter_conditions = {}
        filtered_df = df.copy()

        for column in df.columns:
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
