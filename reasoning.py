import streamlit as st
import pandas as pd
from datetime import timedelta

st.title("📊 3-Day Hourly Alarm Breakdown")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load Excel with correct header row (row 3 = header=2)
        df = pd.read_excel(uploaded_file, header=2)

        required_columns = ['Start Time', 'Node']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns: {required_columns}")
        else:
            # Convert Start Time to datetime
            df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
            df = df.dropna(subset=['Start Time'])

            # Extract date and hour
            df['Date'] = df['Start Time'].dt.date
            df['Hour'] = df['Start Time'].dt.hour

            # Filter last 3 days
            max_date = df['Date'].max()
            last_3_days = [max_date - timedelta(days=i) for i in range(3)]
            df_filtered = df[df['Date'].isin(last_3_days)]

            # Group by Date, Hour, Node
            grouped = df_filtered.groupby(['Date', 'Hour', 'Node']).size().reset_index(name='Alarm Count')

            # Pivot to show Node (alarm type) as columns
            pivot = grouped.pivot_table(index=['Date', 'Hour'], columns='Node', values='Alarm Count', fill_value=0)

            # Display result
            st.subheader("📅 Hourly Alarm Breakdown by Type (Last 3 Days)")
            st.dataframe(pivot)

            # Save to Excel
            output_excel = pd.ExcelWriter("alarm_breakdown.xlsx", engine='openpyxl')
            pivot.to_excel(output_excel)
            output_excel.close()

            with open("alarm_breakdown.xlsx", "rb") as f:
                st.download_button("⬇ Download Breakdown Excel", f, file_name="3_day_alarm_breakdown.xlsx")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to begin.")
