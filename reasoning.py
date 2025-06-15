import streamlit as st
import pandas as pd
from datetime import timedelta

st.title("📊 3-Day Hourly Alarm Trend Generator")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load the Excel file
        df = pd.read_excel(uploaded_file)

        # Ensure required columns exist
        required_columns = ['Start Time']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required column(s): {required_columns}")
        else:
            # Convert 'Start Time' to datetime
            df['Start Time'] = pd.to_datetime(df['Start Time'], format="%m/%d/%Y %I:%M:%S %p", errors='coerce')
            df = df.dropna(subset=['Start Time'])

            # Extract date and hour
            df['Date'] = df['Start Time'].dt.date
            df['Hour'] = df['Start Time'].dt.hour

            # Determine the latest date in data
            max_date = df['Date'].max()
            last_3_days = [max_date - timedelta(days=i) for i in range(3)]

            # Filter for the last 3 days
            df_filtered = df[df['Date'].isin(last_3_days)]

            # Group and pivot
            alarm_counts = df_filtered.groupby(['Date', 'Hour']).size().reset_index(name='Alarm Count')
            pivot = alarm_counts.pivot(index='Date', columns='Hour', values='Alarm Count').fillna(0).astype(int)
            pivot = pivot.reindex(columns=range(24), fill_value=0)  # Ensure 0-23 hours

            # Display result
            st.subheader("📅 Hourly Alarm Count (Last 3 Days)")
            st.dataframe(pivot)

            # Download as Excel
            output_excel = pd.ExcelWriter("alarm_trend.xlsx", engine='openpyxl')
            pivot.to_excel(output_excel, index=True)
            output_excel.close()

            with open("alarm_trend.xlsx", "rb") as f:
                st.download_button("⬇ Download Excel", f, file_name="3_day_hourly_alarm_trend.xlsx")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to begin.")
