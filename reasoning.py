import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.title("📊 Power System Event Duration Table")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        # Required columns check
        required_columns = ['Start Time', 'End Time', 'Node']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns: {required_columns}")
        else:
            # Convert to datetime
            df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
            df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce')
            df = df.dropna(subset=['Start Time', 'End Time'])

            # Duration calculation
            df['Duration (Hours)'] = (df['End Time'] - df['Start Time']).dt.total_seconds() / 3600

            # Add formatted Start Date and Hour
            try:
                df['Start Date'] = df['Start Time'].dt.strftime('%-d-%B, %Y')  # Linux/Mac
            except:
                df['Start Date'] = df['Start Time'].dt.strftime('%#d-%B, %Y')  # Windows fallback

            df['Hour'] = df['Start Time'].dt.hour

            # Display
            st.subheader("📋 All Event Records")
            st.dataframe(
                df[['Start Date', 'Hour', 'Start Time', 'End Time', 'Node', 'Duration (Hours)']].sort_values('Start Time'),
                use_container_width=True,
                hide_index=True
            )

            # Download
            output = BytesIO()
            df.to_excel(output, index=False, sheet_name='Event Data')
            output.seek(0)
            st.download_button(
                label="📥 Download Processed Excel File",
                data=output,
                file_name="event_data_with_hour.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to begin.")
