import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.title("🔔 DC Disconnect Virtual Alarm with Detailed Time Matching")

# Upload files
uploaded_file_1 = st.file_uploader("📂 Upload DC Disconnect Virtual Alarm Excel (with Start/End Time)", type=["xlsx"])
uploaded_file_2 = st.file_uploader("📂 Upload Node Alarms Excel (with Start/End Time)", type=["xlsx"])

if uploaded_file_1 and uploaded_file_2:
    # Read Excel files
    df1 = pd.read_excel(uploaded_file_1)
    df2 = pd.read_excel(uploaded_file_2)

    # Clean column names
    df1.columns = df1.columns.str.strip()
    df2.columns = df2.columns.str.strip()

    st.write("📊 DC Disconnect Columns:", df1.columns.tolist())
    st.write("📊 Node Alarms Columns:", df2.columns.tolist())

    # Validate necessary columns
    if not {'Site', 'Start Time', 'End Time'}.issubset(df1.columns):
        st.error("⚠️ DC Disconnect file must have 'Site', 'Start Time', and 'End Time'.")
    elif not {'Site', 'Node', 'Start Time', 'End Time'}.issubset(df2.columns):
        st.error("⚠️ Node Alarms file must have 'Site', 'Node', 'Start Time', and 'End Time'.")
    else:
        # Convert times
        df1['Start Time'] = pd.to_datetime(df1['Start Time'])
        df1['End Time'] = pd.to_datetime(df1['End Time'])
        df2['Start Time'] = pd.to_datetime(df2['Start Time'])
        df2['End Time'] = pd.to_datetime(df2['End Time'])

        # Create a detailed match table
        matches = []
        for idx1, row1 in df1.iterrows():
            site1, start1, end1 = row1['Site'], row1['Start Time'], row1['End Time']
            matching_alarms = df2[
                (df2['Site'] == site1) &
                (
                    ((df2['Start Time'] >= start1) & (df2['Start Time'] <= end1)) |
                    ((df2['End Time'] >= start1) & (df2['End Time'] <= end1)) |
                    ((df2['Start Time'] <= start1) & (df2['End Time'] >= end1))
                )
            ]
            if not matching_alarms.empty:
                for _, row2 in matching_alarms.iterrows():
                    matches.append({
                        'Site': site1,
                        'DC Start': start1,
                        'DC End': end1,
                        'Matched Alarm': row2['Node'],
                        'Alarm Start': row2['Start Time'],
                        'Alarm End': row2['End Time']
                    })
            else:
                matches.append({
                    'Site': site1,
                    'DC Start': start1,
                    'DC End': end1,
                    'Matched Alarm': '',
                    'Alarm Start': '',
                    'Alarm End': ''
                })

        match_df = pd.DataFrame(matches)

        st.subheader("✅ Detailed Time-Frame Matching Table")
        st.dataframe(match_df)

        # Optionally allow download
        towrite = BytesIO()
        match_df.to_excel(towrite, index=False, sheet_name='Matches')
        towrite.seek(0)
        st.download_button("⬇️ Download Detailed Match Excel", towrite, file_name="Detailed_Alarm_Matches.xlsx")

else:
    st.info("Please upload both Excel files.")
