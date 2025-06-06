import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Alignment

st.title("🔔 DC Disconnect Virtual Alarm Detailed Alarm Match Report")

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

    if not {'Site', 'Start Time', 'End Time'}.issubset(df1.columns):
        st.error("⚠️ DC Disconnect file must have 'Site', 'Start Time', and 'End Time'.")
    elif not {'Site', 'Node', 'Start Time', 'End Time'}.issubset(df2.columns):
        st.error("⚠️ Node Alarms file must have 'Site', 'Node', 'Start Time', and 'End Time'.")
    else:
        df1['Start Time'] = pd.to_datetime(df1['Start Time'], errors='coerce')
        df1['End Time'] = pd.to_datetime(df1['End Time'], errors='coerce')
        df2['Start Time'] = pd.to_datetime(df2['Start Time'], errors='coerce')
        df2['End Time'] = pd.to_datetime(df2['End Time'], errors='coerce')

        alarm_types = sorted(df2['Node'].dropna().unique())

        result_df = df1.copy()
        for alarm in alarm_types:
            result_df[alarm] = ''

        for idx1, row1 in df1.iterrows():
            site1, start1, end1 = row1['Site'], row1['Start Time'], row1['End Time']
            if pd.isna(site1) or pd.isna(start1) or pd.isna(end1):
                continue
            matching_alarms = df2[
                (df2['Site'] == site1) &
                (
                    ((df2['Start Time'] >= start1) & (df2['Start Time'] <= end1)) |
                    ((df2['End Time'] >= start1) & (df2['End Time'] <= end1)) |
                    ((df2['Start Time'] <= start1) & (df2['End Time'] >= end1))
                )
            ]
            for alarm in alarm_types:
                matches = matching_alarms[matching_alarms['Node'] == alarm]
                if not matches.empty:
                    details = "\n\n".join([  # Double space
                        f"{row2['Start Time'].strftime('%Y-%m-%d %H:%M') if pd.notna(row2['Start Time']) else 'Unknown'} ➡ {row2['End Time'].strftime('%H:%M') if pd.notna(row2['End Time']) else 'Unknown'}"
                        for _, row2 in matches.iterrows()
                    ])
                    result_df.at[idx1, alarm] = details

        st.subheader("✅ Detailed Match Table with Alarm Times (Green for Matches)")
        
        def highlight_matches(val):
            return 'background-color: lightgreen' if val != '' else ''
        
        styled_df = result_df.style.applymap(highlight_matches, subset=alarm_types)
        st.dataframe(styled_df, use_container_width=True)

        # Prepare Excel with formatting
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = "Detailed_Matches"

        # Write header
        for col_num, col_name in enumerate(result_df.columns, 1):
            ws.cell(row=1, column=col_num, value=col_name)

        # Write data with formatting
        green_fill = PatternFill(start_color='90EE90', end_color='90EE90', fill_type='solid')
        center_align = Alignment(horizontal="left", vertical="top", wrap_text=True)

        for row_num, row_data in enumerate(result_df.itertuples(index=False), 2):
            for col_num, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_num, column=col_num, value=value)
                cell.alignment = center_align
                if col_num > len(df1.columns) and value != '':
                    cell.fill = green_fill

        wb.save(output)
        output.seek(0)

        st.download_button("⬇️ Download Excel with Green Highlights", output, file_name="Detailed_Alarm_Report.xlsx")

else:
    st.info("Please upload both Excel files.")
