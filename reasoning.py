import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

st.title("🔔 DC Disconnect Virtual Alarm Analyzer with Time Matching & Hourly Trend")

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

        # Custom alarm ordering
        custom_order = [
            "DCDB-01 Primary Disconnect",
            "DCDB-01 Critical Disconnect",
            "Battery Low",
            "Battery Critical",
            "Mains Fail",
            "Rectifier Module Fault",
            "MDB Fault",
            "PG Run",
            "Vibration",
            "Motion"
        ]
        all_alarms = sorted(df2['Node'].unique())
        remaining_alarms = [a for a in all_alarms if a not in custom_order]
        ordered_alarms = custom_order + remaining_alarms

        # Prepare result dataframe
        result_df = df1.copy()
        for alarm in ordered_alarms:
            result_df[alarm] = ''

        # Match alarms by site and time overlap
        for idx1, row1 in df1.iterrows():
            site1, start1, end1 = row1['Site'], row1['Start Time'], row1['End Time']
            matches = df2[
                (df2['Site'] == site1) &
                (
                    ((df2['Start Time'] >= start1) & (df2['Start Time'] <= end1)) |
                    ((df2['End Time'] >= start1) & (df2['End Time'] <= end1)) |
                    ((df2['Start Time'] <= start1) & (df2['End Time'] >= end1))
                )
            ]
            for alarm in ordered_alarms:
                if alarm in matches['Node'].values:
                    result_df.at[idx1, alarm] = '✓'

        # Highlight style
        def highlight(val):
            return 'background-color: lightgreen' if val == '✓' else ''

        st.subheader("✅ Matched Alarms with Time Overlap")
        st.dataframe(result_df.style.applymap(highlight, subset=ordered_alarms))

        # Excel export
        towrite = BytesIO()
        wb = Workbook()
        ws = wb.active
        headers = list(df1.columns) + ordered_alarms
        ws.append(headers)

        green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
        tick_font = Font(name='Calibri', size=11, bold=True)
        align_center = Alignment(horizontal='center', vertical='center')

        for idx, row in result_df.iterrows():
            row_data = [row.get(col, '') for col in headers]
            ws.append(row_data)
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=idx+2, column=col_idx)
                if col_idx > len(df1.columns) and value == '✓':
                    cell.fill = green_fill
                    cell.font = tick_font
                    cell.alignment = align_center

        for col in ws.columns:
            max_len = max(len(str(cell.value)) if cell.value else 0 for cell in col) + 2
            ws.column_dimensions[col[0].column_letter].width = max_len

        wb.save(towrite)
        towrite.seek(0)
        st.download_button("⬇️ Download Highlighted Excel", towrite, file_name="Matched_Alarms.xlsx")

        # Hourly trend chart
        st.subheader("📈 Hourly Alarm Trend")
        df1['Hour'] = df1['Start Time'].dt.floor('H')
        df2['Hour'] = df2['Start Time'].dt.floor('H')
        trend_dc = df1.groupby('Hour').size().reset_index(name='DC Disconnect Count')
        trend_node = df2.groupby('Hour').size().reset_index(name='Node Alarm Count')
        trend = pd.merge(trend_dc, trend_node, on='Hour', how='outer').fillna(0)
        st.line_chart(trend.set_index('Hour'))

else:
    st.info("Please upload both Excel files.")
