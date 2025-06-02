import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

st.title("DC Disconnect Virtual Alarm Analyzer with Time Matching and Hourly Trend")

# File uploads
uploaded_file_1 = st.file_uploader("Upload DC Disconnect Virtual Alarm Excel (with Start/End Time)", type=["xlsx"])
uploaded_file_2 = st.file_uploader("Upload Node Alarms Excel (with Start/End Time)", type=["xlsx"])
uploaded_file_3 = st.file_uploader("Upload DC Disconnect Hourly Excel (Optional)", type=["xlsx"])
uploaded_file_4 = st.file_uploader("Upload Node Alarms Hourly Excel (Optional)", type=["xlsx"])

if uploaded_file_1 and uploaded_file_2:
    df1 = pd.read_excel(uploaded_file_1)
    df2 = pd.read_excel(uploaded_file_2)

    # Clean data
    df1['Site'] = df1['Site'].astype(str).str.strip()
    df2['Site'] = df2['Site'].astype(str).str.strip()
    df1['Start Time'] = pd.to_datetime(df1['Start Time'])
    df1['End Time'] = pd.to_datetime(df1['End Time'])
    df2['Start Time'] = pd.to_datetime(df2['Start Time'])
    df2['End Time'] = pd.to_datetime(df2['End Time'])
    
    # Extract unique alarms and custom order
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
    all_unique_alarms = sorted(df2['Node'].unique())
    remaining_alarms = [alarm for alarm in all_unique_alarms if alarm not in custom_order]
    ordered_alarms = custom_order + remaining_alarms
    
    # Initialize result DataFrame
    result_df = df1.copy()
    for alarm in ordered_alarms:
        result_df[alarm] = ''
    
    # Match logic with time overlap
    for idx1, row1 in df1.iterrows():
        site1, start1, end1 = row1['Site'], row1['Start Time'], row1['End Time']
        matching_rows = df2[(df2['Site'] == site1) & (
            ((df2['Start Time'] >= start1) & (df2['Start Time'] <= end1)) | 
            ((df2['End Time'] >= start1) & (df2['End Time'] <= end1)) |
            ((df2['Start Time'] <= start1) & (df2['End Time'] >= end1))  # Full overlap
        )]
        for alarm in ordered_alarms:
            if alarm in matching_rows['Node'].values:
                result_df.at[idx1, alarm] = '✓'
    
    # Streamlit display with styling
    def highlight_alarms(val):
        if val == '✓':
            return 'background-color: lightgreen'
        return ''
    
    st.subheader("Matched Alarms with Time Overlap")
    st.dataframe(result_df.style.applymap(highlight_alarms, subset=ordered_alarms))
    
    # Prepare Excel for download
    towrite = BytesIO()
    workbook = Workbook()
    sheet = workbook.active
    headers = list(df1.columns) + ordered_alarms
    sheet.append(headers)
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    tick_font = Font(name='Calibri', size=11, bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    
    for row_idx, row in result_df.iterrows():
        row_data = [row[col] for col in headers]
        sheet.append(row_data)
        for col_idx, value in enumerate(row_data, start=1):
            cell = sheet.cell(row=row_idx+2, column=col_idx)
            if col_idx > len(df1.columns) and value == '✓':
                cell.fill = green_fill
                cell.font = tick_font
                cell.alignment = center_align
    
    workbook.save(towrite)
    towrite.seek(0)
    st.download_button("Download Highlighted Excel", data=towrite, file_name="Matched_Alarms.xlsx")
    
    # Optional: Hourly trend visualization
    if uploaded_file_3 and uploaded_file_4:
        df_hourly_dc = pd.read_excel(uploaded_file_3)
        df_hourly_node = pd.read_excel(uploaded_file_4)
        st.subheader("Hourly Trend (DC Disconnect & Node Alarms)")
        
        # Convert times
        df_hourly_dc['Start Time'] = pd.to_datetime(df_hourly_dc['Start Time'])
        df_hourly_node['Start Time'] = pd.to_datetime(df_hourly_node['Start Time'])
        
        # Create hourly bins
        df_hourly_dc['Hour'] = df_hourly_dc['Start Time'].dt.floor('H')
        df_hourly_node['Hour'] = df_hourly_node['Start Time'].dt.floor('H')
        
        trend_dc = df_hourly_dc.groupby('Hour').size().reset_index(name='DC Disconnect Count')
        trend_node = df_hourly_node.groupby('Hour').size().reset_index(name='Node Alarm Count')
        
        trend_df = pd.merge(trend_dc, trend_node, on='Hour', how='outer').fillna(0)
        st.line_chart(trend_df.set_index('Hour'))
else:
    st.warning("Upload at least the first two Excel files to continue.")
