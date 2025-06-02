import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from io import BytesIO

# Streamlit App
st.title("DC Disconnect Virtual Alarm Highlighter & Downloader")

# Upload first Excel (DC Disconnect Virtual Alarm)
uploaded_file_1 = st.file_uploader("Upload DC Disconnect Virtual Alarm Excel", type=["xlsx"])

# Upload second Excel (Node Alarms)
uploaded_file_2 = st.file_uploader("Upload Node Alarms Excel", type=["xlsx"])

if uploaded_file_1 and uploaded_file_2:
    # Read Excel files
    df1 = pd.read_excel(uploaded_file_1)
    df2 = pd.read_excel(uploaded_file_2)
    
    # Clean up site names
    df1['Site'] = df1['Site'].astype(str).str.strip()
    df2['Site'] = df2['Site'].astype(str).str.strip()
    
    # Extract unique alarms from Node column
    all_unique_alarms = sorted(df2['Node'].unique())
    
    # Define custom alarm order
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
    
    # Sort alarms: first custom order, then rest (excluding duplicates)
    remaining_alarms = [alarm for alarm in all_unique_alarms if alarm not in custom_order]
    ordered_alarms = custom_order + remaining_alarms
    
    # Initialize result DataFrame
    result_df = df1.copy()
    for alarm in ordered_alarms:
        result_df[alarm] = ''
    
    # Update the result DataFrame with matches
    for idx, row in result_df.iterrows():
        site = row['Site']
        matching_alarms = df2[df2['Site'] == site]['Node'].tolist()
        for alarm in ordered_alarms:
            if alarm in matching_alarms:
                result_df.at[idx, alarm] = '✓'  # Light thin checkmark
    
    # Show DataFrame with Streamlit styling
    def highlight_alarms(val):
        if val == '✓':
            return 'background-color: lightgreen'
        return ''
    
    st.write("Matched Sites with Alarm Highlights")
    st.dataframe(result_df.style.applymap(highlight_alarms, subset=ordered_alarms))
    
    # Prepare Excel with formatting
    towrite = BytesIO()
    workbook = Workbook()
    sheet = workbook.active
    
    # Write headers
    headers = list(df1.columns) + ordered_alarms
    sheet.append(headers)
    
    # Define styles
    green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    tick_font = Font(name='Calibri', size=11, bold=True)
    center_align = Alignment(horizontal='center', vertical='center')
    
    # Write data rows
    for row_idx, row in result_df.iterrows():
        row_data = [row[col] for col in headers]
        sheet.append(row_data)
        for col_idx, value in enumerate(row_data, start=1):
            cell = sheet.cell(row=row_idx + 2, column=col_idx)
            if col_idx > len(df1.columns) and value == '✓':
                cell.fill = green_fill
                cell.font = tick_font
                cell.alignment = center_align
    
    # Adjust column widths
    for col in sheet.columns:
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col) + 2
        col_letter = col[0].column_letter
        sheet.column_dimensions[col_letter].width = max_length
    
    # Save to BytesIO
    workbook.save(towrite)
    towrite.seek(0)
    
    st.download_button(
        label="Download Highlighted Excel File",
        data=towrite,
        file_name="DC_Disconnect_Alarm_Highlighted.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.warning("Please upload both Excel files to continue.")
