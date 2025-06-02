import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

# Streamlit App
st.title("DC Disconnect Virtual Alarm Highlighter")

# Upload first Excel (DC Disconnect Virtual Alarm)
uploaded_file_1 = st.file_uploader("Upload DC Disconnect Virtual Alarm Excel", type=["xlsx"])

# Upload second Excel (Node Alarms)
uploaded_file_2 = st.file_uploader("Upload Node Alarms Excel", type=["xlsx"])

if uploaded_file_1 and uploaded_file_2:
    # Read the Excel files into DataFrames
    df1 = pd.read_excel(uploaded_file_1)
    df2 = pd.read_excel(uploaded_file_2)
    
    # Normalize site names by stripping spaces
    df1['Site'] = df1['Site'].str.strip()
    df2['Site'] = df2['Site'].str.strip()
    
    # Get unique alarm names from the second Excel's 'Node' column
    unique_alarms = df2['Node'].unique()
    
    # Initialize a result DataFrame based on the first Excel
    result_df = df1.copy()
    
    for alarm in unique_alarms:
        # Create a new column for each alarm, initially empty
        result_df[alarm] = ''
    
    # Fill in the result DataFrame
    for idx, row in result_df.iterrows():
        site = row['Site']
        matching_alarms = df2[df2['Site'] == site]['Node'].tolist()
        for alarm in unique_alarms:
            if alarm in matching_alarms:
                result_df.at[idx, alarm] = '✅'
    
    # Display the result with conditional formatting
    def highlight_alarms(val):
        if val == '✅':
            return 'background-color: lightgreen'
        return ''

    st.write("Matched Sites with Alarm Highlights")
    st.dataframe(result_df.style.applymap(highlight_alarms, subset=unique_alarms))
    
    # Optionally, allow downloading the result
    towrite = BytesIO()
    result_df.to_excel(towrite, index=False)
    towrite.seek(0)
    st.download_button("Download Result Excel", data=towrite, file_name="highlighted_result.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.warning("Please upload both Excel files to continue.")
