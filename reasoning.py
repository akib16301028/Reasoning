import streamlit as st
import pandas as pd
from datetime import timedelta
import matplotlib.pyplot as plt
import numpy as np

st.title("📊 Power System Event Analysis")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load Excel
        df = pd.read_excel(uploaded_file)

        required_columns = ['Start Time', 'End Time', 'Node']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns: {required_columns}")
        else:
            # Convert datetime columns
            df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
            df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce')
            df = df.dropna(subset=['Start Time', 'End Time'])
            
            # Calculate duration in hours
            df['Duration'] = (df['End Time'] - df['Start Time']).dt.total_seconds() / 3600
            
            # Extract date and hour (using start time)
            df['Date'] = df['Start Time'].dt.date
            df['Hour'] = df['Start Time'].dt.hour

            # Filter last 3 days
            max_date = df['Date'].max()
            last_3_days = [max_date - timedelta(days=i) for i in range(3)]
            df_filtered = df[df['Date'].isin(last_3_days)].copy()
            
            # Create a combined date-hour column for plotting
            df_filtered['Date_Hour'] = df_filtered['Date'].astype(str) + ' ' + df_filtered['Hour'].astype(str) + ':00'
            
            # Separate the data into two groups:
            # 1. Power events (Mains Fail, PG Run, DG Run) - for bar chart (duration)
            # 2. Outage events (DCDB-01 Primary Disconnect) - for line chart (count)
            
            power_events = ['Mains Fail', 'PG Run', 'DG Run']
            outage_event = 'DCDB-01 Primary Disconnect'
            
            # Process power events data (duration)
            power_df = df_filtered[df_filtered['Node'].isin(power_events)]
            power_grouped = power_df.groupby(['Date', 'Hour', 'Node'])['Duration'].sum().unstack(fill_value=0)
            
            # Process outage events data (count)
            outage_df = df_filtered[df_filtered['Node'] == outage_event]
            outage_grouped = outage_df.groupby(['Date', 'Hour']).size().reset_index(name='Count')
            
            # Create visualization for each day
            st.subheader("🔌 Power Event Analysis (Last 3 Days)")
            
            for day in last_3_days:
                st.markdown(f"### {day.strftime('%A, %Y-%m-%d')}")
                
                # Filter data for this day
                day_power = power_grouped.loc[day]
                day_outage = outage_grouped[outage_grouped['Date'] == day]
                
                # Create figure with secondary y-axis
                fig, ax1 = plt.subplots(figsize=(14, 7))
                
                # Plot power events as stacked bars (duration)
                day_power.plot(kind='bar', stacked=True, ax=ax1, 
                              color=['#FF6B6B', '#4ECDC4', '#45B7D1'], 
                              width=0.8, position=0)
                
                ax1.set_ylabel('Duration (Hours)', color='#333333')
                ax1.set_xlabel('Hour of Day')
                ax1.set_xticks(range(0, 24))
                ax1.set_xticklabels(range(0, 24), rotation=0)
                ax1.grid(axis='y', linestyle='--', alpha=0.7)
                ax1.set_ylim(0, max(day_power.sum(axis=1).max() * 1.2, 1))
                
                # Create second y-axis for outage count
                ax2 = ax1.twinx()
                
                # Plot outage events as line (count)
                if not day_outage.empty:
                    ax2.plot(day_outage['Hour'], day_outage['Count'], 
                            color='#FFA500', marker='o', linestyle='-', 
                            linewidth=2, markersize=8, label='Outage Count')
                
                ax2.set_ylabel('Outage Event Count', color='#FFA500')
                ax2.tick_params(axis='y', labelcolor='#FFA500')
                ax2.set_ylim(0, max(day_outage['Count'].max() * 1.5, 1) if not day_outage.empty else 0)
                
                # Customize the plot
                plt.title(f"Power System Events on {day.strftime('%Y-%m-%d')}\n"
                         f"Bars show duration (hours), Line shows outage count", pad=20)
                
                # Combine legends from both axes
                lines1, labels1 = ax1.get_legend_handles_labels()
                lines2, labels2 = ax2.get_legend_handles_labels()
                ax1.legend(lines1 + lines2, labels1 + labels2, 
                          loc='upper left', bbox_to_anchor=(1.1, 1))
                
                plt.tight_layout()
                st.pyplot(fig)
                
                # Display summary statistics
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Mains Fail Duration", 
                             f"{day_power['Mains Fail'].sum():.2f} hours")
                with col2:
                    st.metric("Total PG Run Duration", 
                             f"{day_power['PG Run'].sum():.2f} hours")
                with col3:
                    st.metric("Total Outage Events", 
                             f"{day_outage['Count'].sum() if not day_outage.empty else 0}")
                
                st.markdown("---")
            
            # Save to Excel
            output_excel = pd.ExcelWriter("power_analysis.xlsx", engine='openpyxl')
            power_grouped.to_excel(output_excel, sheet_name="Power Events")
            outage_grouped.to_excel(output_excel, sheet_name="Outage Events")
            output_excel.close()
            
            with open("power_analysis.xlsx", "rb") as f:
                st.download_button("⬇ Download Analysis Excel", f, file_name="power_system_analysis.xlsx")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to begin.")
