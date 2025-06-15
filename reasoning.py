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
                day_power = power_grouped.loc[day].reset_index()
                day_outage = outage_grouped[outage_grouped['Date'] == day]
                
                # Create figure with secondary y-axis
                fig, ax1 = plt.subplots(figsize=(16, 8))
                
                # Set width and positions for side-by-side bars
                bar_width = 0.25
                positions = np.arange(24)
                
                # Plot each power event as separate bars
                for i, event in enumerate(power_events):
                    ax1.bar(
                        positions + (i * bar_width), 
                        day_power[event], 
                        width=bar_width, 
                        label=event,
                        color=['#FF6B6B', '#4ECDC4', '#45B7D1'][i]
                    )
                
                ax1.set_ylabel('Duration (Hours)', color='#333333')
                ax1.set_xlabel('Hour of Day')
                ax1.set_xticks(positions + bar_width)
                ax1.set_xticklabels(range(24))
                ax1.grid(axis='y', linestyle='--', alpha=0.7)
                ax1.set_ylim(0, max(day_power[power_events].max().max() * 1.2, 1))
                
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
                cols = st.columns(len(power_events) + 1)
                for i, event in enumerate(power_events):
                    with cols[i]:
                        st.metric(f"Total {event} Duration", 
                                 f"{day_power[event].sum():.2f} hours")
                with cols[-1]:
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
