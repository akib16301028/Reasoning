import streamlit as st
import pandas as pd
from datetime import timedelta
import matplotlib.pyplot as plt
import numpy as np
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.drawing.image import Image as ExcelImage
import io

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
            
            # Power events and outage event definitions
            power_events = ['Mains Fail', 'PG Run', 'DG Run']
            outage_event = 'DCDB-01 Primary Disconnect'
            
            # Process data
            power_df = df_filtered[df_filtered['Node'].isin(power_events)]
            power_grouped = power_df.groupby(['Date', 'Hour', 'Node'])['Duration'].sum().unstack(fill_value=0)
            outage_df = df_filtered[df_filtered['Node'] == outage_event]
            outage_grouped = outage_df.groupby(['Date', 'Hour']).size().reset_index(name='Count')
            
            # Create Excel workbook with charts
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                # Create visualization for each day
                st.subheader("🔌 Power Event Analysis (Last 3 Days)")
                
                for day in last_3_days:
                    day_str = day.strftime('%Y-%m-%d')
                    st.markdown(f"### {day.strftime('%A, %Y-%m-%d')}")
                    
                    # Filter data for this day
                    day_power = power_grouped.loc[day].reset_index()
                    day_outage = outage_grouped[outage_grouped['Date'] == day]
                    
                    # Create figure for Streamlit
                    fig, ax1 = plt.subplots(figsize=(16, 8))
                    self.create_streamlit_chart(ax1, day_power, day_outage, power_events, day_str)
                    st.pyplot(fig)
                    plt.close(fig)
                    
                    # Write data to Excel
                    day_power.to_excel(writer, sheet_name=f"Data_{day_str}", index=False)
                    if not day_outage.empty:
                        day_outage.to_excel(writer, sheet_name=f"Outage_{day_str}", index=False)
                    
                    # Create Excel charts
                    workbook = writer.book
                    self.create_excel_charts(workbook, day_power, day_outage, day_str, power_events)
                    
                    # Display metrics
                    self.display_metrics(day_power, day_outage, power_events)
                    st.markdown("---")
                
                # Save the workbook
                writer.save()
            
            # Download button for Excel file
            excel_buffer.seek(0)
            st.download_button(
                label="⬇ Download Full Analysis Workbook with Charts",
                data=excel_buffer,
                file_name="power_analysis_with_charts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to begin.")

def create_streamlit_chart(self, ax1, day_power, day_outage, power_events, day_str):
    """Create the Streamlit visualization"""
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
    plt.title(f"Power System Events on {day_str}\n"
             f"Bars show duration (hours), Line shows outage count", pad=20)
    
    # Combine legends
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, 
              loc='upper left', bbox_to_anchor=(1.1, 1))

def create_excel_charts(self, workbook, day_power, day_outage, day_str, power_events):
    """Create native Excel charts in the workbook"""
    # Create a new sheet for charts
    chart_sheet = workbook.create_sheet(title=f"Charts_{day_str}")
    
    # Write formatted data to the sheet
    chart_sheet.append(["Hour"] + power_events + ["Outage Count"])
    for hour in range(24):
        row = [hour]
        # Add power event durations
        for event in power_events:
            val = day_power[day_power['Hour'] == hour][event].values
            row.append(val[0] if len(val) > 0 else 0)
        # Add outage count
        outage_val = day_outage[day_outage['Hour'] == hour]['Count'].values
        row.append(outage_val[0] if len(outage_val) > 0 else 0)
        chart_sheet.append(row)
    
    # Create combo chart
    combo_chart = BarChart()
    combo_chart.type = "col"
    combo_chart.style = 10
    combo_chart.title = f"Power Events - {day_str}"
    combo_chart.y_axis.title = 'Duration (Hours)'
    combo_chart.x_axis.title = 'Hour of Day'
    
    # Add power events data
    data = Reference(chart_sheet, min_col=2, min_row=1, max_col=len(power_events)+1, max_row=25)
    cats = Reference(chart_sheet, min_col=1, min_row=2, max_row=25)
    combo_chart.add_data(data, titles_from_data=True)
    combo_chart.set_categories(cats)
    
    # Add outage line to secondary axis
    line_chart = LineChart()
    outage_data = Reference(chart_sheet, min_col=len(power_events)+2, min_row=1, max_col=len(power_events)+2, max_row=25)
    line_chart.add_data(outage_data, titles_from_data=True)
    combo_chart.y_axis.crosses = "max"
    combo_chart += line
    
    # Position and add the chart
    chart_sheet.add_chart(combo_chart, "A30")
    
    # Create separate bar chart for power events
    bar_chart = BarChart()
    bar_chart.type = "col"
    bar_chart.style = 10
    bar_chart.title = f"Power Duration - {day_str}"
    bar_chart.y_axis.title = 'Duration (Hours)'
    bar_chart.add_data(data, titles_from_data=True)
    bar_chart.set_categories(cats)
    chart_sheet.add_chart(bar_chart, "A60")
    
    # Create separate line chart for outages
    if not day_outage.empty:
        line_chart = LineChart()
        line_chart.title = f"Outage Events - {day_str}"
        line_chart.y_axis.title = 'Count'
        line_chart.add_data(outage_data, titles_from_data=True)
        line_chart.set_categories(cats)
        chart_sheet.add_chart(line_chart, "J60")

def display_metrics(self, day_power, day_outage, power_events):
    """Display summary metrics in Streamlit"""
    cols = st.columns(len(power_events) + 1)
    for i, event in enumerate(power_events):
        with cols[i]:
            st.metric(f"Total {event} Duration", 
                     f"{day_power[event].sum():.2f} hours")
    with cols[-1]:
        st.metric("Total Outage Events", 
                 f"{day_outage['Count'].sum() if not day_outage.empty else 0}")
