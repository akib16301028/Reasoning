import streamlit as st
import pandas as pd
from datetime import timedelta
import matplotlib.pyplot as plt
import seaborn as sns

st.title("📊 3-Day Hourly Alarm Breakdown")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Load Excel
        df = pd.read_excel(uploaded_file)

        required_columns = ['Start Time', 'Node']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns: {required_columns}")
        else:
            # Convert Start Time to datetime
            df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
            df = df.dropna(subset=['Start Time'])

            # Extract date and hour
            df['Date'] = df['Start Time'].dt.date
            df['Hour'] = df['Start Time'].dt.hour

            # Filter last 3 days
            max_date = df['Date'].max()
            last_3_days = [max_date - timedelta(days=i) for i in range(3)]
            df_filtered = df[df['Date'].isin(last_3_days)]

            # Group by Date, Hour, Node
            grouped = df_filtered.groupby(['Date', 'Hour', 'Node']).size().reset_index(name='Alarm Count')

            # Pivot to show Node (alarm type) as columns
            pivot = grouped.pivot_table(index=['Date', 'Hour'], columns='Node', values='Alarm Count', fill_value=0)
            
            # Display result
            st.subheader("📅 Hourly Alarm Breakdown by Type (Last 3 Days)")
            st.dataframe(pivot)

            # Graphical Representation - Separate Graphs for Each Day
            st.subheader("📈 Daily Hourly Alarm Trends")
            
            # Get unique dates
            unique_dates = grouped['Date'].unique()
            
            # Select alarm types to display
            all_alarms = grouped['Node'].unique()
            selected_alarms = st.multiselect(
                "Select alarm types to display",
                options=all_alarms,
                default=all_alarms[:min(5, len(all_alarms))]  # Show first 5 by default
            )
            
            # Create a separate graph for each day
            for day in unique_dates:
                st.markdown(f"### {day.strftime('%A, %Y-%m-%d')}")
                
                # Filter data for this day
                day_data = grouped[(grouped['Date'] == day) & (grouped['Node'].isin(selected_alarms))]
                
                # Create figure
                fig, ax = plt.subplots(figsize=(12, 6))
                
                # Plot each selected alarm type
                for alarm in selected_alarms:
                    alarm_data = day_data[day_data['Node'] == alarm]
                    ax.plot(alarm_data['Hour'], alarm_data['Alarm Count'], 
                            marker='o', label=alarm)
                
                # Customize plot
                ax.set_xlabel('Hour of Day')
                ax.set_ylabel('Alarm Count')
                ax.set_title(f'Hourly Alarm Counts on {day.strftime("%Y-%m-%d")}')
                ax.set_xticks(range(0, 24))
                ax.grid(True, linestyle='--', alpha=0.7)
                ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
                
                plt.tight_layout()
                st.pyplot(fig)

            # Save to Excel
            output_excel = pd.ExcelWriter("alarm_breakdown.xlsx", engine='openpyxl')
            pivot.to_excel(output_excel)
            output_excel.close()

            with open("alarm_breakdown.xlsx", "rb") as f:
                st.download_button("⬇ Download Breakdown Excel", f, file_name="3_day_alarm_breakdown.xlsx")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to begin.")
