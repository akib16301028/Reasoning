import streamlit as st
import pandas as pd
from datetime import timedelta
import matplotlib.pyplot as plt
import seaborn as sns

st.title("📊 3-Day Hourly Alarm Breakdown")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # ✅ Load Excel, taking column headers from Row 1
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

            # Graphical Representation Section
            st.subheader("📈 Graphical Representation")
            
            # Option to select visualization type
            viz_type = st.selectbox("Select Visualization Type", 
                                   ["Line Chart", "Bar Chart", "Heatmap"])
            
            # Reset index for plotting
            plot_df = pivot.reset_index()
            plot_df['Date'] = pd.to_datetime(plot_df['Date'])
            plot_df['Date_Hour'] = plot_df['Date'].astype(str) + ' ' + plot_df['Hour'].astype(str) + ':00'
            
            if viz_type == "Line Chart":
                st.write("### Alarm Trends Over Time")
                fig, ax = plt.subplots(figsize=(12, 6))
                for node in pivot.columns.get_level_values(0).unique():
                    if node not in ['Date', 'Hour']:
                        ax.plot(plot_df['Date_Hour'], plot_df[node], label=node)
                plt.xticks(rotation=45)
                plt.xlabel('Date & Hour')
                plt.ylabel('Alarm Count')
                plt.title('Hourly Alarm Count by Type')
                plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
                plt.tight_layout()
                st.pyplot(fig)
                
            elif viz_type == "Bar Chart":
                st.write("### Hourly Alarm Distribution")
                selected_date = st.selectbox("Select Date to View", plot_df['Date'].dt.date.unique())
                date_df = plot_df[plot_df['Date'].dt.date == selected_date]
                
                fig, ax = plt.subplots(figsize=(12, 6))
                date_df.set_index('Hour').drop(['Date', 'Date_Hour'], axis=1).plot(kind='bar', stacked=True, ax=ax)
                plt.xlabel('Hour of Day')
                plt.ylabel('Alarm Count')
                plt.title(f'Alarm Distribution by Hour on {selected_date}')
                plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
                plt.tight_layout()
                st.pyplot(fig)
                
            elif viz_type == "Heatmap":
                st.write("### Alarm Heatmap by Hour and Type")
                # Select top N alarm types to show
                top_n = st.slider("Select number of top alarm types to display", 3, 15, 5)
                
                # Get top alarm types by total count
                alarm_totals = plot_df.drop(['Date', 'Hour', 'Date_Hour'], axis=1).sum().sort_values(ascending=False)
                top_alarms = alarm_totals.head(top_n).index.tolist()
                
                # Prepare data for heatmap
                heatmap_data = plot_df.set_index(['Date', 'Hour'])[top_alarms]
                
                fig, ax = plt.subplots(figsize=(12, 8))
                sns.heatmap(heatmap_data.T, cmap="YlOrRd", annot=True, fmt="d", linewidths=.5, ax=ax)
                plt.title(f'Top {top_n} Alarm Types by Hour and Date')
                plt.xlabel('Date & Hour')
                plt.ylabel('Alarm Type')
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
