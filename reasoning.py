import streamlit as st
import pandas as pd
from datetime import timedelta
import matplotlib.pyplot as plt
import numpy as np

st.title("📊 Power System Event Duration Analysis")

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
            
            # Group all events by type and hour to get total duration
            event_grouped = df_filtered.groupby(['Date', 'Hour', 'Node'])['Duration'].sum().unstack(fill_value=0)
            
            # Get unique event types for coloring
            event_types = df_filtered['Node'].unique()
            
            # Create visualization for each day
            st.subheader("🔌 Event Duration Analysis (Last 3 Days)")
            
            for day in last_3_days:
                st.markdown(f"### {day.strftime('%A, %Y-%m-%d')}")
                
                # Filter data for this day
                day_events = event_grouped.loc[day].reset_index()
                
                # Create figure
                fig, ax = plt.subplots(figsize=(16, 8))
                
                # Set width and positions for side-by-side bars
                bar_width = 0.8 / len(event_types)  # Dynamic width based on number of event types
                positions = np.arange(24)
                
                # Plot each event type as separate bars
                for i, event in enumerate(event_types):
                    if event in day_events.columns:  # Check if event occurred this day
                        ax.bar(
                            positions + (i * bar_width), 
                            day_events[event], 
                            width=bar_width, 
                            label=event,
                            color=plt.cm.tab20(i)  # Use a colormap for distinct colors
                        )
                
                ax.set_ylabel('Duration (Hours)', color='#333333')
                ax.set_xlabel('Hour of Day')
                ax.set_xticks(positions + (len(event_types) * bar_width) / 2)
                ax.set_xticklabels(range(24))
                ax.grid(axis='y', linestyle='--', alpha=0.7)
                
                # Set y-axis limit with some padding
                max_duration = day_events[event_types].max().max() if not day_events.empty else 0
                ax.set_ylim(0, max(max_duration * 1.2, 1))
                
                # Customize the plot
                plt.title(f"Power System Event Durations on {day.strftime('%Y-%m-%d')}", pad=20)
                
                # Add legend outside the plot
                ax.legend(loc='upper left', bbox_to_anchor=(1.05, 1))
                
                plt.tight_layout()
                st.pyplot(fig)
                
                # Display summary statistics
                cols = st.columns(len(event_types))
                for i, event in enumerate(event_types):
                    if event in day_events.columns:
                        with cols[i % len(cols)]:
                            st.metric(f"Total {event} Duration", 
                                     f"{day_events[event].sum():.2f} hours")
                
                st.markdown("---")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to begin.")
