import streamlit as st
import pandas as pd
from datetime import timedelta
import matplotlib.pyplot as plt
import numpy as np
from io import BytesIO

st.title("📊 Power System Event Duration Analysis")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)

        required_columns = ['Start Time', 'End Time', 'Node']
        if not all(col in df.columns for col in required_columns):
            st.error(f"Missing required columns: {required_columns}")
        else:
            # Convert datetime columns
            df['Start Time'] = pd.to_datetime(df['Start Time'], errors='coerce')
            df['End Time'] = pd.to_datetime(df['End Time'], errors='coerce')
            df = df.dropna(subset=['Start Time', 'End Time'])

            # Duration in hours
            df['Duration'] = (df['End Time'] - df['Start Time']).dt.total_seconds() / 3600

            # Extract Date and Hour
            df['Date'] = df['Start Time'].dt.date
            df['Hour'] = df['Start Time'].dt.hour

            # Sort and display raw data
            st.subheader("📋 All Event Data")
            st.dataframe(
                df.sort_values('Start Time'),
                column_config={
                    "Start Time": "Start Time",
                    "End Time": "End Time",
                    "Node": "Event Type",
                    "Duration": st.column_config.NumberColumn("Duration (hours)", format="%.2f")
                },
                use_container_width=True,
                hide_index=True
            )

            # Group events
            event_grouped = df.groupby(['Date', 'Hour', 'Node'])['Duration'].sum().unstack(fill_value=0)
            event_types = df['Node'].unique()
            all_dates = sorted(df['Date'].unique())

            st.subheader("📈 Daily Event Duration Analysis")

            for day in all_dates:
                st.markdown(f"### {pd.to_datetime(day).strftime('%A, %Y-%m-%d')}")

                if day in event_grouped.index:
                    day_events = event_grouped.loc[day].reset_index()

                    fig, ax = plt.subplots(figsize=(16, 8))
                    bar_width = 0.8 / len(event_types)
                    positions = np.arange(24)

                    for i, event in enumerate(event_types):
                        if event in day_events.columns:
                            ax.bar(
                                positions + (i * bar_width),
                                day_events[event],
                                width=bar_width,
                                label=event,
                                color=plt.cm.tab20(i)
                            )

                    ax.set_ylabel('Duration (Hours)', color='#333333')
                    ax.set_xlabel('Hour of Day')
                    ax.set_xticks(positions + (len(event_types) * bar_width) / 2)
                    ax.set_xticklabels(range(24))
                    ax.grid(axis='y', linestyle='--', alpha=0.7)

                    max_duration = day_events[event_types].max().max() if not day_events.empty else 0
                    ax.set_ylim(0, max(max_duration * 1.2, 1))
                    plt.title(f"Event Durations on {day}", pad=20)
                    ax.legend(loc='upper left', bbox_to_anchor=(1.05, 1))
                    plt.tight_layout()
                    st.pyplot(fig)

                    cols = st.columns(len(event_types))
                    for i, event in enumerate(event_types):
                        if event in day_events.columns:
                            with cols[i % len(cols)]:
                                st.metric(f"Total {event} Duration", f"{day_events[event].sum():.2f} hours")

                else:
                    st.info("No events recorded on this day.")
                st.markdown("---")

            # Download filtered data
            st.subheader("📥 Download Processed Data")
            output = BytesIO()
            df.to_excel(output, index=False, sheet_name='Event Data')
            output.seek(0)
            st.download_button(
                label="📥 Download Excel File",
                data=output,
                file_name="processed_event_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
else:
    st.info("Please upload an Excel file to begin.")
