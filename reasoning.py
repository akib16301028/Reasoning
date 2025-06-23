import streamlit as st
import pandas as pd

st.title("🔁 Virtual vs All Alarm Matcher")

# File uploaders
virtual_file = st.file_uploader("Upload Virtual Alarm Excel file", type=["xlsx"], key="virtual")
all_file = st.file_uploader("Upload All Alarm Excel file", type=["xlsx"], key="all")

# Required columns
required_cols = ["Rms Station", "Site Alias", "Zone", "Node", "Cluster", "Tenant", "Start Time", "End Time"]

def validate_columns(df, name):
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        st.error(f"❌ {name} file is missing columns: {', '.join(missing)}")
        return False
    return True

# Run matching logic
if virtual_file and all_file:
    try:
        df_virtual = pd.read_excel(virtual_file)
        df_all = pd.read_excel(all_file)

        # Ensure datetime format
        df_virtual["Start Time"] = pd.to_datetime(df_virtual["Start Time"])
        df_virtual["End Time"] = pd.to_datetime(df_virtual["End Time"])
        df_all["Start Time"] = pd.to_datetime(df_all["Start Time"])
        df_all["End Time"] = pd.to_datetime(df_all["End Time"])

        # Validate columns
        if validate_columns(df_virtual, "Virtual Alarm") and validate_columns(df_all, "All Alarm"):
            
            matched_nodes = []

            # Iterate through virtual alarm rows
            for idx, v_row in df_virtual.iterrows():
                site = v_row["Site Alias"]
                start = v_row["Start Time"]
                end = v_row["End Time"]

                # Filter all_alarm for matching Site Alias and time within virtual time range
                matched_rows = df_all[
                    (df_all["Site Alias"] == site) &
                    (df_all["Start Time"] >= start) &
                    (df_all["Start Time"] <= end)
                ]

                # Get matching node names as comma-separated string
                node_list = matched_rows["Node"].dropna().unique().tolist()
                matched_nodes.append(", ".join(map(str, node_list)))

            # Add to virtual df
            df_virtual["Matched Nodes from All Alarm"] = matched_nodes

            st.success("✅ Matching complete!")

            with st.expander("📄 Result: Virtual Alarm with Matched Nodes"):
                st.dataframe(df_virtual)

            # Option to download result
            st.download_button(
                label="📥 Download Result as Excel",
                data=df_virtual.to_excel(index=False, engine='openpyxl'),
                file_name="virtual_with_matched_nodes.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"⚠️ Error processing files: {e}")

else:
    st.info("⬆ Please upload both Excel files to continue.")
