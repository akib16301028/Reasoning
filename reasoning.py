import streamlit as st
import pandas as pd
import io

st.title("🔁 Virtual vs All Alarm Matcher (Excel, CSV, TXT supported)")

# Allowed file types
allowed_types = ["xlsx", "xls", "csv", "txt"]

# Upload files
virtual_file = st.file_uploader("Upload Virtual Alarm File", type=allowed_types, key="virtual")
all_file = st.file_uploader("Upload All Alarm File", type=allowed_types, key="all")

# Required columns
required_cols = ["Rms Station", "Site Alias", "Zone", "Node", "Cluster", "Tenant", "Start Time", "End Time"]

# File reading function
def read_file(uploaded_file):
    try:
        if uploaded_file.name.endswith((".xlsx", ".xls")):
            return pd.read_excel(uploaded_file)
        else:
            # Try to detect delimiter
            content = uploaded_file.read()
            uploaded_file.seek(0)
            sample = content.decode(errors='ignore')
            delimiter = "," if "," in sample else "\t"
            return pd.read_csv(io.StringIO(sample), delimiter=delimiter)
    except Exception as e:
        st.error(f"Failed to read file {uploaded_file.name}: {e}")
        return None

# Validation
def validate_columns(df, name):
    missing = [col for col in required_cols if col not in df.columns]
    if missing:
        st.error(f"❌ {name} file is missing columns: {', '.join(missing)}")
        return False
    return True

# Matching logic
if virtual_file and all_file:
    df_virtual = read_file(virtual_file)
    df_all = read_file(all_file)

    if df_virtual is not None and df_all is not None:
        # Convert time columns
        try:
            df_virtual["Start Time"] = pd.to_datetime(df_virtual["Start Time"])
            df_virtual["End Time"] = pd.to_datetime(df_virtual["End Time"])
            df_all["Start Time"] = pd.to_datetime(df_all["Start Time"])
            df_all["End Time"] = pd.to_datetime(df_all["End Time"])
        except Exception as e:
            st.error(f"Date parsing error: {e}")
        else:
            # Validate column names
            if validate_columns(df_virtual, "Virtual Alarm") and validate_columns(df_all, "All Alarm"):
                matched_nodes = []

                for idx, v_row in df_virtual.iterrows():
                    site = v_row["Site Alias"]
                    start = v_row["Start Time"]
                    end = v_row["End Time"]

                    # Find matches in All Alarm
                    matched_rows = df_all[
                        (df_all["Site Alias"] == site) &
                        (df_all["Start Time"] >= start) &
                        (df_all["Start Time"] <= end)
                    ]

                    node_list = matched_rows["Node"].dropna().unique().tolist()
                    matched_nodes.append(", ".join(map(str, node_list)))

                df_virtual["Matched Nodes from All Alarm"] = matched_nodes

                st.success("✅ Matching completed successfully!")

                with st.expander("📄 Preview Result"):
                    st.dataframe(df_virtual)

                # Download result
                output = io.BytesIO()
                df_virtual.to_excel(output, index=False, engine="openpyxl")
                st.download_button(
                    label="📥 Download Result as Excel",
                    data=output.getvalue(),
                    file_name="virtual_with_matched_nodes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
else:
    st.info("⬆ Please upload both alarm files (Excel, CSV, or TXT)")
