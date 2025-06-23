import streamlit as st
import pandas as pd

st.title("📁 Step 1: Upload Alarm Files")

# Upload Virtual Alarm file
virtual_file = st.file_uploader("Upload Virtual Alarm Excel", type=["xlsx", "xls"], key="virtual")

# Upload All Alarm file
all_file = st.file_uploader("Upload All Alarm Excel", type=["xlsx", "xls"], key="all")

# Required columns
required_columns = [
    "Rms Station", "Site Alias", "Zone", "Node",
    "Cluster", "Tenant", "Start Time", "End Time"
]

def validate_columns(df, file_name):
    missing = [col for col in required_columns if col not in df.columns]
    if missing:
        st.error(f"❌ {file_name} is missing columns: {', '.join(missing)}")
        return False
    return True

if virtual_file and all_file:
    try:
        df_virtual = pd.read_excel(virtual_file)
        df_all = pd.read_excel(all_file)

        # Show file previews
        with st.expander("🔍 Preview: Virtual Alarm"):
            st.dataframe(df_virtual)

        with st.expander("🔍 Preview: All Alarm"):
            st.dataframe(df_all)

        # Validate
        if validate_columns(df_virtual, "Virtual Alarm") and validate_columns(df_all, "All Alarm"):
            st.success("✅ Both files uploaded and verified successfully.")
        else:
            st.stop()

    except Exception as e:
        st.error(f"⚠️ Error loading files: {e}")
else:
    st.info("⬆ Please upload both Excel files to proceed.")
