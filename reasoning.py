import streamlit as st
import pandas as pd
import io

st.title("⚡ Fast Virtual vs All Alarm Matcher")

# File upload
virtual_file = st.file_uploader("Upload Virtual Alarm file", type=["xlsx","xls","csv","txt"])
all_file = st.file_uploader("Upload All Alarm file", type=["xlsx","xls","csv","txt"])

def read_file(f):
    if f is None: return None
    if f.name.endswith((".xlsx","xls")):
        return pd.read_excel(f)
    content = f.getvalue().decode(errors='ignore')
    sep = ',' if content.count(',') > content.count('\t') else '\t'
    return pd.read_csv(io.StringIO(content), sep=sep)

if virtual_file and all_file:

    df_v = read_file(virtual_file)
    df_a = read_file(all_file)

    try:
        for df in [df_v, df_a]:
            df["Start Time"] = pd.to_datetime(df["Start Time"])
            df["End Time"] = pd.to_datetime(df["End Time"])
    except Exception as e:
        st.error(f"Error parsing dates: {e}")
        st.stop()

    required = ["Site Alias", "Node", "Start Time", "End Time"]
    for name, df in [("Virtual",df_v),("All",df_a)]:
        miss = [c for c in required if c not in df.columns]
        if miss:
            st.error(f"{name} file missing columns: {', '.join(miss)}")
            st.stop()

    st.info("💡 Starting fast merge...")

    # Explode virtual time windows into intervals
    df_v = df_v.assign(idx=range(len(df_v)))
    
    # Use cartesian merge on Site Alias
    merged = df_v.merge(df_a[['Site Alias','Node','Start Time','End Time']], on='Site Alias', how='left', suffixes=('','_all'))
    
    # Filter rows where all-alarm start is within the virtual window
    cond = (merged['Start Time_all'] >= merged['Start Time']) & (merged['Start Time_all'] <= merged['End Time'])
    filtered = merged[cond]

    # Group nodes
    grouped = filtered.groupby('idx')['Node_all'].agg(lambda nodes: ", ".join(sorted(set(nodes)))).reset_index()
    
    # Merge results back
    df_v = df_v.merge(grouped, on='idx', how='left').rename(columns={'Node_all': 'Matched Nodes from All Alarm'})
    df_v['Matched Nodes from All Alarm'] = df_v['Matched Nodes from All Alarm'].fillna("")

    st.success("✅ Matching done in seconds!")

    with st.expander("📄 Preview Result"):
        st.dataframe(df_v)

    out = io.BytesIO()
    df_v.to_excel(out, index=False, engine='openpyxl')
    st.download_button("📥 Download Excel", out.getvalue(), "matched_result.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("⬆ Please upload both alarm files.")
