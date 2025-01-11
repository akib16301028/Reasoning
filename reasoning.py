import streamlit as st
import pandas as pd

def main():
    st.title("Group Remarks and Calculate Elapsed Time")

    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Load the Excel file into a Pandas DataFrame
        try:
            df = pd.read_excel(uploaded_file)
            st.write("### Uploaded Data:")
            st.dataframe(df)

            # Check if required columns exist
            if {'Remarks', 'Elapsed Time Count'}.issubset(df.columns):
                # Group by 'Remarks' and sum 'Elapsed Time Count'
                grouped_data = df.groupby('Remarks', as_index=False)['Elapsed Time Count'].sum()

                st.write("### Grouped Data with Elapsed Time Sum:")
                st.dataframe(grouped_data)

                # Option to download the grouped data
                download_data = grouped_data.to_csv(index=False)
                st.download_button(
                    label="Download Grouped Data as CSV",
                    data=download_data,
                    file_name="grouped_data.csv",
                    mime="text/csv"
                )
            else:
                st.error("The uploaded file must contain 'Remarks' and 'Elapsed Time Count' columns.")
        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")

if __name__ == "__main__":
    main()
