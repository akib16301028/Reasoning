import streamlit as st
import pandas as pd

def main():
    st.title("Group Reasoning and Display Corresponding Remarks with Elapsed Time")

    # File upload
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

    if uploaded_file is not None:
        # Load the Excel file into a Pandas DataFrame
        try:
            df = pd.read_excel(uploaded_file)
            st.write("### Uploaded Data:")
            st.dataframe(df)

            # Check if required columns exist
            if {'Remarks', 'Elapsed Time Count', 'Reasoning'}.issubset(df.columns):
                # Group by 'Reasoning' and aggregate Remarks with corresponding Elapsed Time Count
                grouped_data = (
                    df.groupby('Reasoning', as_index=False)
                    .apply(lambda group: pd.DataFrame({
                        'Reasoning': [group['Reasoning'].iloc[0]] * len(group),
                        'Remarks': group['Remarks'],
                        'Elapsed Time Count': group['Elapsed Time Count']
                    }))
                    .reset_index(drop=True)
                )

                # Sort by Elapsed Time Count in descending order
                grouped_data = grouped_data.sort_values(by='Elapsed Time Count', ascending=False)

                st.write("### Grouped Data by Reasoning with Remarks and Elapsed Time Count:")
                st.dataframe(grouped_data)

                # Option to download the grouped data
                download_data = grouped_data.to_csv(index=False)
                st.download_button(
                    label="Download Grouped Data as CSV",
                    data=download_data,
                    file_name="grouped_data_by_reasoning.csv",
                    mime="text/csv"
                )
            else:
                st.error("The uploaded file must contain 'Remarks', 'Elapsed Time Count', and 'Reasoning' columns.")
        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")

if __name__ == "__main__":
    main()
