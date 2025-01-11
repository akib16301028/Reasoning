import streamlit as st
import pandas as pd


def main():
    st.title("Group Remarks and Calculate Elapsed Time with Reasoning")

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
                # Group by 'Remarks', aggregate elapsed time sum, and list reasoning values
                grouped_data = (
                    df.groupby('Remarks', as_index=False)
                    .agg({
                        'Elapsed Time Count': 'sum',
                        'Reasoning': lambda x: ', '.join(sorted(set(x)))  # Combine unique reasoning values
                    })
                )

                st.write("### Grouped Data with Elapsed Time Sum and Reasoning:")
                st.dataframe(grouped_data)

                # Option to download the grouped data
                download_data = grouped_data.to_csv(index=False)
                st.download_button(
                    label="Download Grouped Data as CSV",
                    data=download_data,
                    file_name="grouped_data_with_reasoning.csv",
                    mime="text/csv"
                )
            else:
                st.error("The uploaded file must contain 'Remarks', 'Elapsed Time Count', and 'Reasoning' columns.")
        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")


if __name__ == "__main__":
    main()
