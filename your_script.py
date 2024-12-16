import streamlit as st
import pandas as pd
import os
from io import BytesIO

def merge_excel_files(uploaded_files):
    """
    Merge multiple uploaded Excel files into a single DataFrame.

    Parameters:
        uploaded_files (list): List of uploaded files.

    Returns:
        pd.DataFrame: Merged DataFrame.
    """
    merged_df = pd.DataFrame()

    for file in uploaded_files:
        try:
            # Load the Excel file
            excel_data = pd.ExcelFile(file)
            # Iterate through all sheets
            for sheet_name in excel_data.sheet_names:
                df = excel_data.parse(sheet_name)
                df['Source_File'] = file.name  # Track source file
                df['Sheet_Name'] = sheet_name  # Track sheet name
                merged_df = pd.concat([merged_df, df], ignore_index=True)
        except Exception as e:
            st.error(f"Error processing {file.name}: {e}")

    return merged_df

def download_link(df):
    """
    Create a download link for the merged Excel file.

    Parameters:
        df (pd.DataFrame): The merged DataFrame.
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Merged_Data')
    output.seek(0)
    return output

# Streamlit app
st.title("Excel File Merger with AI Assistance")
st.write("Upload multiple Excel files to merge them into a single file.")

# File uploader
uploaded_files = st.file_uploader("Upload Excel files", type=['xlsx', 'xls'], accept_multiple_files=True)

if uploaded_files:
    st.write(f"You uploaded {len(uploaded_files)} files.")

    # Merge the files
    merged_df = merge_excel_files(uploaded_files)

    if not merged_df.empty:
        st.write("Preview of Merged Data:")
        st.dataframe(merged_df.head())

        # Create a download link
        st.write("Download the merged Excel file:")
        excel_output = download_link(merged_df)
        st.download_button(
            label="Download Merged File",
            data=excel_output,
            file_name="merged_file.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No data to display in the merged file.")
