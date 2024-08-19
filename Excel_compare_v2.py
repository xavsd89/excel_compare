import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import streamlit as st
import tempfile

def extract_and_merge_columns(old_file, new_file, old_cols, new_cols):
    """
    Extracts specific columns from two Excel files, merges them into a new dataset, and saves it to a file.

    Parameters:
    - old_file (pd.DataFrame): DataFrame loaded from the old Excel file.
    - new_file (pd.DataFrame): DataFrame loaded from the new Excel file.
    - old_cols (list): List of column names to extract from the old file.
    - new_cols (list): List of column names to extract from the new file.

    Returns:
    - str: Path to the output Excel file containing the merged columns.
    """
    # Extract specified columns from old and new files
    old_data = old_file[old_cols]
    new_data = new_file[new_cols]

    # Merge the extracted columns (assuming we are merging based on a common key, e.g., 'ID')
    # Adjust 'key' to the actual column name used for merging
    merged_data = pd.merge(old_data, new_data, left_index=True, right_index=True, suffixes=('_old', '_new'))

    # Create a temporary file to save the merged data
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        merged_file_path = temp_file.name

    with pd.ExcelWriter(merged_file_path, engine='openpyxl') as writer:
        merged_data.to_excel(writer, sheet_name='Merged', index=False)

    return merged_file_path

def highlight_col(file_path):
    """
    Highlights cells in the 'Differences' sheet of an Excel file based on their merge status.

    Parameters:
    - file_path (str): The path to the Excel file to be processed.
    """
    wb = load_workbook(file_path)
    ws_diff = wb['Differences']

    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    pink_fill = PatternFill(start_color='FFC0CB', end_color='FFC0CB', fill_type='solid')

    for row in ws_diff.iter_rows(min_row=2, max_col=ws_diff.max_column):
        merge_output = row[-1].value  # The merge status is assumed to be in the last column
        for cell in row:
            if merge_output == 'Source Only':
                cell.fill = yellow_fill
            elif merge_output == 'Target Only':
                cell.fill = pink_fill

    wb.save(file_path)

def main():
    st.title('Excel Comparison & Merging Tool (v3)')

    # Upload widgets for the old and new Excel files for column extraction
    old_file_upload = st.file_uploader("Upload Excel A for merging", type=['xlsx'], key='old')
    new_file_upload = st.file_uploader("Upload Excel B for merging", type=['xlsx'], key='new')

    # Specify the columns to extract from the old and new files
    old_columns = st.text_input("Enter Column names to extract from the Excel A, separated by commas", "Column1,Column2..etc")
    new_columns = st.text_input("Enter Column names to extract from the Excel B, separated by commas", "Column1,Column2..etc")

    if old_file_upload and new_file_upload:
        try:
            # Read the uploaded Excel files into pandas DataFrames
            old_file = pd.read_excel(old_file_upload)
            new_file = pd.read_excel(new_file_upload)

            # Convert input columns to lists
            old_cols = [col.strip() for col in old_columns.split(',')]
            new_cols = [col.strip() for col in new_columns.split(',')]

            # Extract and merge specific columns
            merged_file_path = extract_and_merge_columns(old_file, new_file, old_cols, new_cols)

            # Provide download link for the merged file
            st.write("Merged columns have been successfully extracted and merged.")
            with open(merged_file_path, "rb") as f:
                st.download_button("Download Merged File", f, file_name="merged_output.xlsx")

            # Clean up temporary file
            os.remove(merged_file_path)

        except Exception as e:
            st.error(f"An error occurred: {e}")

    # File upload widgets for the source and target files for comparison
    uploaded_source_file = st.file_uploader("Upload SOURCE File for Comparison", type=['xlsx'], key='source')
    uploaded_target_file = st.file_uploader("Upload TARGET File for Comparison", type=['xlsx'], key='target')

    if uploaded_source_file and uploaded_target_file:
        try:
            # Read the uploaded Excel files into pandas DataFrames
            source_file = pd.read_excel(uploaded_source_file)
            target_file = pd.read_excel(uploaded_target_file)

            # Strip whitespace from string entries in the DataFrames
            source_file = source_file.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)
            target_file = target_file.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)

            # Merge the DataFrames to find differences
            diff = source_file.merge(target_file, indicator=True, how='outer')
            diff['_merge'].replace({'left_only': 'Source Only', 'right_only': 'Target Only'}, inplace=True)
            diff.rename(columns={'_merge': 'Merge_Output'}, inplace=True)
            diff = diff[diff['Merge_Output'] != 'both']

            # Path for the output Excel file
            diff_file_path = 'diff_output.xlsx'

            # Check if there are differences to write to the Excel file
            if not diff.empty:
                with pd.ExcelWriter(diff_file_path) as writer:
                    source_file.to_excel(writer, sheet_name='Source', index=False)
                    target_file.to_excel(writer, sheet_name='Target', index=False)
                    diff.to_excel(writer, sheet_name='Differences', index=False)
                
                # Highlight differences in the Excel file
                highlight_col(diff_file_path)

                # Display the number of differences and provide a download button for the file
                st.write(f"Found {len(diff)} differences.")
                with open(diff_file_path, "rb") as f:
                    st.download_button("Download Differences File", f, file_name="diff_output.xlsx")
            else:
                # If no differences found, write an empty DataFrame and a message to the Excel file
                with pd.ExcelWriter(diff_file_path) as writer:
                    source_file.to_excel(writer, sheet_name='Source', index=False)
                    target_file.to_excel(writer, sheet_name='Target', index=False)
                    pd.DataFrame(columns=source_file.columns).to_excel(writer, sheet_name='Differences', index=False)
                    writer.sheets['Differences'].write(0, 0, "No difference found between source and target files.")
                
                # Display a message and provide a download button for the file
                st.write("No differences found.")
                with open(diff_file_path, "rb") as f:
                    st.download_button("Download Differences File", f, file_name="diff_output.xlsx")

        except Exception as e:
            st.error(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
