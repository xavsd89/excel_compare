import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import streamlit as st
import tempfile

def extract_and_merge_columns(old_file, new_file, old_keys, new_keys, old_cols, new_cols):
    """
    Extracts specified columns from two Excel files and merges them based on unique keys.
    """

    # Create a unique key in the old file by concatenating values of specified columns
    old_file['UniqueKey'] = old_file[old_keys].astype(str).agg('_'.join, axis=1)
    # Create a unique key in the new file by concatenating values of specified columns
    new_file['UniqueKey'] = new_file[new_keys].astype(str).agg('_'.join, axis=1)

    # Extract only the unique key and specified columns from the old file
    old_data = old_file[['UniqueKey'] + old_cols]
    # Extract only the unique key and specified columns from the new file
    new_data = new_file[['UniqueKey'] + new_cols]

    # Merge the old and new data on the unique key, keeping only matching rows
    merged_data = pd.merge(old_data, new_data, on='UniqueKey', how='inner')

    # Remove any duplicate rows based on the unique key
    merged_data = merged_data.drop_duplicates(subset='UniqueKey')

    # Create a temporary file to save the merged data
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        merged_file_path = temp_file.name

    # Write the merged data to the temporary Excel file
    with pd.ExcelWriter(merged_file_path, engine='openpyxl') as writer:
        merged_data.to_excel(writer, sheet_name='Merged', index=False)

    return merged_file_path

def highlight_col(file_path):
    """
    Highlights differences in an Excel file based on merge status.
    """

    # Load the workbook from the given file path
    wb = load_workbook(file_path)
    # Access the 'Differences' sheet
    ws_diff = wb['Differences']

    # Define fill colors for highlighting cells
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    pink_fill = PatternFill(start_color='FFC0CB', end_color='FFC0CB', fill_type='solid')
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    # Iterate over rows in the 'Differences' sheet, starting from the second row
    for row in ws_diff.iter_rows(min_row=2, max_col=ws_diff.max_column):
        # Get the merge status from the last column of the row
        merge_output = row[-1].value

        # Highlight the entire row based on the merge status
        for cell in row:
            if merge_output == 'Source Only':
                cell.fill = yellow_fill
            elif merge_output == 'Target Only':
                cell.fill = pink_fill

    # Calculate the number of columns excluding the last column (merge status)
    num_cols = len(ws_diff.columns) - 1

    # Iterate over rows to compare cell values
    for row in ws_diff.iter_rows(min_row=2, max_col=num_cols):
        # Get the merge status from the last column of the row
        merge_output = row[-1].value
        if merge_output in ['Source Only', 'Target Only']:
            # Compare cells from two halves of the row
            for i in range(num_cols // 2):
                source_cell = row[i]
                target_cell = row[i + num_cols // 2]

                # Highlight cells in red if they are different
                if source_cell.value != target_cell.value:
                    source_cell.fill = red_fill
                    target_cell.fill = red_fill

    # Save the workbook with the applied highlights
    wb.save(file_path)

def main():
    """
    Main function to run the Streamlit application for comparing and merging Excel files.
    """
    # Set the title of the Streamlit app
    st.title('Excel Compare & Merge Tool (v3)')

    # Create file uploader widgets for old and new Excel files for merging
    old_file_upload = st.file_uploader("Upload Excel A for merging (both files must have the same primary keys)", type=['xlsx'], key='old')
    new_file_upload = st.file_uploader("Upload Excel B for merging (both files must have the same primary keys)", type=['xlsx'], key='new')

    # Input fields for primary key columns and columns to extract
    old_keys = st.text_input("Unique Key - Enter Primary Key Column names for Excel A (separated by commas & no space)", "PrimaryKeyA1,PrimaryKeyA2")
    new_keys = st.text_input("Unique Key - Enter Primary Key Column names for Excel B (separated by commas & no space)", "PrimaryKeyB1,PrimaryKeyB2")
    old_columns = st.text_input("Enter Column names (include primary keys) to extract from Excel A (separated by commas & no space)", "Column1,Column2..etc")
    new_columns = st.text_input("Enter Column names to extract from Excel B (separated by commas & no space)", "Column1,Column2..etc")

    # Check if the "Extract and Merge Columns" button is clicked
    if st.button("Extract and Merge Columns"):
        if old_file_upload and new_file_upload:
            try:
                # Read the uploaded Excel files into DataFrames
                old_file = pd.read_excel(old_file_upload)
                new_file = pd.read_excel(new_file_upload)

                # Convert input columns to lists
                old_keys_list = [key.strip() for key in old_keys.split(',')]
                new_keys_list = [key.strip() for key in new_keys.split(',')]
                old_cols = [col.strip() for col in old_columns.split(',')]
                new_cols = [col.strip() for col in new_columns.split(',')]

                # Extract and merge columns from the old and new files
                merged_file_path = extract_and_merge_columns(old_file, new_file, old_keys_list, new_keys_list, old_cols, new_cols)

                # Notify the user of successful merge and provide download link
                st.write("Merged columns have been successfully extracted and merged.")
                with open(merged_file_path, "rb") as f:
                    st.download_button("Download Merged File", f, file_name="merged_output.xlsx")

                # Clean up temporary file
                os.remove(merged_file_path)

            except Exception as e:
                # Display error message if an exception occurs
                st.error(f"An error occurred: {e}")
        else:
            # Display error message if both files are not uploaded
            st.error("Please upload both Excel files.")

    # Create file uploader widgets for source and target files for comparison
    uploaded_source_file = st.file_uploader("Upload SOURCE File for comparison (both files must be same format & same no. of columns)", type=['xlsx'], key='source')
    uploaded_target_file = st.file_uploader("Upload TARGET File for comparison (both files must be same format & same no. of columns)", type=['xlsx'], key='target')

    # Check if the "Compare Excel Files" button is clicked
    if st.button("Compare Excel Files"):
        if uploaded_source_file and uploaded_target_file:
            try:
                # Read the uploaded Excel files into DataFrames
                source_file = pd.read_excel(uploaded_source_file)
                target_file = pd.read_excel(uploaded_target_file)

                # Strip whitespace from string values in both DataFrames
                source_file = source_file.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)
                target_file = target_file.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)

                # Perform an outer merge to find differences between source and target files
                diff = source_file.merge(target_file, indicator=True, how='outer')
                # Replace merge indicator values with custom labels
                diff['_merge'].replace({'left_only': 'Source Only', 'right_only': 'Target Only'}, inplace=True)
                # Rename the merge indicator column to 'Merge_Output'
                diff.rename(columns={'_merge': 'Merge_Output'}, inplace=True)
                # Filter out rows where both source and target have the same data
                diff = diff[diff['Merge_Output'] != 'both']

                # Define the path for the differences file
                diff_file_path = 'diff_output.xlsx'

                if not diff.empty:
                    # Write source, target, and differences to the Excel file
                    with pd.ExcelWriter(diff_file_path) as writer:
                        source_file.to_excel(writer, sheet_name='Source', index=False)
                        target_file.to_excel(writer, sheet_name='Target', index=False)
                        diff.to_excel(writer, sheet_name='Differences', index=False)
                    
                    # Apply highlighting to the differences sheet
                    highlight_col(diff_file_path)

                    # Notify the user of found differences and provide download link
                    st.write(f"Found {len(diff)} differences.")
                    with open(diff_file_path, "rb") as f:
                        st.download_button("Download Differences File", f, file_name="diff_output.xlsx")
                else:
                    # Create an Excel file with no differences
                    with pd.ExcelWriter(diff_file_path) as writer:
                        source_file.to_excel(writer, sheet_name='Source', index=False)
                        target_file.to_excel(writer, sheet_name='Target', index=False)
                        pd.DataFrame(columns=source_file.columns).to_excel(writer, sheet_name='Differences', index=False)
                        writer.sheets['Differences'].write(0, 0, "No difference found between source and target files.")
                    
                    # Notify the user that no differences were found and provide download link
                    st.write("No differences found.")
                    with open(diff_file_path, "rb") as f:
                        st.download_button("Download Differences File", f, file_name="diff_output.xlsx")

            except Exception as e:
                # Display error message if an exception occurs
                st.error(f"An error occurred: {e}")
        else:
            # Display error message if both files are not uploaded
            st.error("Please upload both Excel files.")

if __name__ == "__main__":
    # Run the main function to start the Streamlit application
    main()
