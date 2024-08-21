import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import streamlit as st
import tempfile

def extract_and_merge_columns(old_file, new_file, old_keys, new_keys, old_cols, new_cols):
    # Concatenate primary keys to create unique identifiers
    old_file['UniqueKey'] = old_file[old_keys].astype(str).agg('_'.join, axis=1)
    new_file['UniqueKey'] = new_file[new_keys].astype(str).agg('_'.join, axis=1)

    # Extract specified columns from old and new files
    old_data = old_file[['UniqueKey'] + old_cols]
    new_data = new_file[['UniqueKey'] + new_cols]

    # Perform a VLOOKUP-like merge based on the unique identifiers
    merged_data = pd.merge(old_data, new_data, on='UniqueKey', how='left')

    # Drop duplicate rows based on the UniqueKey column
    merged_data = merged_data.drop_duplicates(subset='UniqueKey')

    # Create a temporary file to save the merged data
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
        merged_file_path = temp_file.name

    with pd.ExcelWriter(merged_file_path, engine='openpyxl') as writer:
        merged_data.to_excel(writer, sheet_name='Merged', index=False)

    return merged_file_path

def highlight_col(file_path):
    wb = load_workbook(file_path)
    ws_diff = wb['Differences']

    # Define fill colors
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    pink_fill = PatternFill(start_color='FFC0CB', end_color='FFC0CB', fill_type='solid')
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    # Iterate over rows in the 'Differences' sheet
    for row in ws_diff.iter_rows(min_row=2, max_col=ws_diff.max_column):
        merge_output = row[-1].value  # Merge status assumed to be in the last column

        # Apply row-level highlighting based on merge status
        for cell in row:
            if merge_output == 'Source Only':
                cell.fill = yellow_fill
            elif merge_output == 'Target Only':
                cell.fill = pink_fill

    # Apply red highlight to cells with differences
    num_cols = len(ws_diff.columns) - 1  # Number of columns excluding the last one for merge status

    for row in ws_diff.iter_rows(min_row=2, max_col=num_cols):
        merge_output = row[-1].value  # Merge status assumed to be in the last column
        if merge_output in ['Source Only', 'Target Only']:
            for i in range(num_cols // 2):  # Adjust this based on how columns are split
                source_cell = row[i]
                target_cell = row[i + num_cols // 2]  # Compare with corresponding cell in the other half

                if source_cell.value != target_cell.value:
                    source_cell.fill = red_fill
                    target_cell.fill = red_fill

    wb.save(file_path)

def main():
    st.title('Excel Compare & Merge Tool (v3)')

    # Upload widgets for the old and new Excel files for column extraction
    old_file_upload = st.file_uploader("Upload Excel A for merging", type=['xlsx'], key='old')
    new_file_upload = st.file_uploader("Upload Excel B for merging", type=['xlsx'], key='new')

    # Specify the primary keys and columns to extract
    old_keys = st.text_input("Unique Key - Enter Primary Key Column names for Excel A (separated by commas & no space)", "PrimaryKeyA1,PrimaryKeyA2")
    new_keys = st.text_input("Unique Key - Enter Primary Key Column names for Excel B (separated by commas & no space)", "PrimaryKeyB1,PrimaryKeyB2")
    old_columns = st.text_input("Enter Column names (include primary keys) to extract from Excel A (separated by commas & no space)", "Column1,Column2..etc")
    new_columns = st.text_input("Enter Column names to extract from Excel B (separated by commas & no space)", "Column1,Column2..etc")

    if st.button("Extract and Merge Columns"):
        if old_file_upload and new_file_upload:
            try:
                old_file = pd.read_excel(old_file_upload)
                new_file = pd.read_excel(new_file_upload)

                # Convert input columns to lists
                old_keys_list = [key.strip() for key in old_keys.split(',')]
                new_keys_list = [key.strip() for key in new_keys.split(',')]
                old_cols = [col.strip() for col in old_columns.split(',')]
                new_cols = [col.strip() for col in new_columns.split(',')]

                merged_file_path = extract_and_merge_columns(old_file, new_file, old_keys_list, new_keys_list, old_cols, new_cols)

                st.write("Merged columns have been successfully extracted and merged.")
                with open(merged_file_path, "rb") as f:
                    st.download_button("Download Merged File", f, file_name="merged_output.xlsx")

                # Clean up temporary file
                os.remove(merged_file_path)

            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Please upload both Excel files.")

    # File upload widgets for the source and target files for comparison
    uploaded_source_file = st.file_uploader("Upload SOURCE File for comparison (both files must be same format & same no. of columns)", type=['xlsx'], key='source')
    uploaded_target_file = st.file_uploader("Upload TARGET File for comparison (both files must be same format & same no. of columns)", type=['xlsx'], key='target')

    if st.button("Compare Excel Files"):
        if uploaded_source_file and uploaded_target_file:
            try:
                source_file = pd.read_excel(uploaded_source_file)
                target_file = pd.read_excel(uploaded_target_file)

                source_file = source_file.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)
                target_file = target_file.applymap(lambda x: str(x).strip() if isinstance(x, str) else x)

                diff = source_file.merge(target_file, indicator=True, how='outer')
                diff['_merge'].replace({'left_only': 'Source Only', 'right_only': 'Target Only'}, inplace=True)
                diff.rename(columns={'_merge': 'Merge_Output'}, inplace=True)
                diff = diff[diff['Merge_Output'] != 'both']

                diff_file_path = 'diff_output.xlsx'

                if not diff.empty:
                    with pd.ExcelWriter(diff_file_path) as writer:
                        source_file.to_excel(writer, sheet_name='Source', index=False)
                        target_file.to_excel(writer, sheet_name='Target', index=False)
                        diff.to_excel(writer, sheet_name='Differences', index=False)
                    
                    highlight_col(diff_file_path)

                    st.write(f"Found {len(diff)} differences.")
                    with open(diff_file_path, "rb") as f:
                        st.download_button("Download Differences File", f, file_name="diff_output.xlsx")
                else:
                    with pd.ExcelWriter(diff_file_path) as writer:
                        source_file.to_excel(writer, sheet_name='Source', index=False)
                        target_file.to_excel(writer, sheet_name='Target', index=False)
                        pd.DataFrame(columns=source_file.columns).to_excel(writer, sheet_name='Differences', index=False)
                        writer.sheets['Differences'].write(0, 0, "No difference found between source and target files.")
                    
                    st.write("No differences found.")
                    with open(diff_file_path, "rb") as f:
                        st.download_button("Download Differences File", f, file_name="diff_output.xlsx")

            except Exception as e:
                st.error(f"An error occurred: {e}")
        else:
            st.error("Please upload both Excel files.")

if __name__ == "__main__":
    main()
