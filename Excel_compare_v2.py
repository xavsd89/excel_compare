import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import streamlit as st
from pathlib import Path

def highlight_col(file_path):
    """
    Highlights cells in the 'Differences' sheet of an Excel file based on their merge status.

    Parameters:
    - file_path (str): The path to the Excel file to be processed.
    """
    # Load the workbook and select the 'Differences' sheet
    wb = load_workbook(file_path)
    ws_diff = wb['Differences']

    # Define fill colors for highlighting
    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    pink_fill = PatternFill(start_color='FFC0CB', end_color='FFC0CB', fill_type='solid')

    # Iterate through each row in the 'Differences' sheet starting from the second row
    for row in ws_diff.iter_rows(min_row=2, max_col=ws_diff.max_column):
        merge_output = row[-1].value  # The merge status is assumed to be in the last column
        for cell in row:
            # Apply the appropriate fill color based on the merge status
            if merge_output == 'Source Only':
                cell.fill = yellow_fill
            elif merge_output == 'Target Only':
                cell.fill = pink_fill

    # Save the workbook with the highlighted cells
    wb.save(file_path)

def main():
    """
    Main function to run the Streamlit app for comparing two Excel files and highlighting differences.
    """
    st.title('Excel Comparison Tool')

    # File uploader widgets for source and target files
    uploaded_source_file = st.file_uploader("Upload Source File", type=['xlsx'])
    uploaded_target_file = st.file_uploader("Upload Target File", type=['xlsx'])

    # Check if both files are uploaded
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
            # Display an error message if an exception occurs
            st.error(f"An error occurred: {e}")

# Run the main function if the script is executed directly
if __name__ == "__main__":
    main()
