import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import streamlit as st
from pathlib import Path

def highlight_col(file_path):
    wb = load_workbook(file_path)
    ws_diff = wb['Differences']

    yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')
    pink_fill = PatternFill(start_color='FFC0CB', end_color='FFC0CB', fill_type='solid')

    for row in ws_diff.iter_rows(min_row=2, max_col=ws_diff.max_column):
        merge_output = row[-1].value
        for cell in row:
            if merge_output == 'Source Only':
                cell.fill = yellow_fill
            elif merge_output == 'Target Only':
                cell.fill = pink_fill

    wb.save(file_path)

def main():
    st.title('Logistics Data Comparison Tool')

    uploaded_source_file = st.file_uploader("Upload Source File", type=['xlsx'])
    uploaded_target_file = st.file_uploader("Upload Target File", type=['xlsx'])

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

if __name__ == "__main__":
    main()
