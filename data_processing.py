import pandas as pd
from openpyxl import load_workbook

def check_nuget_sheet(excel_file):
    """Checks if the 'NuGet' sheet exists in the Excel file."""
    book = load_workbook(excel_file, read_only=True)
    if "NuGet" not in book.sheetnames:
        raise ValueError(f"The sheet 'NuGet' is not present in {excel_file}.")


def process_data(csv_file, excel_file, column_vars, column_dest_vars):
    """Processes the CSV file and appends the selected columns to the NuGet sheet in an Excel file."""
    df = pd.read_csv(csv_file)

    # Load the Excel file
    book = load_workbook(excel_file)
    
    if "NuGet" not in book.sheetnames:
        raise ValueError(f"The sheet 'NuGet' is not present in {excel_file}.")
    
    sheet = book["NuGet"]

    for column, var in column_vars.items():
        if var.get():
            dest_column = column_dest_vars[column].get().strip().upper()
            if not dest_column:
                raise ValueError(f"Destination column not specified for '{column}'.")
            
            # Convert column letter to index (A=1, B=2, etc.)
            if len(dest_column) != 1 or not ('A' <= dest_column <= 'Z'):
                raise ValueError(f"Invalid column letter '{dest_column}' specified for '{column}'.")

            dest_col_idx = ord(dest_column) - ord('A') + 1
            
            # Ensure destination column index is at least 1
            if dest_col_idx < 1:
                raise ValueError(f"Destination column index {dest_col_idx} is invalid for column '{column}'.")

            col_data = df[column].tolist()
            for i, value in enumerate(col_data):
                # Ensure row index is at least 10
                if (10 + i) < 1:
                    raise ValueError(f"Invalid row index {10 + i} for column '{column}'.")
                sheet.cell(row=10 + i, column=dest_col_idx, value=value)

    # Save the updated Excel file
    book.save(excel_file)





