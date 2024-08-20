import pandas as pd
from openpyxl import load_workbook

def check_nuget_sheet(excel_file):
    """Checks if the specified sheet exists in the Excel file."""
    book = load_workbook(excel_file, read_only=True)
    if "NuGet" not in book.sheetnames:
        raise ValueError(f"The sheet 'NuGet' is not present in {excel_file}.")

def process_data(csv_file, excel_file, column_vars, column_dest_vars, start_row_var, sheet_name_var, dropdown_option, dropdown_cell_var):
    """Processes the CSV file and appends the selected columns to the specified sheet in an Excel file."""
    df = pd.read_csv(csv_file)

    # Load the Excel file
    book = load_workbook(excel_file)
    
    # Check if the specified sheet exists
    if sheet_name_var.get() not in book.sheetnames:
        raise ValueError(f"The sheet '{sheet_name_var.get()}' is not present in {excel_file}.")
    
    sheet = book[sheet_name_var.get()]

    # Write the dropdown option to the specified cell
    sheet[dropdown_cell_var.get()] = dropdown_option

    # Process and copy selected columns to the specified destination columns
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
                # Ensure row index is at least the specified start row
                if (start_row_var.get() + i) < 1:
                    raise ValueError(f"Invalid row index {start_row_var.get() + i} for column '{column}'.")
                sheet.cell(row=start_row_var.get() + i, column=dest_col_idx, value=value)

    # Save the updated Excel file
    book.save(excel_file)
