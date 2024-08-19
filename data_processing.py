import pandas as pd
from openpyxl import load_workbook

def check_nuget_sheet(excel_file):
    """Checks if the 'NuGet' sheet exists in the Excel file."""
    book = load_workbook(excel_file, read_only=True)
    if "NuGet" not in book.sheetnames:
        raise ValueError(f"The sheet 'NuGet' is not present in {excel_file}.")


def process_data(csv_file, excel_file, column_vars, start_row=1, start_col=1):
    """Processes the CSV file and appends the selected columns to the NuGet sheet in an Excel file."""
    df = pd.read_csv(csv_file)

    # Get selected columns
    selected_columns = [col for col, var in column_vars.items() if var.get()]
    if not selected_columns:
        raise ValueError("No columns selected.")

    # Load the Excel file
    book = load_workbook(excel_file)
    
    # Check if the "NuGet" sheet exists
    if "NuGet" not in book.sheetnames:
        raise ValueError(f"The sheet 'NuGet' is not present in {excel_file}.")
    
    # Select the NuGet sheet
    sheet = book["NuGet"]

    # Append the selected columns to the NuGet sheet
    for idx, column in enumerate(selected_columns):
        # Insert column data into the specified location
        col_data = df[column].tolist()
        for i, value in enumerate(col_data):
            sheet.cell(row=start_row + i, column=start_col + idx, value=value)

    # Save the updated Excel file
    book.save(excel_file)



