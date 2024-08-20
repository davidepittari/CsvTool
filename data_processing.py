import pandas as pd
from openpyxl import load_workbook

def check_nuget_sheet(excel_file):
    """Checks if the specified sheet exists in the Excel file."""
    book = load_workbook(excel_file, read_only=True)
    if "NuGet" not in book.sheetnames:
        raise ValueError(f"The sheet 'NuGet' is not present in {excel_file}.")

def process_data(csv_file, excel_file, column_vars, column_dest_vars, start_row_var, sheet_name_var, dropdown_option, dropdown_cell_var, attachment_var, attachment_column_var, attachment_row_var):
    """Processes the CSV file and appends the selected columns to the specified sheet in an Excel file."""
    df = pd.read_csv(csv_file)

    book = load_workbook(excel_file)

    if sheet_name_var.get() not in book.sheetnames:
        raise ValueError(f"The sheet '{sheet_name_var.get()}' is not present in {excel_file}.")

    sheet = book[sheet_name_var.get()]

    sheet[dropdown_cell_var.get()] = dropdown_option

    for column, var in column_vars.items():
        if var.get():
            dest_column = column_dest_vars[column].get().strip().upper()
            if not dest_column:
                raise ValueError(f"Destination column not specified for '{column}'.")

            if len(dest_column) != 1 or not ('A' <= dest_column <= 'Z'):
                raise ValueError(f"Invalid column letter '{dest_column}' specified for '{column}'.")

            dest_col_idx = ord(dest_column) - ord('A') + 1

            if dest_col_idx < 1:
                raise ValueError(f"Destination column index {dest_col_idx} is invalid for column '{column}'.")

            col_data = df[column].tolist()
            for i, value in enumerate(col_data):
                if (start_row_var.get() + i) < 1:
                    raise ValueError(f"Invalid row index {start_row_var.get() + i} for column '{column}'.")
                sheet.cell(row=start_row_var.get() + i, column=dest_col_idx, value=value)

    # Aggiungere l'allegato alla riga e colonna specificata
    if attachment_var.get() and attachment_column_var.get() and attachment_row_var.get() > 0:
        dest_col_idx = ord(attachment_column_var.get().strip().upper()) - ord('A') + 1
        sheet.cell(row=attachment_row_var.get(), column=dest_col_idx, value=attachment_var.get())

    book.save(excel_file)


