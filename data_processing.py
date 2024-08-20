import pandas as pd
from openpyxl import load_workbook

def check_sheet_exists(book, sheet_name):
    """Check if a sheet exists in the Excel file."""
    if sheet_name not in book.sheetnames:
        raise ValueError(f"The sheet '{sheet_name}' is not present in the Excel file.")

def insert_data_into_sheet(sheet, data, start_row, dest_column):
    """Inserts data into the specified sheet starting from the given row and column."""
    for i, value in enumerate(data):
        row = start_row + i
        sheet.cell(row=row, column=dest_column, value=value)

def process_data(csv_file, excel_file, column_vars, column_dest_vars, start_row_var, sheet_name_var, dropdown_option, dropdown_cell_var, attachment_var, attachment_column_var, attachment_row_var):
    df = pd.read_csv(csv_file)
    book = load_workbook(excel_file)

    check_sheet_exists(book, sheet_name_var.get())

    sheet = book[sheet_name_var.get()]
    sheet[dropdown_cell_var.get()] = dropdown_option

    for column, var in column_vars.items():
        if var.get():
            dest_column = column_dest_vars[column].get().strip().upper()
            dest_col_idx = ord(dest_column) - ord('A') + 1
            insert_data_into_sheet(sheet, df[column].tolist(), start_row_var.get(), dest_col_idx)

    if attachment_var.get() and attachment_column_var.get() and attachment_row_var.get() > 0:
        dest_col_idx = ord(attachment_column_var.get().strip().upper()) - ord('A') + 1
        sheet.cell(row=attachment_row_var.get(), column=dest_col_idx, value=attachment_var.get())

    book.save(excel_file)
