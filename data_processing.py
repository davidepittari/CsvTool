import pandas as pd

def process_data(csv_file, excel_file, column_vars):
    """Processes the CSV file and saves the selected columns to an Excel file."""
    df = pd.read_csv(csv_file)

    # Get selected columns
    selected_columns = [col for col, var in column_vars.items() if var.get()]
    if not selected_columns:
        raise ValueError("No columns selected.")

    # Select the columns to copy
    df_selected = df[selected_columns]

    # Write the selected dataframe to an Excel file
    df_selected.to_excel(excel_file, index=False, engine='openpyxl')
