"""
CSV processing functionality for the CSV to Excel converter.
"""
import pandas as pd

def get_csv_columns(csv_file):
    """
    Gets the column names from a CSV file.
    
    Args:
        csv_file (str): Path to the CSV file
        
    Returns:
        list: List of column names
    """
    try:
        df = pd.read_csv(csv_file)
        return df.columns.tolist()
    except Exception as e:
        raise Exception(f"Failed to read CSV file: {e}")

def process_csv_to_excel(csv_file, excel_file, selected_columns):
    """
    Processes a CSV file and creates an Excel file with selected columns.
    
    Args:
        csv_file (str): Path to the input CSV file
        excel_file (str): Path to the output Excel file
        selected_columns (list): List of column names to include
        
    Returns:
        None
    """
    try:
        # Read the CSV file
        df = pd.read_csv(csv_file)
        
        # Select the columns to copy
        df_selected = df[selected_columns]
        
        # Write the selected dataframe to an Excel file
        df_selected.to_excel(excel_file, index=False, engine='openpyxl')
        
    except Exception as e:
        raise Exception(f"Error processing CSV to Excel: {e}")