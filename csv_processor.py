"""
CSV processing functionality for the CSV to Excel converter.
"""
import pandas as pd
from pandas.errors import EmptyDataError

SUPPORTED_ENCODINGS = ['utf-8', 'latin1', 'iso-8859-1', 'cp1252', 'ascii']
COMMON_DELIMITERS = [',', ';', '\t', '|', ':']

def get_csv_columns(csv_file, delimiter=',', encoding='utf-8'):
    """
    Gets the column names from a CSV file.
    
    Args:
        csv_file (str): Path to the CSV file
        delimiter (str): CSV delimiter character
        encoding (str): File encoding
        
    Returns:
        list: List of column names
    """
    try:
        # Try reading with specified parameters
        df = pd.read_csv(csv_file, delimiter=delimiter, encoding=encoding, nrows=0)
        return df.columns.tolist()
    except UnicodeDecodeError:
        # Try common alternative encodings if default fails
        for enc in [e for e in SUPPORTED_ENCODINGS if e != encoding]:
            try:
                df = pd.read_csv(csv_file, delimiter=delimiter, encoding=enc, nrows=0)
                return df.columns.tolist()
            except:
                continue
        raise Exception("Failed to read CSV file. Please try a different encoding.")
    except EmptyDataError:
        raise Exception("The CSV file appears to be empty.")
    except pd.errors.ParserError:
        raise Exception("Could not parse CSV file. Please check the delimiter.")
    except Exception as e:
        raise Exception(f"Failed to read CSV file: {str(e)}")

def process_csv_to_excel(csv_file, excel_file, selected_columns, delimiter=',', encoding='utf-8'):
    """
    Processes a CSV file and creates an Excel file with selected columns.
    
    Args:
        csv_file (str): Path to the input CSV file
        excel_file (str): Path to the output Excel file
        selected_columns (list): List of column names to include
        delimiter (str): CSV delimiter character
        encoding (str): File encoding
        
    Returns:
        None
    """
    try:
        # Read the CSV file with error handling for encoding
        try:
            df = pd.read_csv(csv_file, delimiter=delimiter, encoding=encoding)
        except UnicodeDecodeError:
            # Try common alternative encodings if default fails
            for enc in [e for e in SUPPORTED_ENCODINGS if e != encoding]:
                try:
                    df = pd.read_csv(csv_file, delimiter=delimiter, encoding=enc)
                    break
                except:
                    continue
            else:
                raise Exception("Failed to read CSV file. Please try a different encoding.")
        
        # Verify all selected columns exist
        missing_cols = [col for col in selected_columns if col not in df.columns]
        if missing_cols:
            raise Exception(f"Columns not found in CSV: {', '.join(missing_cols)}")
        
        # Select the columns to copy
        df_selected = df[selected_columns]
        
        # Write to Excel with error handling
        try:
            df_selected.to_excel(excel_file, index=False, engine='openpyxl')
        except PermissionError:
            raise Exception("Permission denied. Please close the Excel file if it's open.")
        except Exception as e:
            raise Exception(f"Failed to write Excel file: {str(e)}")
        
    except pd.errors.ParserError:
        raise Exception("Could not parse CSV file. Please check the delimiter.")
    except Exception as e:
        raise Exception(f"Error processing CSV to Excel: {str(e)}")
