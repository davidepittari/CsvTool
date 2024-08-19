from tkinter import filedialog

def select_input_file():
    """Opens a dialog to select the input CSV file."""
    file_path = filedialog.askopenfilename(
        filetypes=[("CSV Files", "*.csv")],
        title="Select Input CSV File"
    )
    return file_path

def select_output_file():
    """Opens a dialog to select the output Excel file."""
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Select Output Excel File"
    )
    return file_path
