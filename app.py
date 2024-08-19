import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

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

def start_procedure():
    """Function called by the button to start the data processing."""
    csv_file = file_input_var.get()
    excel_file = file_output_var.get()
    
    if not csv_file:
        messagebox.showwarning("No File Selected", "No CSV file selected. The operation has been canceled.")
        return

    if not excel_file:
        messagebox.showwarning("No File Selected", "No Excel file selected. The operation has been canceled.")
        return

    try:
        # Read the CSV file
        df = pd.read_csv(csv_file)

        # Get selected columns
        selected_columns = [col for col, var in column_vars.items() if var.get()]
        if not selected_columns:
            messagebox.showwarning("No Columns Selected", "No columns selected. The operation has been canceled.")
            return

        # Select the columns to copy
        df_selected = df[selected_columns]

        # Write the selected dataframe to an Excel file
        df_selected.to_excel(excel_file, index=False, engine='openpyxl')

        messagebox.showinfo("Success", f"Selected columns have been copied to {excel_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def choose_input_file():
    """Updates the input field with the path of the selected file."""
    file_path = select_input_file()
    if file_path:
        file_input_var.set(file_path)
        update_column_checkboxes(file_path)

def choose_output_file():
    """Updates the output field with the path of the selected file."""
    file_path = select_output_file()
    if file_path:
        file_output_var.set(file_path)

def update_column_checkboxes(file_path):
    """Updates the list of column checkboxes based on the CSV file."""
    global column_vars

    # Clear existing checkboxes
    for widget in column_frame.winfo_children():
        widget.destroy()

    try:
        df = pd.read_csv(file_path)
        columns = df.columns.tolist()
        
        column_vars = {}
        for column in columns:
            var = tk.BooleanVar()
            column_vars[column] = var
            tk.Checkbutton(column_frame, text=column, variable=var).pack(anchor='w')
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while reading the CSV file: {e}")

def create_interface():
    """Creates the graphical user interface."""
    # Create the main window
    root = tk.Tk()
    root.title("CSV to Excel Application")

    # Variables to store file paths
    global file_input_var, file_output_var, column_vars
    file_input_var = tk.StringVar()
    file_output_var = tk.StringVar()
    column_vars = {}

    # Create GUI elements
    tk.Label(root, text="Input CSV File:").pack(pady=5)
    tk.Entry(root, textvariable=file_input_var, width=50).pack(pady=5)
    tk.Button(root, text="Select CSV File", command=choose_input_file).pack(pady=5)

    tk.Label(root, text="Output Excel File:").pack(pady=5)
    tk.Entry(root, textvariable=file_output_var, width=50).pack(pady=5)
    tk.Button(root, text="Select Excel File", command=choose_output_file).pack(pady=5)

    tk.Label(root, text="Select Columns to Copy:").pack(pady=5)
    
    global column_frame
    column_frame = tk.Frame(root)
    column_frame.pack(pady=5)

    tk.Button(root, text="Start Procedure", command=start_procedure).pack(pady=20)

    # Run the GUI
    root.mainloop()

if __name__ == "__main__":
    create_interface()
