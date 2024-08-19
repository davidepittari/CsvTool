import tkinter as tk
from tkinter import messagebox
import pandas as pd
from file_utils import select_input_file, select_output_file
from data_processing import check_nuget_sheet, process_data

def choose_input_file(file_input_var, column_frame, column_vars, column_dest_vars):
    """Updates the input field with the path of the selected file."""
    file_path = select_input_file()
    if file_path:
        file_input_var.set(file_path)
        update_column_checkboxes(file_path, column_frame, column_vars, column_dest_vars)

def choose_output_file(file_output_var):
    """Updates the output field with the path of the selected file."""
    file_path = select_output_file()
    if file_path:
        file_output_var.set(file_path)

def update_column_checkboxes(file_path, column_frame, column_vars, column_dest_vars):
    """Updates the list of column checkboxes and destination column entries based on the CSV file."""
    for widget in column_frame.winfo_children():
        widget.destroy()

    try:
        df = pd.read_csv(file_path)
        columns = df.columns.tolist()
        
        column_vars.clear()
        column_dest_vars.clear()
        for column in columns:
            var = tk.BooleanVar()
            column_vars[column] = var
            
            frame = tk.Frame(column_frame)
            frame.pack(fill='x', pady=2)
            
            tk.Checkbutton(frame, text=column, variable=var).pack(side='left')
            tk.Label(frame, text=" -> Column in NuGet:").pack(side='left')
            
            dest_var = tk.StringVar()
            column_dest_vars[column] = dest_var
            tk.Entry(frame, textvariable=dest_var, width=5).pack(side='left')
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred while reading the CSV file: {e}")


def start_procedure(file_input_var, file_output_var, column_vars, column_dest_vars):
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
        check_nuget_sheet(excel_file)
        process_data(csv_file, excel_file, column_vars, column_dest_vars)
        messagebox.showinfo("Success", f"Selected columns have been copied to the NuGet sheet in {excel_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def create_interface(root):
    """Creates the graphical user interface."""
    main_frame = tk.Frame(root)
    main_frame.pack(fill='both', expand=True)

    canvas = tk.Canvas(main_frame)
    canvas.pack(side='left', fill='both', expand=True)

    scrollbar = tk.Scrollbar(main_frame, orient='vertical', command=canvas.yview)
    scrollbar.pack(side='right', fill='y')

    content_frame = tk.Frame(canvas)
    canvas.create_window((0, 0), window=content_frame, anchor='nw')
    content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox('all')))

    file_input_var = tk.StringVar()
    file_output_var = tk.StringVar()
    column_vars = {}
    column_dest_vars = {}  # New dictionary for destination columns

    tk.Label(content_frame, text="Input CSV File:").pack(pady=5)
    tk.Entry(content_frame, textvariable=file_input_var, width=50).pack(pady=5)
    tk.Button(content_frame, text="Select CSV File", command=lambda: choose_input_file(file_input_var, column_frame, column_vars, column_dest_vars)).pack(pady=5)

    tk.Label(content_frame, text="Output Excel File:").pack(pady=5)
    tk.Entry(content_frame, textvariable=file_output_var, width=50).pack(pady=5)
    tk.Button(content_frame, text="Select Excel File", command=lambda: choose_output_file(file_output_var)).pack(pady=5)

    tk.Label(content_frame, text="Select Columns to Copy:").pack(pady=5)
    
    column_frame = tk.Frame(content_frame)
    column_frame.pack(pady=5, fill='x')

    tk.Button(content_frame, text="Start Procedure", command=lambda: start_procedure(file_input_var, file_output_var, column_vars, column_dest_vars)).pack(pady=20)

    canvas.config(yscrollcommand=scrollbar.set)
