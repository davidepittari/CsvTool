import tkinter as tk
from tkinter import messagebox
import pandas as pd
from file_utils import select_input_file, select_output_file
from data_processing import process_data
from tkinter import Menu

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

def open_options_window(start_row_var, sheet_name_var):
    """Opens a new window to set options like start row and sheet name."""
    options_window = tk.Toplevel()
    options_window.title("Opzioni")
    
    # Create a frame to hold all the content
    content_frame = tk.Frame(options_window, padx=20, pady=20)
    content_frame.pack(fill='both', expand=True)

    tk.Label(content_frame, text="Riga di partenza:").grid(row=0, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=start_row_var).grid(row=0, column=1, pady=5)

    tk.Label(content_frame, text="Nome della scheda:").grid(row=1, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=sheet_name_var).grid(row=1, column=1, pady=5)

    tk.Button(content_frame, text="Salva", command=options_window.destroy).grid(row=2, column=0, columnspan=2, pady=10)

    # Configure column weights for proper centering
    content_frame.grid_columnconfigure(0, weight=1)
    content_frame.grid_columnconfigure(1, weight=1)

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
            tk.Label(frame, text=" -> Colonna in NuGet:").pack(side='left')
            
            dest_var = tk.StringVar()
            column_dest_vars[column] = dest_var
            tk.Entry(frame, textvariable=dest_var, width=5).pack(side='left')
    except Exception as e:
        messagebox.showerror("Errore", f"Errore durante la lettura del file CSV: {e}")

def start_procedure(file_input_var, file_output_var, column_vars, column_dest_vars, start_row_var, sheet_name_var):
    """Function called by the button to start the data processing."""
    csv_file = file_input_var.get()
    excel_file = file_output_var.get()
    
    if not csv_file:
        messagebox.showwarning("Nessun file selezionato", "Non è stato selezionato alcun file CSV. Operazione annullata.")
        return

    if not excel_file:
        messagebox.showwarning("Nessun file selezionato", "Non è stato selezionato alcun file Excel. Operazione annullata.")
        return

    try:
        process_data(csv_file, excel_file, column_vars, column_dest_vars, start_row_var.get(), sheet_name_var.get())
        messagebox.showinfo("Successo", f"Le colonne selezionate sono state copiate nel file {excel_file}")
    except Exception as e:
        messagebox.showerror("Errore", f"Si è verificato un errore: {e}")


def create_interface(root):
    """Creates the graphical user interface."""

    # Variabili per la configurazione delle opzioni
    start_row_var = tk.IntVar(value=10)  # Default start row
    sheet_name_var = tk.StringVar(value="NuGet")  # Default sheet name

    # Menu Bar
    menubar = Menu(root)
    root.config(menu=menubar)
    
    # Create a File Menu and add items
    file_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Esci", command=root.quit)
    
    # Create an Options Menu
    options_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Opzioni", menu=options_menu)
    options_menu.add_command(label="Imposta Opzioni", command=lambda: open_options_window(start_row_var, sheet_name_var))
    
    # Create a Help Menu
    help_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Aiuto", menu=help_menu)
    help_menu.add_command(label="Informazioni", command=lambda: messagebox.showinfo("Informazioni", "CSV to Excel Application v1.0"))

    # Create content frame
    content_frame = tk.Frame(root)
    content_frame.pack(padx=20, pady=20, fill='both', expand=True)

    # GUI elements
    file_input_var = tk.StringVar()
    file_output_var = tk.StringVar()
    column_vars = {}
    column_dest_vars = {}

    tk.Label(content_frame, text="File CSV di input:").grid(row=0, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=file_input_var, width=50).grid(row=0, column=1, pady=5)
    tk.Button(content_frame, text="Seleziona file CSV", command=lambda: choose_input_file(file_input_var, column_frame, column_vars, column_dest_vars)).grid(row=0, column=2, pady=5)

    tk.Label(content_frame, text="File Excel di output:").grid(row=1, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=file_output_var, width=50).grid(row=1, column=1, pady=5)
    tk.Button(content_frame, text="Seleziona file Excel", command=lambda: choose_output_file(file_output_var)).grid(row=1, column=2, pady=5)

    tk.Label(content_frame, text="Seleziona le colonne da copiare:").grid(row=2, column=0, pady=5, sticky='e')

    column_frame = tk.Frame(content_frame)
    column_frame.grid(row=3, column=0, columnspan=3, pady=5, sticky='nsew')

    tk.Button(content_frame, text="Avvia procedura", command=lambda: start_procedure(file_input_var, file_output_var, column_vars, column_dest_vars, start_row_var, sheet_name_var)).grid(row=4, column=0, columnspan=3, pady=20)

    # Configure column and row weights for proper resizing
    content_frame.grid_columnconfigure(0, weight=1)
    content_frame.grid_columnconfigure(1, weight=1)
    content_frame.grid_columnconfigure(2, weight=1)
    content_frame.grid_rowconfigure(3, weight=1)
