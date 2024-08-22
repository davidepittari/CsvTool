import tkinter as tk
from tkinter import messagebox
import pandas as pd
from file_utils import select_final_file, select_input_file, select_output_file
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
        
def choose_finale_file(file_final_var):
    """Updates the output field with the path of the selected file."""
    file_path = select_final_file()
    if file_path:
        file_final_var.set(file_path)
        
def on_checkbox_toggle(checkbox_var):
    # Mostra lo stato della checkbox (selezionata o no)
    print(f"Checkbox is {'selected' if checkbox_var.get() else 'not selected'}")

def open_options_window(start_row_var, sheet_name_var, dropdown_options_var, update_dropdown, dropdown_cell_var):
    """Opens a new window to set options like start row, sheet name, and dropdown options."""
    options_window = tk.Toplevel()
    options_window.title("Opzioni")
    
    # Create a frame to hold all the content
    content_frame = tk.Frame(options_window, padx=20, pady=20)
    content_frame.pack(fill='both', expand=True)

    tk.Label(content_frame, text="Riga di partenza:").grid(row=0, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=start_row_var).grid(row=0, column=1, pady=5)
    
    tk.Label(content_frame, text="Cella da scrivere:").grid(row=1, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=dropdown_cell_var).grid(row=1, column=1, pady=5)

    tk.Label(content_frame, text="Nome della scheda:").grid(row=2, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=sheet_name_var).grid(row=2, column=1, pady=5)

    tk.Label(content_frame, text="Opzioni Dropdown (separate da virgola):").grid(row=3, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=dropdown_options_var, width=30).grid(row=3, column=1, pady=5)

    def save_and_update():
        update_dropdown()
        options_window.destroy()

    tk.Button(content_frame, text="Salva", command=save_and_update).grid(row=4, column=0, columnspan=2, pady=10)

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
        
        # Creazione di un Canvas e un Frame all'interno per gestire lo scrolling
        canvas = tk.Canvas(column_frame)
        scrollbar = tk.Scrollbar(column_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Posizionamento del Canvas e della Scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        for column in columns:
            var = tk.BooleanVar()
            column_vars[column] = var
            
            frame = tk.Frame(scrollable_frame)
            frame.pack(fill='x', pady=2)
            
            tk.Checkbutton(frame, text=column, variable=var).pack(side='left')
            tk.Label(frame, text=" -> Colonna in NuGet:").pack(side='left')
            
            dest_var = tk.StringVar()
            column_dest_vars[column] = dest_var
            tk.Entry(frame, textvariable=dest_var, width=5).pack(side='left')
    except Exception as e:
        messagebox.showerror("Errore", f"Errore durante la lettura del file CSV: {e}")

def start_procedure(file_input_var, file_output_var, file_final_var, column_vars, column_dest_vars, start_row_var, sheet_name_var, selected_option_var, dropdown_cell_var, is_final):
    """Function called by the button to start the data processing."""
    csv_file = file_input_var.get()
    excel_file = file_output_var.get()
    excel_file_final = file_final_var.get()
    dropdown_option = selected_option_var.get()  # Get the selected option from the dropdown
    
    if not csv_file:
        messagebox.showwarning("No File Selected", "No CSV file selected. The operation has been canceled.")
        return

    if not excel_file:
        messagebox.showwarning("No File Selected", "No Excel file selected. The operation has been canceled.")
        return

    try:
        process_data(csv_file, excel_file, excel_file_final, column_vars, column_dest_vars, start_row_var, sheet_name_var, dropdown_option, dropdown_cell_var, is_final)
        messagebox.showinfo("Success", f"Selected columns have been copied to {excel_file}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def create_interface(root):
    """Creates the graphical user interface."""

    # Variabili per la configurazione delle opzioni
    start_row_var = tk.IntVar(value=10)
    sheet_name_var = tk.StringVar(value="NuGet")
    dropdown_options_var = tk.StringVar(value="Test, Test")
    selected_option_var = tk.StringVar()
    dropdown_cell_var = tk.StringVar(value="A9")
    
    def update_dropdown():
        menu = dropdown["menu"]
        menu.delete(0, "end")
        new_options = dropdown_options_var.get().split(", ")
        for option in new_options:
            menu.add_command(label=option, command=lambda value=option: selected_option_var.set(value))
        selected_option_var.set(new_options[0])

    menubar = Menu(root)
    root.config(menu=menubar)

    file_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Esci", command=root.quit)

    options_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Opzioni", menu=options_menu)
    options_menu.add_command(label="Imposta Opzioni", command=lambda: open_options_window(start_row_var, sheet_name_var, dropdown_options_var, update_dropdown, dropdown_cell_var))

    help_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Aiuto", menu=help_menu)
    help_menu.add_command(label="Informazioni", command=lambda: messagebox.showinfo("Informazioni", "CSV to Excel Application v1.0"))

    content_frame = tk.Frame(root)
    content_frame.pack(padx=20, pady=20, fill='both', expand=True)

    file_input_var = tk.StringVar()
    file_output_var = tk.StringVar()
    file_final_var = tk.StringVar()
    column_vars = {}
    column_dest_vars = {}
    checkbox_var = tk.IntVar()

    tk.Label(content_frame, text="File CSV di input:").grid(row=0, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=file_input_var, width=50).grid(row=0, column=1, pady=5)
    tk.Button(content_frame, text="Seleziona file CSV", command=lambda: choose_input_file(file_input_var, column_frame, column_vars, column_dest_vars)).grid(row=0, column=2, pady=5)

    tk.Label(content_frame, text="File Excel di output:").grid(row=1, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=file_output_var, width=50).grid(row=1, column=1, pady=5)
    tk.Button(content_frame, text="Seleziona file Excel", command=lambda: choose_output_file(file_output_var)).grid(row=1, column=2, pady=5)
    
    tk.Label(content_frame, text="Ultima Aggiunta?:").grid(row=2, column=0, pady=5, sticky='e')
    tk.Checkbutton(content_frame, variable=checkbox_var, command=lambda: on_checkbox_toggle(checkbox_var)).grid(row=2, column=1, pady=5)
    
    tk.Label(content_frame, text="File Excel di finale:").grid(row=3, column=0, pady=5, sticky='e')
    tk.Entry(content_frame, textvariable=file_final_var, width=50).grid(row=3, column=1, pady=5)
    tk.Button(content_frame, text="Seleziona file Excel", command=lambda: choose_finale_file(file_final_var)).grid(row=3, column=2, pady=5)

    tk.Label(content_frame, text="Seleziona le colonne da copiare:").grid(row=4, column=0, pady=5, sticky='e')

    column_frame = tk.Frame(content_frame)
    column_frame.grid(row=5, column=0, columnspan=3, pady=5, sticky='nsew')

    tk.Button(content_frame, text="Avvia procedura", command=lambda: start_procedure(file_input_var, file_output_var, file_final_var, column_vars, column_dest_vars, start_row_var, sheet_name_var, selected_option_var, dropdown_cell_var, checkbox_var.get())).grid(row=4, column=0, columnspan=3, pady=20)

    tk.Label(content_frame, text="Seleziona un'opzione:").grid(row=6, column=0, pady=5, sticky='e')
    dropdown_options = dropdown_options_var.get().split(", ")
    dropdown = tk.OptionMenu(content_frame, selected_option_var, *dropdown_options)
    dropdown.grid(row=6, column=1, pady=5)

    content_frame.grid_columnconfigure(0, weight=1)
    content_frame.grid_columnconfigure(1, weight=1)
    content_frame.grid_columnconfigure(2, weight=1)
    content_frame.grid_columnconfigure(3, weight=1)
    content_frame.grid_columnconfigure(4, weight=1)
    content_frame.grid_rowconfigure(5, weight=1)

    update_dropdown()