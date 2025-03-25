"""
GUI implementation for the CSV to Excel converter application.
"""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from csv_processor import process_csv_to_excel, get_csv_columns
import os

class CsvToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to Excel Converter Pro")
        self.root.minsize(500, 600)
        
        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TButton', padding=5)
        self.style.configure('TLabel', padding=5)
        self.style.configure('TFrame', background='#f0f0f0')
        
        # Variables
        self.file_input_var = tk.StringVar()
        self.file_output_var = tk.StringVar()
        self.delimiter_var = tk.StringVar(value=",")
        self.encoding_var = tk.StringVar(value="utf-8")
        self.column_vars = {}
        
        self._create_menu_bar()
        self._create_widgets()
    
    def _create_menu_bar(self):
        """Creates the application menu bar."""
        menubar = tk.Menu(self.root)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Open CSV", command=self._choose_input_file)
        file_menu.add_command(label="Set Output Excel", command=self._choose_output_file)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self._show_about)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def _create_widgets(self):
        """Creates all GUI elements."""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Input file section
        input_frame = ttk.LabelFrame(main_frame, text="Input CSV File", padding="10")
        input_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(input_frame, text="CSV File Path:").grid(row=0, column=0, sticky="w")
        ttk.Entry(input_frame, textvariable=self.file_input_var, width=50).grid(row=1, column=0, sticky="ew")
        ttk.Button(input_frame, text="Browse...", command=self._choose_input_file).grid(row=1, column=1, padx=5)
        
        # CSV options
        options_frame = ttk.Frame(input_frame)
        options_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=5)
        
        ttk.Label(options_frame, text="Delimiter:").pack(side=tk.LEFT)
        delimiter_combo = ttk.Combobox(options_frame, textvariable=self.delimiter_var, 
                                     values=[",", ";", "\t", "|", ":"], width=3)
        delimiter_combo.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(options_frame, text="Encoding:").pack(side=tk.LEFT, padx=(10,0))
        encoding_combo = ttk.Combobox(options_frame, textvariable=self.encoding_var, 
                                    values=["utf-8", "latin1", "iso-8859-1", "cp1252", "ascii"], width=10)
        encoding_combo.pack(side=tk.LEFT)
        
        # Output file section
        output_frame = ttk.LabelFrame(main_frame, text="Output Excel File", padding="10")
        output_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(output_frame, text="Excel File Path:").grid(row=0, column=0, sticky="w")
        ttk.Entry(output_frame, textvariable=self.file_output_var, width=50).grid(row=1, column=0, sticky="ew")
        ttk.Button(output_frame, text="Browse...", command=self._choose_output_file).grid(row=1, column=1, padx=5)
        
        # Column selection section
        column_frame = ttk.LabelFrame(main_frame, text="Select Columns to Include", padding="10")
        column_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Create a canvas with scrollbar for many columns
        self.canvas = tk.Canvas(column_frame, highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(column_frame, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = ttk.Frame(self.canvas)
        
        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
        
        # Add select all/none buttons
        buttons_frame = ttk.Frame(self.scrollable_frame)
        buttons_frame.pack(fill=tk.X, pady=5)
        
        ttk.Button(buttons_frame, text="Select All", 
                  command=lambda: self._toggle_all_columns(True)).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="Deselect All", 
                  command=lambda: self._toggle_all_columns(False)).pack(side=tk.LEFT)
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(main_frame, textvariable=self.status_var, 
                                   relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X, pady=(5,0))
        
        # Process button
        ttk.Button(main_frame, text="Convert to Excel", command=self._start_procedure, 
                  style='Accent.TButton').pack(pady=10)
        
        # Configure style for accent button
        self.style.configure('Accent.TButton', foreground='white', background='#0078d7')
        self.style.map('Accent.TButton', 
                      background=[('active', '#005499'), ('pressed', '#004080')])
    
    def _show_about(self):
        """Shows the about dialog."""
        about_text = "CSV to Excel Converter Pro\n\nVersion 2.0\n\nA simple tool to convert CSV files to Excel format."
        messagebox.showinfo("About", about_text)
    
    def _update_status(self, message):
        """Updates the status bar."""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def _choose_input_file(self):
        """Opens a dialog to select the input CSV file and updates column checkboxes."""
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")],
            title="Select Input CSV File"
        )
        if file_path:
            self.file_input_var.set(file_path)
            self._update_status(f"Loading: {os.path.basename(file_path)}")
            self._update_column_checkboxes(file_path)
            self._update_status("Ready")
    
    def _choose_output_file(self):
        """Opens a dialog to select the output Excel file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            title="Select Output Excel File"
        )
        if file_path:
            self.file_output_var.set(file_path)
            self._update_status(f"Output set to: {os.path.basename(file_path)}")
    
    def _update_column_checkboxes(self, file_path):
        """Updates the column checkboxes based on the selected CSV file."""
        # Clear existing checkboxes (except the buttons frame)
        for widget in self.scrollable_frame.winfo_children():
            if not isinstance(widget, ttk.Frame):
                widget.destroy()
        
        try:
            delimiter = self.delimiter_var.get()
            encoding = self.encoding_var.get()
            columns = get_csv_columns(file_path, delimiter=delimiter, encoding=encoding)
            
            # Add column checkboxes
            self.column_vars = {}
            for column in columns:
                var = tk.BooleanVar(value=True)  # Default to selected
                self.column_vars[column] = var
                cb = ttk.Checkbutton(self.scrollable_frame, text=column, variable=var)
                cb.pack(anchor='w', pady=2)
                
            self._update_status(f"Loaded {len(columns)} columns from CSV")
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read CSV file:\n{str(e)}")
            self._update_status("Error loading CSV file")
    
    def _toggle_all_columns(self, state):
        """Select or deselect all columns."""
        for var in self.column_vars.values():
            var.set(state)
    
    def _start_procedure(self):
        """Process the CSV to Excel conversion."""
        csv_file = self.file_input_var.get()
        excel_file = self.file_output_var.get()
        
        if not csv_file:
            messagebox.showwarning("Input Error", "Please select an input CSV file.")
            return
            
        if not excel_file:
            messagebox.showwarning("Output Error", "Please specify an output Excel file.")
            return
        
        selected_columns = [col for col, var in self.column_vars.items() if var.get()]
        if not selected_columns:
            messagebox.showwarning("Selection Error", "Please select at least one column to include.")
            return
            
        try:
            self._update_status("Converting... Please wait")
            self.root.update()  # Force UI update
            
            process_csv_to_excel(
                csv_file, 
                excel_file, 
                selected_columns,
                delimiter=self.delimiter_var.get(),
                encoding=self.encoding_var.get()
            )
            
            messagebox.showinfo("Success", f"File successfully converted to:\n{excel_file}")
            self._update_status("Conversion completed successfully")
        except Exception as e:
            messagebox.showerror("Conversion Error", f"Failed to convert file:\n{str(e)}")
            self._update_status("Conversion failed")


def create_interface():
    """Creates and runs the main application interface."""
    root = tk.Tk()
    
    # Set window icon if available
    try:
        root.iconbitmap('icon.ico')  # Placeholder for icon file
    except:
        pass
    
    app = CsvToExcelApp(root)
    root.mainloop()
