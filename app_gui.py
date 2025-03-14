"""
GUI implementation for the CSV to Excel converter application.
"""
import tkinter as tk
from tkinter import filedialog, messagebox
from csv_processor import process_csv_to_excel, get_csv_columns

class CsvToExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("CSV to Excel Application")
        self.file_input_var = tk.StringVar()
        self.file_output_var = tk.StringVar()
        self.column_vars = {}
        
        self._create_widgets()
    
    def _create_widgets(self):
        """Creates all GUI elements."""
        # Input file section
        tk.Label(self.root, text="Input CSV File:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.file_input_var, width=50).pack(pady=5)
        tk.Button(self.root, text="Select CSV File", command=self._choose_input_file).pack(pady=5)
        
        # Output file section
        tk.Label(self.root, text="Output Excel File:").pack(pady=5)
        tk.Entry(self.root, textvariable=self.file_output_var, width=50).pack(pady=5)
        tk.Button(self.root, text="Select Excel File", command=self._choose_output_file).pack(pady=5)
        
        # Column selection section
        tk.Label(self.root, text="Select Columns to Copy:").pack(pady=5)
        self.column_frame = tk.Frame(self.root)
        self.column_frame.pack(pady=5, fill=tk.BOTH, expand=True)
        
        # Process button
        tk.Button(self.root, text="Start Procedure", command=self._start_procedure).pack(pady=20)
    
    def _choose_input_file(self):
        """Opens a dialog to select the input CSV file and updates column checkboxes."""
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv")],
            title="Select Input CSV File"
        )
        if file_path:
            self.file_input_var.set(file_path)
            self._update_column_checkboxes(file_path)
    
    def _choose_output_file(self):
        """Opens a dialog to select the output Excel file."""
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="Select Output Excel File"
        )
        if file_path:
            self.file_output_var.set(file_path)
    
    def _update_column_checkboxes(self, file_path):
        """Updates the column checkboxes based on the selected CSV file."""
        # Clear existing checkboxes
        for widget in self.column_frame.winfo_children():
            widget.destroy()
        
        try:
            columns = get_csv_columns(file_path)
            
            # Create a canvas with scrollbar for many columns
            canvas = tk.Canvas(self.column_frame)
            scrollbar = tk.Scrollbar(self.column_frame, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas)
            
            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)
            
            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Add select all/none buttons
            buttons_frame = tk.Frame(scrollable_frame)
            buttons_frame.pack(fill="x", pady=5)
            
            tk.Button(buttons_frame, text="Select All", 
                      command=lambda: self._toggle_all_columns(True)).pack(side="left", padx=5)
            tk.Button(buttons_frame, text="Deselect All", 
                      command=lambda: self._toggle_all_columns(False)).pack(side="left")
            
            # Add column checkboxes
            self.column_vars = {}
            for column in columns:
                var = tk.BooleanVar(value=True)  # Default to selected
                self.column_vars[column] = var
                tk.Checkbutton(scrollable_frame, text=column, variable=var).pack(anchor='w')
                
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while reading the CSV file: {e}")
    
    def _toggle_all_columns(self, state):
        """Select or deselect all columns."""
        for var in self.column_vars.values():
            var.set(state)
    
    def _start_procedure(self):
        """Process the CSV to Excel conversion."""
        csv_file = self.file_input_var.get()
        excel_file = self.file_output_var.get()
        
        if not csv_file:
            messagebox.showwarning("No File Selected", "No CSV file selected. The operation has been canceled.")
            return
            
        if not excel_file:
            messagebox.showwarning("No File Selected", "No Excel file selected. The operation has been canceled.")
            return
        
        selected_columns = [col for col, var in self.column_vars.items() if var.get()]
        if not selected_columns:
            messagebox.showwarning("No Columns Selected", "No columns selected. The operation has been canceled.")
            return
            
        try:
            process_csv_to_excel(csv_file, excel_file, selected_columns)
            messagebox.showinfo("Success", f"Selected columns have been copied to {excel_file}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


def create_interface():
    """Creates and runs the main application interface."""
    root = tk.Tk()
    app = CsvToExcelApp(root)
    root.mainloop()