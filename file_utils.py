from tkinter import filedialog

def select_file(dialog_type="open", file_type=None, title="Select File"):
    if dialog_type == "open":
        file_path = filedialog.askopenfilename(
            filetypes=[file_type] if file_type else [],
            title=title
        )
    else:
        file_path = filedialog.asksaveasfilename(
            defaultextension=file_type[1] if file_type else None,
            filetypes=[file_type] if file_type else [],
            title=title
        )
    return file_path

def select_input_file():
    return select_file(dialog_type="open", file_type=("CSV Files", "*.csv"), title="Select Input CSV File")

def select_output_file():
    return select_file(dialog_type="save", file_type=("Excel Files", "*.xlsx"), title="Select Output Excel File")

def select_final_file():
    return select_file(dialog_type="open", file_type=("Excel Files", "*.xlsx"), title="Select Final Excel File")
