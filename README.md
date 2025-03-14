# CSV to Excel Converter

## Overview
This application provides a user-friendly GUI for converting CSV files into Excel format. Users can select a CSV file, choose specific columns, and export the data to an Excel file.

## Features
- **Graphical User Interface (GUI):** Built with Tkinter for easy interaction.
- **Column Selection:** Allows users to choose which columns to include in the Excel file.
- **Error Handling:** Displays messages for missing files or invalid selections.
- **Excel Export:** Saves selected data as an `.xlsx` file using Pandas.

## Requirements
- Python 3.x
- Required Python packages:
  - `pandas`
  - `tkinter`
  - `openpyxl`

To install the required packages, run:
```bash
pip install pandas openpyxl
```

## Installation
1. Clone the repository:
```bash
git clone <repository_url>
cd <repository_folder>
```
2. Create a virtual environment:
```bash
python -m venv venv
```
3. Activate the virtual environment:
   - On Windows:
     ```bash
     venv\Scripts\activate
     ```
   - On macOS/Linux:
     ```bash
     source venv/bin/activate
     ```
4. Install dependencies (if not installed already):
```bash
pip install pandas openpyxl
```

## Usage
1. Run the application:
```bash
python main.py
```
2. Select an input CSV file.
3. Choose the columns you want to export.
4. Select an output Excel file.
5. Click **Start Procedure** to generate the Excel file.

## Project Structure
```
📂 project_folder
│── main.py             # Entry point of the application
│── app_gui.py          # GUI implementation
│── csv_processor.py    # CSV processing logic
```

## Error Handling
- If the selected CSV file is invalid, an error message will be displayed.
- If no columns are selected, the process will be canceled.
- If an issue occurs during file processing, an error dialog will notify the user.

## License
This project is licensed under the MIT License.

## Contributions
Contributions are welcome! Feel free to submit issues or pull requests.

