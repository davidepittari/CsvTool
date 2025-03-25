"""
Unit tests for the CSV to Excel converter.
"""
import unittest
import pandas as pd
import os
import tempfile
from csv_processor import get_csv_columns, process_csv_to_excel

class TestCsvProcessor(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        """Create test files for all tests."""
        cls.temp_dir = tempfile.TemporaryDirectory()
        
        # Create test CSV files
        cls.csv_file = os.path.join(cls.temp_dir.name, "test.csv")
        cls.csv_file_semicolon = os.path.join(cls.temp_dir.name, "test_semicolon.csv")
        cls.csv_file_tab = os.path.join(cls.temp_dir.name, "test_tab.csv")
        cls.empty_csv_file = os.path.join(cls.temp_dir.name, "empty.csv")
        
        # Create normal CSV
        df = pd.DataFrame({
            'Name': ['Alice', 'Bob', 'Charlie'],
            'Age': [25, 30, 35],
            'City': ['New York', 'London', 'Paris']
        })
        df.to_csv(cls.csv_file, index=False)
        
        # Create semicolon-delimited CSV
        df.to_csv(cls.csv_file_semicolon, index=False, sep=';')
        
        # Create tab-delimited CSV
        df.to_csv(cls.csv_file_tab, index=False, sep='\t')
        
        # Create empty CSV
        pd.DataFrame().to_csv(cls.empty_csv_file, index=False)
    
    @classmethod
    def tearDownClass(cls):
        """Clean up test files."""
        cls.temp_dir.cleanup()
    
    def test_get_csv_columns(self):
        """Test getting columns from CSV file."""
        columns = get_csv_columns(self.csv_file)
        self.assertEqual(columns, ['Name', 'Age', 'City'])
    
    def test_get_csv_columns_semicolon(self):
        """Test getting columns from semicolon-delimited CSV."""
        columns = get_csv_columns(self.csv_file_semicolon, delimiter=';')
        self.assertEqual(columns, ['Name', 'Age', 'City'])
    
    def test_get_csv_columns_tab(self):
        """Test getting columns from tab-delimited CSV."""
        columns = get_csv_columns(self.csv_file_tab, delimiter='\t')
        self.assertEqual(columns, ['Name', 'Age', 'City'])
    
    def test_get_csv_columns_empty(self):
        """Test getting columns from empty CSV."""
        with self.assertRaises(Exception):
            get_csv_columns(self.empty_csv_file)
    
    def test_process_csv_to_excel(self):
        """Test CSV to Excel conversion."""
        excel_file = os.path.join(self.temp_dir.name, "test_output.xlsx")
        process_csv_to_excel(self.csv_file, excel_file, ['Name', 'City'])
        
        # Verify the Excel file was created
        self.assertTrue(os.path.exists(excel_file))
        
        # Verify the content
        df = pd.read_excel(excel_file)
        self.assertEqual(list(df.columns), ['Name', 'City'])
        self.assertEqual(len(df), 3)
    
    def test_process_csv_to_excel_missing_columns(self):
        """Test CSV to Excel with missing columns."""
        excel_file = os.path.join(self.temp_dir.name, "test_output.xlsx")
        with self.assertRaises(Exception):
            process_csv_to_excel(self.csv_file, excel_file, ['Name', 'Country'])
    
    def test_process_csv_to_excel_wrong_delimiter(self):
        """Test CSV to Excel with wrong delimiter."""
        excel_file = os.path.join(self.temp_dir.name, "test_output.xlsx")
        with self.assertRaises(Exception):
            process_csv_to_excel(self.csv_file_semicolon, excel_file, ['Name', 'City'], delimiter=',')

if __name__ == '__main__':
    unittest.main()
