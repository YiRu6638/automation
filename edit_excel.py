import openpyxl
from openpyxl import Workbook
import os

def add_row_to_excel(data):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_file = os.path.join(script_dir, "output", "output-record.xlsx")
    
    try:
        workbook = openpyxl.load_workbook(output_file)
        print(f"Successfully opened the Excel file: {output_file}")
    except Exception as e:
        print(f"Error occured. {e}")
        exit(0)
        
    sheet = workbook.active
    
    sheet.append(data)
    
    workbook.save(output_file)
    print(f"File saved successfully at: {output_file}")
    
if __name__ == "__main__":
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    output_file = os.path.join(script_dir, "output", "output-record.xlsx")
    
    add_row_to_excel([0, 1, 2])