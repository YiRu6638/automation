from docx import Document
import os
import openpyxl

# Define the paths
script_dir = os.path.dirname(os.path.abspath(__file__))
data_path = os.path.join(script_dir, "output", "output-record.xlsx")
template_path = os.path.join(script_dir, "template", "EmployeeName_ConfirmationOfAgreementToTheTaxReturn.docx")

data = openpyxl.load_workbook(data_path)
data_sheet = data.active

# Define placeholder replacements
placeholders = (
    "{{name}}",
    "{{TIN number}}"
)

if __name__ == "__main__":
    # Replace placeholders in the document
    for row in range(3, data_sheet.max_row + 1):
        row_data = list(data_sheet.iter_rows(min_row=row, max_row=row, min_col=1, max_col=26, values_only=True))[0]
        
        doc = Document(template_path)
        
        for paragraph in doc.paragraphs:
            paragraph.text = paragraph.text.replace("{{name}}", row_data[1])
            paragraph.text = paragraph.text.replace("{{TIN number}}", row_data[0])
                    
        # Replace placeholders in tables (if needed)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    cell.text = cell.text.replace("{{name}}", row_data[1])
                    cell.text = cell.text.replace("{{TIN number}}", row_data[0])

        new_filepath = os.path.join(script_dir, "coa", row_data[1] + "_ConfirmationOfAgreementToTheTaxReturn.docx")
        # Save the modified document
        doc.save(new_filepath)
        print(f"Document saved to {new_filepath}")
