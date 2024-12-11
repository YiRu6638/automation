import openpyxl
import os

script_dir = os.path.dirname(os.path.abspath(__file__))
data_path = os.path.join(script_dir, "output", "output-record.xlsx")
template_path = os.path.join(script_dir, "template", "EmployeeName_PersonalIncomeTaxComputation_YA2024.xlsx")
template2_path = os.path.join(script_dir, "template", "EmployeeName_PersonalIncomeTaxComputation_YA2024_Non-resident.xlsx")
data = openpyxl.load_workbook(data_path)
data_sheet = data.active



def generate_tax_comp():
    for row in range(3, data_sheet.max_row + 1):
        row_data = list(data_sheet.iter_rows(min_row=row, max_row=row, min_col=1, max_col=26, values_only=True))[0]

        if row_data[21] == 'R':
            template_workbook = openpyxl.load_workbook(template_path)
            template_sheet = template_workbook.active
            
            template_sheet['C4'] = row_data[1]
            template_sheet['C5'] = row_data[0]
            template_sheet['E14'] = row_data[3]
            template_sheet['E15'] = row_data[4]
            template_sheet['E16'] = row_data[5]
            template_sheet['E17'] = row_data[6]
            template_sheet['E18'] = row_data[7]
            template_sheet['E19'] = row_data[8]
            template_sheet['E20'] = row_data[9]
            template_sheet['E21'] = row_data[10]
            template_sheet['E22'] = row_data[11]
            template_sheet['E23'] = row_data[12]
            template_sheet['E24'] = row_data[13]
            template_sheet['E27'] = row_data[25]
            template_sheet['E39'] = row_data[18]
            template_sheet['E42'] = row_data[17] * -1
            
            new_filepath = os.path.join(script_dir, "tax_comps", row_data[1] + "_PersonalIncomeTaxComputation_YA2024.xlsx")
            template_workbook.save(new_filepath)
        else:
            template_workbook = openpyxl.load_workbook(template2_path)
            template_sheet = template_workbook.active
            
            template_sheet['C4'] = row_data[1]
            template_sheet['C5'] = row_data[0]
            template_sheet['E14'] = row_data[3]
            template_sheet['E15'] = row_data[4]
            template_sheet['E16'] = row_data[5]
            template_sheet['E17'] = row_data[6]
            template_sheet['E18'] = row_data[7]
            template_sheet['E19'] = row_data[8]
            template_sheet['E20'] = row_data[9]
            template_sheet['E21'] = row_data[10]
            template_sheet['E22'] = row_data[11]
            template_sheet['E23'] = row_data[12]
            template_sheet['E24'] = row_data[13]
            template_sheet['E32'] = row_data[18]
            template_sheet['E35'] = row_data[17] * -1

            new_filepath = os.path.join(script_dir, "tax_comps", row_data[1] + "_PersonalIncomeTaxComputation_YA2024_Non-resident.xlsx")
            template_workbook.save(new_filepath)

if __name__ == "__main__":
    generate_tax_comp()
    pass