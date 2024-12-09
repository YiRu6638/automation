import os
import re
import edit_excel

# Column, Description, Regex pattern
patterns = [
    ("A", "Tax Identification No. (TIN)", r"No\. Pengenalan Cukai \(TIN\) (.+)"),
    ##### SECTION A #####
    ("B", "Name", r"1\. Nama Penuh Pekerja/Pesara \(Encik/Cik/Puan\) (.+)"),
    ("C", "No. Kad Pengenalan / No. Passport", r"5\. No\. Pasport (.+)"),
    ##### SECTION B #####
    ("D", "Gross salary", r"1\. \(a\) Gaji kasar, upah atau gaji cuti \(termasuk gaji lebih masa\) (.+)"),
    ("E", "Fees", r"\(b\) Fi \(termasuk fi pengarah\), komisen atau bonus (.+)"),
    ("F", "Gross tips", r"\(c\) Tip kasar, perkuisit, penerimaan sagu hati atau elaun-elaun lain \(Perihal pembayaran\.: ............ a\) (.+)"),
    ("G", "Income tax paid", r"\(d\) Cukai Pendapatan yang dibayar oleh Majikan bagi pihak Pekerja (.+)"),
    ("H", "Manfaat Skim Opsyen Saham Pekerja (ESOS)", r"\(e\) Manfaat Skim Opsyen Saham Pekerja \(ESOS\) (.+)"),
    ("I", "Remuneration for the period", r"\(f\) Ganjaran bagi tempoh dari \.\.\.\. hingga \.\.\.\.\.\. (.+)"),
    ("J", "Details of arrears", r"2\. Butiran bayaran tunggakan dan lain-lain bagi tahun-tahun terdahulu dalam tahun semasa (.+)"),
    ("K", "Benefit in kinds", r"3\. Manfaat berupa barangan \(Nyatakan: .................. o\) (.+)"),
    ("L", "Value of living accommodation", r"4\. Nilai tempat kediaman \(Alamat: ........ (.+)"),
    ("M", "Refund from unapproved", r"5\. Bayaran balik daripada Kumpulan Wang Simpanan/Pencen yang tidak diluluskan (.+)"),
    ("N", "Compensation for loss", r"6\. Pampasan kerana kehilangan pekerjaan (.+)"),
    ##### FORMULA
    ("O", "Formula Total Gross Employment Income", r"nanikore"),
    ##### SECTION C #####
    ("P", "Pension", r"1. Pencen (.+)"),
    ("Q", "Annuities or other periodical payments", r"2\. Anuiti atau Bayaran Berkala yang lain (.+)"),
    ##### SECTION D #####
    ("R", "Monthly tax deductions (MTD) remitted to LHDNM", r"1\. Potongan Cukai Bulanan \(PCB\) yang dibayar kepada LHDNM (.+)"),
    ("S", "Zakat paid via salary deduction", r"3\. Zakat yang dibayar melalui potongan gaji (.+)"),
    ##### SECTION E #####
    ("T", "Relief - EPF (employee's contribution)", r"Amaun caruman yang wajib dibayar \(nyatakan bahagian pekerja sahaja\) RM (.+)"),
    ("U", "Relief - SOCSO / EIS (employee's contribution)", r"2\. PERKESO : Amaun caruman yang wajib dibayar \(nyatakan bahagian pekerja sahaja\) RM (.+)"),
    ##### SECTION F #####
    # r"JUMLAH ELAUN / PERKUISIT / PEMBERIAN / MANFAAT YANG DIKECUALIKAN CUKAI RM (.+)"
]

def find_in_text(large_text):
    results = {}
    
    for i, (label, description, pattern) in enumerate(patterns):
        if pattern:  # Only process if a regex pattern is provided
            match = re.search(pattern, large_text)
            if match:
                results[description] = match.group(1)
            else:
                results[description] = None
        else:
            results[description] = "No pattern provided"
            
    return results

def extract_text_to_excel(text_folder):
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Define the folder containing the files
    folder_path = os.path.join(script_dir, text_folder)
    
    os.makedirs(text_folder, exist_ok=True)

    # List all files in the folder
    if os.path.exists(folder_path):
        file_paths = sorted(os.listdir(folder_path))
        for file_name in file_paths:
            file_path = os.path.join(folder_path, file_name)
            
            # Check if the current item is a file (not a subdirectory)
            if os.path.isfile(file_path):
                print(f"Processing file: {file_path}")
                
                # Open and process the file
                with open(file_path, 'r') as file:
                    file_content = file.read()
                
                results = find_in_text(file_content)
                
                # Convert dict to list
                converted_data = list(results.values())
                edit_excel.add_row_to_excel(converted_data)
    else:
        print(f"Folder does not exist: {folder_path}")
    

if __name__ == "__main__":
    # script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # single_txt = os.path.join(script_dir, "dummy_text", "EDIT1-2023-Form-EA_Part2.txt")
    
    # with open(single_txt, 'r') as file:
    #     file_content = file.read()
    
    # results = find_in_text(file_content)
    # print(results)
    
    extract_text_to_excel("dummy_text")