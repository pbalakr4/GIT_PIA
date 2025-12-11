
import os
import shutil
import pandas as pd

def consolidate_pdfs(source_folder, consolidated_folder, newfiles_base_path):
    if not os.path.exists(source_folder):
        print(f"Source folder '{source_folder}' does not exist.")
        return
    if not os.path.exists(consolidated_folder):
        print(f"Consolidated folder '{consolidated_folder}' does not exist.")
        return

    excel_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.xlsx')]
    if not excel_files:
        print("No Excel file found in source folder.")
        return

    excel_path = os.path.join(source_folder, excel_files[0])
    print(f"Reading Excel file: {excel_path}")

    try:
        df = pd.read_excel(excel_path, usecols=[0])
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    excel_numbers = df.iloc[:, 0].dropna().astype(str).tolist()

    source_folder_name = os.path.basename(source_folder.rstrip("\\/"))
    new_folder_name = f"newfiles_{source_folder_name}"
    new_folder_path = os.path.join(newfiles_base_path, new_folder_name)
    os.makedirs(new_folder_path, exist_ok=True)

    pdf_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.pdf')]
    new_files_count = 0

    for pdf in pdf_files:
        try:
            number_part = pdf.split('_')[-1].split('.')[0]
        except IndexError:
            continue

        if number_part in excel_numbers:
            source_pdf_path = os.path.join(source_folder, pdf)
            consolidated_pdf_path = os.path.join(consolidated_folder, pdf)

            if not os.path.exists(consolidated_pdf_path):
                shutil.copy2(source_pdf_path, consolidated_pdf_path)
                shutil.copy2(source_pdf_path, os.path.join(new_folder_path, pdf))
                new_files_count += 1

    print(f"\nProcess completed. Total new PDFs copied: {new_files_count}")
    print(f"New files folder: {new_folder_path}")


if __name__ == "__main__":
    # OPTION 1: Interactive input
    # source_folder = input("Enter source folder path: ").strip()
    # consolidated_folder = input("Enter consolidated folder path: ").strip()
    # newfiles_base_path = input("Enter base path for newfiles folder: ").strip()

    # OPTION 2: Hardcode paths safely using raw strings
    source_folder = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Dec 2025"
    consolidated_folder = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidatedpdfs"
    newfiles_base_path = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Monthlynewfiles"
    consolidate_pdfs(source_folder, consolidated_folder, newfiles_base_path)
