import os
import re
import pandas as pd
from PyPDF2 import PdfReader
from openpyxl import load_workbook

# Hardcoded paths
MASTER_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidated_Master.xlsx"
EXTRACT_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Extract.xlsx"
PDF_FOLDER = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidatedpdfs"

# Columns to copy from master
COLUMNS_TO_COPY = ["ID", "Name", "Stage", "Date created", "Respondent", "Date submitted", "Date completed"]

# Multiple search phrases
SEARCH_PHRASES = [
    "Provide a detailed, non-technical description of the objectives and goals of the activity.",
    "Provide a brief, non-technical description of the project objectives for your system, product, or service.",
]

# Stop strings
STOP_STRINGS = ["Risks", "Comments", "Assessment questions"]

# Column name
COMBINED_COLUMN = "Description"

# ---------------- PART 1: Update Extract.xlsx ----------------
def update_extract():
    master_df = pd.read_excel(MASTER_PATH)
    extract_df = pd.read_excel(EXTRACT_PATH, sheet_name="Raw Extract")

    # Ensure required columns exist
    for col in COLUMNS_TO_COPY:
        if col not in extract_df.columns:
            extract_df[col] = ""

    # Ensure combined column exists
    if COMBINED_COLUMN not in extract_df.columns:
        extract_df[COMBINED_COLUMN] = pd.Series([""] * len(extract_df), dtype="object")
    else:
        extract_df[COMBINED_COLUMN] = extract_df[COMBINED_COLUMN].astype("object")

    # Sync rows from master to extract
    for _, row in master_df.iterrows():
        row_id = row["ID"]
        if row_id in extract_df["ID"].values:
            idx = extract_df[extract_df["ID"] == row_id].index[0]
            for col in COLUMNS_TO_COPY:
                if pd.notna(row[col]) and extract_df.at[idx, col] != row[col]:
                    extract_df.at[idx, col] = row[col]
        else:
            new_row = {col: row[col] if col in row else "" for col in COLUMNS_TO_COPY}
            new_row[COMBINED_COLUMN] = ""
            extract_df = pd.concat([extract_df, pd.DataFrame([new_row])], ignore_index=True)

    # Write back only to "Raw Extract" sheet without deleting others
    with pd.ExcelWriter(EXTRACT_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        extract_df.to_excel(writer, sheet_name="Raw Extract", index=False)

    print("✅ Extract.xlsx updated successfully (Raw Extract sheet only).")

# ---------------- PART 2: Clean and Extract Text ----------------
def clean_text(text):
    lines = text.splitlines()
    cleaned_lines = []
    prev_line = ""
    for line in lines:
        line = line.strip()
        if line and line != prev_line:
            cleaned_lines.append(line)
        prev_line = line

    text = " ".join(cleaned_lines)

    # Remove page numbers, timestamps, dates
    text = re.sub(r"\bPage\s*\d+\b", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", "", text)
    text = re.sub(r"\b\d{1,2}:\d{2}(?:\s?[APMapm]{2})?\b", "", text)
    text = re.sub(r"\b(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{1,2},\s+\d{4}\b", "", text)

    # Remove "Assessment questions"
    text = re.sub(r"Assessment questions", "", text, flags=re.IGNORECASE)

    # Remove section headers globally (but keep bullet numbers)
    text = re.sub(r"\bSection\s+\d+(\.\d+)*\b", "", text, flags=re.IGNORECASE)

    # Normalize spaces
    text = re.sub(r"\s+", " ", text).strip()

    return text

def extract_text_from_pdf(pdf_path, phrase, stop_strings):
    reader = PdfReader(pdf_path)
    raw_text = "\n".join(page.extract_text() for page in reader.pages if page.extract_text())
    raw_text = clean_text(raw_text)

    if phrase in raw_text:
        phrase_index = raw_text.find(phrase)
        response_index = raw_text.find("Response", phrase_index)
        if response_index != -1:
            after_response = raw_text[response_index + len("Response"):].lstrip()
            after_response = re.sub(r"^Response\s*", "", after_response, flags=re.IGNORECASE)

            # Stop pattern: section markers OR stop strings
            stop_pattern = r"\b\d+\.\d+\b|\bSection\s+\d+(\.\d+)*\b|\b(" + "|".join(map(re.escape, stop_strings)) + r")\b"
            stop_match = re.search(stop_pattern, after_response)
            if stop_match:
                return after_response[:stop_match.start()].strip()
            else:
                return after_response.strip()
    return None

# ---------------- PART 3: Process PDFs ----------------
def process_pdfs():
    extract_df = pd.read_excel(EXTRACT_PATH, sheet_name="Raw Extract")

        for idx, row in extract_df.iterrows():
        row_id = str(row["ID"])
        pdf_found = False
        for file in os.listdir(PDF_FOLDER):
            if file.endswith(".pdf"):
                file_id = file.split("_")[-1].split(".")[0]
                if file_id == row_id:
                    pdf_found = True
                    pdf_path = os.path.join(PDF_FOLDER, file)

                    responses = []
                    for phrase in SEARCH_PHRASES:
                        extracted_text = extract_text_from_pdf(pdf_path, phrase, STOP_STRINGS)
                        if extracted_text:
                            if extracted_text not in responses:
                                responses.append(extracted_text)
                            print(f"✅ Extracted for ID {row_id} | Phrase: {phrase} | Text: {extracted_text}")
                        else:
                            print(f"⚠️ No match for phrase '{phrase}' in PDF for ID {row_id}")

                    if responses:
                        combined_text = "; ".join(responses)
                        extract_df.at[idx, COMBINED_COLUMN] = combined_text
                    else:
                        extract_df.at[idx, COMBINED_COLUMN] = "Not found in PDF"
                    break
        if not pdf_found:
            extract_df.at[idx, COMBINED_COLUMN] = "Not found in PDF"
            print(f"❌ No PDF found for ID {row_id}")

    # Write back only to "Raw Extract" sheet without deleting others
    with pd.ExcelWriter(EXTRACT_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        extract_df.to_excel(writer, sheet_name="Raw Extract", index=False)

    print("✅ PDF processing completed and Raw Extract sheet updated.")

# ---------------- MAIN ----------------
if __name__ == "__main__":
    update_extract()
    process_pdfs()
