
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
    "Whose/What Personal Data is involved in this activity?",
    "Whose data is involved in this activity?"
]

# Stop strings
STOP_STRINGS = [
    "Risks", "Comments",
    "What is the estimated number of data subjects whose data will be processed?",
    "What operations will be performed on the personal data?",
    "What operations will be performed on the data?",
    "Assessment questions"
]

# Phrases to remove
REMOVE_PHRASES = [
    "Consumer","Classifications Protected by Law","Internet and Mobile Network Activity",
    "Inferred/Derived Information","Commercial Information","Non T-Mobile Customers or Prospects",
    "Select all that apply","Regular Identifiers","Customer Proprietary Network Information",
    "What is the estimated number of data subjects whose data will be processed?",
    "What operations will be performed on the personal data?",
    "What operations will be performed on the data?",
    "Assessment questions",
    "Select the groups of individuals you are processing data about",
    "If you did not select any data elements in the previous question"
]

# Column name
COMBINED_COLUMN = "What Personal Data is involved"

# Exception phrases to keep intact
EXCEPTION_PHRASES = [
    "Postal Code",
    "Interactions with third party mobile applications",
    "Wireless User Cellular Latitude and Longitude",
    "Individual's Language use or preference",
    "Interactions with advertisements",
    "Web Cookies or tracking tokens",
    "Interactions with third party internet websites"
]

# ---------------- Utility: Fix broken phrases ----------------
def fix_exceptions(text, exception_phrases):
    sep = r"(?:\s+|\s*\|\s*)"
    for phrase in exception_phrases:
        tokens = phrase.split()
        if len(tokens) < 2:
            continue
        pattern = r"\b" + sep.join(re.escape(tok) for tok in tokens) + r"\b"
        text = re.sub(pattern, " ".join(tokens), text, flags=re.IGNORECASE)
    return text

# ---------------- PART 1: Update Extract.xlsx ----------------
def update_extract():
    master_df = pd.read_excel(MASTER_PATH)
    extract_df = pd.read_excel(EXTRACT_PATH, sheet_name="Raw Extract")

    for col in COLUMNS_TO_COPY:
        if col not in extract_df.columns:
            extract_df[col] = ""

    if COMBINED_COLUMN not in extract_df.columns:
        extract_df[COMBINED_COLUMN] = pd.Series([""] * len(extract_df), dtype="object")
    else:
        extract_df[COMBINED_COLUMN] = extract_df[COMBINED_COLUMN].astype("object")

    for _, row in master_df.iterrows():
        row_id = row["ID"]
        if row_id in extract_df["ID"].values:
            idx = extract_df[extract_df["ID"] == row_id].index[0]
            for col in COLUMNS_TO_COPY:
                if pd.notna(row.get(col)) and extract_df.at[idx, col] != row[col]:
                    extract_df.at[idx, col] = row[col]
        else:
            new_row = {col: row[col] if col in row else "" for col in COLUMNS_TO_COPY}
            new_row[COMBINED_COLUMN] = ""
            extract_df = pd.concat([extract_df, pd.DataFrame([new_row])], ignore_index=True)

    with pd.ExcelWriter(EXTRACT_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        extract_df.to_excel(writer, sheet_name="Raw Extract", index=False)

    print("✅ Extract.xlsx updated successfully.")

# ---------------- PART 2: Extract text from PDFs ----------------
def clean_extracted_text(raw_text):
    # Split lines and remove empty
    lines = [line.strip() for line in raw_text.splitlines() if line.strip()]

    # Remove duplicates while preserving order
    seen = set()
    unique_lines = []
    for line in lines:
        if line not in seen:
            seen.add(line)
            unique_lines.append(line)

    # Join with separator
    text = " | ".join(unique_lines)

    # Global deduplication after joining
    parts = text.split(" | ")
    text = " | ".join(dict.fromkeys(parts))  # preserves order

    # Fix exceptions where ' | ' splits phrases
    text = fix_exceptions(text, EXCEPTION_PHRASES)

    # Normalize spacing
    text = re.sub(r"\s*\|\s*", " | ", text).strip(" |")
    return text

def remove_unwanted_phrases(text, phrases_to_remove):
    for phrase in phrases_to_remove:
        text = re.sub(r"\b" + re.escape(phrase) + r"\b", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\b\d+\s*/\s*\d+\b", "", text)
    text = re.sub(r"\d{4}\s+\d{1,2}:\d{2}\s*(AM|PM)?", "", text)
    text = re.sub(r"\d{1,2}:\d{2}\s*(AM|PM)?", "", text)
    text = re.sub(r"\d{2}/\d{2}/\d{4}", "", text)
    text = re.sub(r"\s*\|\s*", " | ", text).strip(" |")
    return text

def extract_text_from_pdf(pdf_path, phrase, stop_strings):
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() for page in reader.pages if page.extract_text())

    # Debug: Show first 500 chars of raw text
    print(f"\n[DEBUG] Extracted text from {os.path.basename(pdf_path)} (first 500 chars):")
    print(text[:500])

    if phrase in text:
        phrase_index = text.find(phrase)
        response_index = text.find("Response", phrase_index)
        if response_index != -1:
            after_response = text[response_index + len("Response"):].lstrip()
            after_response = re.sub(r"^Response\s*", "", after_response, flags=re.IGNORECASE)

            stop_pattern = r"(" + "|".join(map(re.escape, stop_strings)) + r")"
            section_pattern = r"\b\d+\.\d+\b"

            stop_match = re.search(stop_pattern, after_response)
            section_match = re.search(section_pattern, after_response)

            matches = [m.start() for m in [stop_match, section_match] if m]
            if matches:
                cut_pos = min(matches)
                raw_text = after_response[:cut_pos].strip()
            else:
                raw_text = after_response.strip()

            # Debug: Show raw extracted section before cleaning
            print(f"[DEBUG] Raw extracted section for phrase '{phrase}':")
            print(raw_text[:500])

            cleaned_text = clean_extracted_text(raw_text)
            cleaned_text = remove_unwanted_phrases(cleaned_text, REMOVE_PHRASES)

            # Debug: Show cleaned text
            print(f"[DEBUG] Cleaned text after processing:")
            print(cleaned_text)

            return cleaned_text
    return None

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
                    print(f"\n📄 Processing PDF for ID {row_id}: {pdf_path}")

                    responses = []
                    for phrase in SEARCH_PHRASES:
                        extracted_text = extract_text_from_pdf(pdf_path, phrase, STOP_STRINGS)
                        if extracted_text:
                            responses.append(extracted_text)
                            print(f"✅ Extracted for ID {row_id} | Phrase: {phrase}")
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

    with pd.ExcelWriter(EXTRACT_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        extract_df.to_excel(writer, sheet_name="Raw Extract", index=False)

    print("\n✅ PDF processing completed and Raw Extract sheet updated.")

# ---------------- MAIN ----------------
if __name__ == "__main__":
    update_extract()
    process_pdfs()
