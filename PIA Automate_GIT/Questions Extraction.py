
import os
import re
import sys
from typing import List, Dict, Optional, Tuple
import pandas as pd
from PyPDF2 import PdfReader

# --------- CONFIGURE THESE PATHS ----------
SOURCE_DIR = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidatedpdfs"
DEST_EXCEL = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Desktop\PIA Rough work\Questions.xlsx"
SHEET_NAME = "List of Questions"
# ------------------------------------------


SECTION_LINE_RE = re.compile(
    r"""^
        (?:Section|Sec\.?)?\s*       # optional 'Section' or 'Sec.'
        (?P<maj>\d+)\.(?P<min>\d+)   # X.Y section number, e.g., 1.2 or 6.13
        \b                           # word boundary
        (?:\s*[\)\.\:\-–—])?         # optional punctuation ')' '.' ':' '-' en/em dash
        \s*                          # spaces
        (?P<rest>.*)                 # remainder of the line (question text if present)
        $""",
    re.IGNORECASE | re.VERBOSE
)


def clean_text_for_lines(text: str) -> List[str]:
    """
    Normalize PDF extracted text into a list of non-empty lines.
    Also attempts to fix hyphenation across line breaks.
    """
    if not text:
        return []

    # Normalize newlines
    text = text.replace("\r", "\n")

    # De-hyphenate words split across line breaks: "exam-\nple" -> "example"
    text = re.sub(r"(\w)-\n(\w)", r"\1\2", text)

    # Collapse multiple spaces
    text = re.sub(r"[ \t]+", " ", text)

    # Split to lines and strip
    lines = [ln.strip() for ln in text.split("\n")]
    return [ln for ln in lines if ln.strip()]


def clean_question_text(text: str) -> str:
    """
    Clean common noise from the question text portion after the section number.
    """
    # Remove leading labels or bullets
    text = re.sub(r"^\s*(Q(?:uestion)?[:\.\)]\s*)", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^\s*([•\-\u2022]\s*)", "", text)
    text = re.sub(r"^\s*(\d+\)|\(\d+\)|\d+\.\d+\s+)", "", text)
    text = re.sub(r"^\s*[:\-\–—\.]\s*", "", text)  # strip lingering separators
    text = re.sub(r"\s+", " ", text).strip()
    return text


def extract_questions_from_pdf(pdf_path: str) -> List[Dict[str, str]]:
    """
    Extract records of (question, section_number, pdf_name) strictly where a line starts
    with a section number 'X.Y'. If the line has only the section number, the next non-empty
    line (even across page boundaries) is treated as the question.
    """
    results: List[Dict[str, str]] = []
    seen: set[Tuple[str, str, str]] = set()  # (question_lower, section, pdf_name)
    pdf_name = os.path.basename(pdf_path)

    try:
        reader = PdfReader(pdf_path)
    except Exception as e:
        print(f"[WARN] Failed to read PDF: {pdf_path} | {e}", file=sys.stderr)
        return results

    pending_section: Optional[str] = None  # carry over when section line has no text

    # Iterate all pages
    for page_idx, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception as e:
            print(f"[WARN] Failed to extract text from {pdf_name} page {page_idx+1}: {e}", file=sys.stderr)
            continue

        lines = clean_text_for_lines(text)

        i = 0
        while i < len(lines):
            line = lines[i]

            # If we are waiting to capture the question for a pending section
            if pending_section:
                if SECTION_LINE_RE.match(line):
                    # A new section starts before we found the question for the pending one; drop pending
                    pending_section = None
                    # Do not advance i; let normal processing handle this same line
                else:
                    q_text = clean_question_text(line)
                    if q_text:
                        key = (q_text.lower(), pending_section, pdf_name)
                        if key not in seen:
                            seen.add(key)
                            results.append({
                                "List of Questions": q_text,
                                "Section #": pending_section,
                                "PIA name": pdf_name
                            })
                        pending_section = None
                    i += 1
                    continue  # move to next line

            # Normal detection: line must begin with X.Y section number
            m = SECTION_LINE_RE.match(line)
            if m:
                section = f"{m.group('maj')}.{m.group('min')}"
                rest = m.group("rest").strip()

                if rest:
                    q_text = clean_question_text(rest)
                    if q_text:
                        key = (q_text.lower(), section, pdf_name)
                        if key not in seen:
                            seen.add(key)
                            results.append({
                                "List of Questions": q_text,
                                "Section #": section,
                                "PIA name": pdf_name
                            })
                else:
                    # No text on same line: capture from the next non-empty line
                    pending_section = section

                i += 1
                continue

            # No match and no pending: skip
            i += 1

    # If pending_section remains at end of file without a question line following, ignore gracefully.
    return results


def main():
    if not os.path.isdir(SOURCE_DIR):
        print(f"[ERROR] Source directory does not exist:\n{SOURCE_DIR}")
        return

    all_records: List[Dict[str, str]] = []

    pdf_files = [f for f in os.listdir(SOURCE_DIR) if f.lower().endswith(".pdf")]
    if not pdf_files:
        print(f"[WARN] No PDF files found in:\n{SOURCE_DIR}")

    for fname in sorted(pdf_files):
        pdf_path = os.path.join(SOURCE_DIR, fname)
        print(f"[INFO] Processing: {pdf_path}")
        recs = extract_questions_from_pdf(pdf_path)
        all_records.extend(recs)

    # Create DataFrame with required headers and write to Excel
    df = pd.DataFrame(all_records, columns=["List of Questions", "Section #", "PIA name"])

    # Ensure destination folder exists
    dest_dir = os.path.dirname(DEST_EXCEL)
    if dest_dir and not os.path.isdir(dest_dir):
        try:
            os.makedirs(dest_dir, exist_ok=True)
        except Exception as e:
            print(f"[ERROR] Could not create destination directory '{dest_dir}': {e}", file=sys.stderr)

    try:
        with pd.ExcelWriter(DEST_EXCEL, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name=SHEET_NAME, index=False)
        print(f"[DONE] Wrote {len(df)} rows to:\n{DEST_EXCEL} (sheet: {SHEET_NAME})")
    except Exception as e:
        print(f"[ERROR] Failed to write Excel:\n{DEST_EXCEL}\nReason: {e}")


if __name__ == "__main__":
    main()
