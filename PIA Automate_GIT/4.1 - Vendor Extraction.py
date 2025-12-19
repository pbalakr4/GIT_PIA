
import os
import re
import sys
from typing import Dict, List, Optional, Tuple
import pandas as pd
from PyPDF2 import PdfReader

# ========= USER CONFIG =========
EXTRACT_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Extract.xlsx"
PDF_FOLDER   = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidatedpdfs"
MASTER_PATH  = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidated_Master.xlsx"
MASTER_SHEET = "All up"
EXTRACT_SHEET = "Vendor Extraction"

# ========= LOGGING =========
def log(msg: str) -> None:
    print(f"[INFO] {msg}")

def warn(msg: str) -> None:
    print(f"[WARN] {msg}", file=sys.stderr)

# ========= ID / FILENAME UTILITIES =========
def normalize_id(val) -> Optional[str]:
    """
    Normalize an ID cell value to the first continuous digit sequence as a string.
    Returns None if no digits found.
    """
    if pd.isna(val):
        return None
    s = str(val)
    m = re.search(r'(\d+)', s)
    return m.group(1) if m else None

def extract_id_from_filename(filename: str) -> Optional[str]:
    """
    Extract trailing numeric ID that appears after the LAST '_' in the filename (before extension).
    Example: 'SomeFile_12345.pdf' -> '12345'
    """
    base = os.path.basename(filename)
    name, _ = os.path.splitext(base)
    # Look for final underscore + digits at the end of the name
    m = re.search(r'_(\d+)$', name)
    return m.group(1) if m else None

# ========= PDF PARSING =========
# SECTION NUMBER RULE (updated):
# - Must be dotted: X.Y (e.g., 1.3, 3.41, 13.2)
# - X = 1–99 -> [1-9]\d?
# - Y = 0–99 -> \d{1,2}
# - Appears at the start of a line, followed optionally by punctuation like ')', '.', '-', '–' and whitespace.
STRICT_SECTION_RE = re.compile(
    r'^\s*([1-9]\d?\.\d{1,2})\s*(?:[\)\.\-–]\s*)?.*$'
)

# Whole-word regexes for vendor names (case-insensitive)
BLIS_WORD_RE   = re.compile(r'\bblis\b', flags=re.IGNORECASE)
VISTAR_WORD_RE = re.compile(r'\bvistar\b', flags=re.IGNORECASE)

def parse_pdf_occurrences(pdf_path: str) -> List[Tuple[str, str]]:
    """
    Scan a PDF and return a list of occurrences: [("Blis", "2.1"), ("Vistar", "4.2"), ("Blis", "Cover Page"), ...]
    - Mentions *before* the first valid section are labeled "Cover Page".
    - Mentions *under* a section record the section number only (no question text).
    - Only whole-word matches for 'Blis' or 'Vistar' are counted.
    """
    occurrences: List[Tuple[str, str]] = []
    try:
        reader = PdfReader(pdf_path)
    except Exception as e:
        warn(f"Failed to open PDF '{pdf_path}': {e}")
        return occurrences

    current_section_num: Optional[str] = None
    first_section_seen: bool = False

    for page_idx, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception as e:
            warn(f"Failed to extract text from page {page_idx} in '{pdf_path}': {e}")
            continue

        # Optional normalization (enable if you see hyphenation artifacts)
        # text = text.replace('\u00ad', '')  # remove soft hyphens

        for line in text.splitlines():
            line_stripped = line.strip()

            # If line starts with a valid dotted section number, update current_section_num
            m = STRICT_SECTION_RE.match(line_stripped)
            if m:
                sec_num = m.group(1).strip()
                current_section_num = sec_num
                first_section_seen = True
                continue

            # Whole-word checks (case-insensitive)
            has_blis   = bool(BLIS_WORD_RE.search(line_stripped))
            has_vistar = bool(VISTAR_WORD_RE.search(line_stripped))

            # Before any section appears -> "Cover Page"
            if not first_section_seen:
                if has_blis:
                    occurrences.append(("Blis", "Cover Page"))
                if has_vistar:
                    occurrences.append(("Vistar", "Cover Page"))
            else:
                # Under a known section -> record the section number
                if current_section_num:
                    if has_blis:
                        occurrences.append(("Blis", current_section_num))
                    if has_vistar:
                        occurrences.append(("Vistar", current_section_num))

    return occurrences

def detect_vendors_in_filename(filename: str) -> List[str]:
    """
    Detect vendor names present in the filename (case-insensitive).
    Returns any of ['Blis', 'Vistar'].
    Only matches whole words, so 'published.pdf' will NOT match 'Blis'.
    """
    lower = filename.lower()
    vendors: List[str] = []

    # Whole-word checks in filenames
    if re.search(r'\bblis\b', lower, flags=re.IGNORECASE):
        vendors.append("Blis")
    if re.search(r'\bvistar\b', lower, flags=re.IGNORECASE):
        vendors.append("Vistar")

    return vendors

# ========= EXCEL IO =========
def read_master() -> pd.DataFrame:
    log(f"Loading master from {MASTER_PATH} (sheet '{MASTER_SHEET}')")
    df = pd.read_excel(MASTER_PATH, sheet_name=MASTER_SHEET, engine="openpyxl")
    if df.empty:
        raise ValueError("Master sheet is empty.")
    return df

def ensure_extract_headers(master_columns: List[str]) -> pd.DataFrame:
    """
    Ensure Extract.xlsx exists with sheet 'Vendor Extraction'.
    If sheet is missing or empty, initialize with master headers.
    Return current sheet as DataFrame (with at least master columns).
    """
    # Create file and sheet if needed
    if not os.path.exists(EXTRACT_PATH):
        log(f"Creating new extract workbook at {EXTRACT_PATH}")
        df_new = pd.DataFrame(columns=master_columns)
        with pd.ExcelWriter(EXTRACT_PATH, engine="openpyxl", mode="w") as writer:
            df_new.to_excel(writer, sheet_name=EXTRACT_SHEET, index=False)

    # Load sheet
    df = pd.read_excel(EXTRACT_PATH, sheet_name=EXTRACT_SHEET, engine="openpyxl")

    # Align to master columns (extras will be handled later)
    for col in master_columns:
        if col not in df.columns:
            df[col] = pd.NA

    # Reorder: master columns first, then any extras that were present
    df = df[[*master_columns, *[c for c in df.columns if c not in master_columns]]]
    return df

def enforce_vendor_foundin_source_at_PQR(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure Column P (16th, index 15) = 'Vendor',
           Column Q (17th, index 16) = 'Found in',
           Column R (18th, index 17) = 'SourceFileName'.
    Preserves existing columns/order; pads with blank columns if fewer than 18 columns.
    """
    # Add targets if missing
    for col in ["Vendor", "Found in", "SourceFileName"]:
        if col not in df.columns:
            df[col] = pd.NA

    cols = list(df.columns)
    base_cols = [c for c in cols if c not in ("Vendor", "Found in", "SourceFileName")]

    # Pad up to index 15 for Vendor
    while len(base_cols) < 15:
        pad_name = f"_Pad_{len(base_cols) + 1}"
        if pad_name not in df.columns:
            df[pad_name] = pd.NA
        base_cols.append(pad_name)

    # Insert Vendor at P (index 15)
    base_cols.insert(15, "Vendor")

    # Pad up to index 16 for Found in
    while len(base_cols) < 16:
        pad_name = f"_Pad_{len(base_cols) + 1}"
        if pad_name not in df.columns:
            df[pad_name] = pd.NA
        base_cols.append(pad_name)

    # Insert Found in at Q (index 16)
    base_cols.insert(16, "Found in")

    # Pad up to index 17 for SourceFileName
    while len(base_cols) < 17:
        pad_name = f"_Pad_{len(base_cols) + 1}"
        if pad_name not in df.columns:
            df[pad_name] = pd.NA
        base_cols.append(pad_name)

    # Insert SourceFileName at R (index 17)
    base_cols.insert(17, "SourceFileName")

    # Append any remaining columns (including prior extras)
    remaining = [c for c in df.columns if c not in base_cols]
    final_cols = base_cols + remaining
    df = df[final_cols]
    return df

def save_extract_df(df: pd.DataFrame) -> None:
    with pd.ExcelWriter(EXTRACT_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=EXTRACT_SHEET, index=False)
    log(f"Saved updates to {EXTRACT_PATH} (sheet '{EXTRACT_SHEET}')")

# ========= FILE MATCHING =========
def collect_pdf_matches(pdf_folder: str) -> Dict[str, List[str]]:
    """
    Walk the folder and return a mapping: normalized_id -> list of PDF paths.
    Only includes files whose name contains a trailing _<digits>.
    """
    mapping: Dict[str, List[str]] = {}
    for root, _, files in os.walk(pdf_folder):
        for f in files:
            if not f.lower().endswith(".pdf"):
                continue
            fullpath = os.path.join(root, f)
            fid = extract_id_from_filename(f)
            if fid:
                mapping.setdefault(fid, []).append(fullpath)

    log(f"Indexed {sum(len(v) for v in mapping.values())} PDF(s) across {len(mapping)} ID(s).")
    return mapping

# ========= MAIN PROCESS =========
def main() -> None:
    master_df = read_master()

    # Build normalized ID column (assume Column A is the first column in master)
    id_col_name = master_df.columns[0]
    master_df["_NormalizedID"] = master_df[id_col_name].apply(normalize_id)

    # Lookup dict: normalized_id -> full row dict (excluding helper column)
    master_columns = [c for c in master_df.columns if c != "_NormalizedID"]
    lookup: Dict[str, Dict[str, object]] = {}
    for _, row in master_df.iterrows():
        nid = row["_NormalizedID"]
        if pd.isna(nid):
            continue
        lookup[str(nid)] = {col: row[col] for col in master_columns}

    # Prepare extract DF with master headers, then enforce Vendor/Found in/SourceFileName at P/Q/R
    extract_df = ensure_extract_headers(master_columns)
    extract_df = enforce_vendor_foundin_source_at_PQR(extract_df)

    # Collect PDFs by ID (by trailing _<digits> in filename)
    id_to_pdfs = collect_pdf_matches(PDF_FOLDER)
    rows_appended = 0

    # Process each ID that has at least one matching PDF
    for nid, master_row in lookup.items():
        pdfs = id_to_pdfs.get(nid, [])
        if not pdfs:
            # No matched PDF for this ID -> skip (per your scope).
            continue

        # Gather occurrences for this ID (per PDF, with dedup)
        occurrences: List[Tuple[str, str, str]] = []  # (Vendor, Found in, SourceFileName)

        for pdf_path in pdfs:
            base = os.path.basename(pdf_path)

            # 1) Filename occurrences: add one row per vendor (per file)
            vendors_in_name = detect_vendors_in_filename(base)
            for v in vendors_in_name:
                occurrences.append((v, "Filename", base))

            # 2) Content occurrences: deduplicate per (vendor, section) within this PDF
            raw_occ = parse_pdf_occurrences(pdf_path)  # List[(vendor, section or 'Cover Page')]
            seen_pairs = set()
            for v, fin in raw_occ:
                key = (v, fin)
                if key not in seen_pairs:
                    seen_pairs.add(key)
                    occurrences.append((v, fin, base))

        if not occurrences:
            # No vendor found anywhere -> add one "Other" row with Found in = "Not Applicable"
            new_row = {c: pd.NA for c in extract_df.columns}
            for col in master_columns:
                new_row[col] = master_row.get(col, pd.NA)
            new_row["Vendor"] = "Other"
            new_row["Found in"] = "Not Applicable"
            new_row["SourceFileName"] = os.path.basename(pdfs[0]) if pdfs else pd.NA

            extract_df = pd.concat([extract_df, pd.DataFrame([new_row])], ignore_index=True)
            rows_appended += 1
            continue

        # Add each occurrence as its own row
        for vendor, found_in_display, source_file in occurrences:
            new_row = {c: pd.NA for c in extract_df.columns}
            for col in master_columns:
                new_row[col] = master_row.get(col, pd.NA)
            new_row["Vendor"] = vendor
            new_row["Found in"] = found_in_display
            new_row["SourceFileName"] = source_file

            extract_df = pd.concat([extract_df, pd.DataFrame([new_row])], ignore_index=True)
            rows_appended += 1

    # Save
    extract_df = enforce_vendor_foundin_source_at_PQR(extract_df)  # Re-ensure final placement
    save_extract_df(extract_df)
    log(f"Appended {rows_appended} row(s) into '{EXTRACT_SHEET}'.")
    log("Completed.")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        warn(f"Script failed: {e}")
