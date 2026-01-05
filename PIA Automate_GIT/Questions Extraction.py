import os
import re
import sys
from typing import Dict, List, Optional, Tuple
import pandas as pd
from PyPDF2 import PdfReader

# ========= USER CONFIG =========
EXTRACT_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Extract.xlsx"
PDF_FOLDER = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidatedpdfs"
MASTER_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidated_Master.xlsx"
MASTER_SHEET = "All up"
EXTRACT_SHEET = "Vendor Extraction"

# >>> Configurable stop words for QUESTIONS (single word line; case-insensitive)
STOP_WORDS = ["Comments", "Response", "Risks"]  # add more: "Notes", "Findings", ...

# >>> Configurable start words for RESPONSE block (single word line; case-insensitive)
RESPONSE_START_WORDS = ["Response"]  # add variants: "Responses", "Vendor Response"


# ========= LOGGING =========
def log(msg: str) -> None:
    print(f"[INFO] {msg}")


def warn(msg: str) -> None:
    print(f"[WARN] {msg}", file=sys.stderr)


# ========= ID / FILENAME UTILITIES =========
def normalize_id(val) -> Optional[str]:
    """Normalize an ID cell value to the first continuous digit sequence as a string. Returns None if no digits found."""
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
    m = re.search(r'_(\d+)$', name)
    return m.group(1) if m else None


# ========= PDF PARSING =========
# SECTION NUMBER RULE:
# - Must be dotted: X.Y (e.g., 1.3, 3.41, 13.2)
# - X = 1–99 -> [1-9]\d?
# - Y = 0–99 -> \d{1,2}
# - Appears at the start of a line, optionally followed by ')', '.', '-', '–' and whitespace,
#   then any trailing text on the same line (captured as group 2).
SECTION_LINE_RE = re.compile(r'^\s*([1-9]\d?\.\d{1,2})\s*(?:[)\.\-–]\s*)?(.*)$')


def build_stop_regex(words: List[str]) -> re.Pattern:
    """
    Build a regex matching a line that is exactly one of `words` (case-insensitive),
    allowing leading/trailing whitespace.
    """
    escaped = [re.escape(w) for w in words if w.strip()]
    if not escaped:
        return re.compile(r'^\b\B$')  # never matches if list is empty
    pattern = r'^\s*(?:' + '|'.join(escaped) + r')\s*$'
    return re.compile(pattern, flags=re.IGNORECASE)


STOP_WORD_RE = build_stop_regex(STOP_WORDS)
RESPONSE_START_RE = build_stop_regex(RESPONSE_START_WORDS)

# Whole-word regexes for vendor names (case-insensitive)
BLIS_WORD_RE = re.compile(r'\bblis\b', flags=re.IGNORECASE)
VISTAR_WORD_RE = re.compile(r'\bvistar\b', flags=re.IGNORECASE)


# ===== Helpers (normalize / filters) =====
def _normalize_line_for_questions(line: str) -> str:
    """
    Normalize a line for 'Questions' capture:
      - remove soft hyphens
      - collapse internal whitespace to single spaces
      - strip ends
    """
    if not line:
        return ""
    s = line.replace("\u00ad", "")
    s = " ".join(s.split())
    return s.strip()


def _normalize_line_for_response(line: str) -> str:
    """Keep formatting (for bullets), but remove soft hyphen and trim."""
    if not line:
        return ""
    s = line.replace("\u00ad", "")
    return s.strip()


def _norm_key(s: str) -> str:
    """Case-insensitive, whitespace-collapsed key for dedup."""
    return " ".join(s.split()).strip().lower()


def _is_page_number_line(line: str) -> bool:
    """
    Heuristics to skip standalone page markers / footers/headers:
      - 'Page 3', 'page 3', '3', '3 of 10', '- 3 -'
      - FRACTION formats like '1/13', '2/14', '13/13', also with spaces: '1 / 13'
      - with label: 'Page 1/13', 'page 1 / 13'
    """
    l = line.strip()
    if not l:
        return False

    # Simple number or "Page N" or "N of M" or dashed patterns
    if re.match(r'^(?:page\s*)?\d+\s*(?:of\s*\d+)?$', l, flags=re.IGNORECASE):
        return True
    if re.match(r'^[-–—]?\s*\d+\s*[-–—]?$', l):
        return True

    # Fraction style: "N/M" or "Page N/M" with optional spaces around '/'
    if re.match(r'^(?:page\s*)?\d+\s*/\s*\d+$', l, flags=re.IGNORECASE):
        return True

    return False


def _is_punctuation_only(line: str) -> bool:
    """
    True if the line contains only punctuation/graphics (e.g., '/', '—', '---', '***')
    Used to drop layout separators without affecting bullets like '- item'.
    """
    s = line.strip()
    if not s:
        return False
    return bool(re.fullmatch(r'[\s\-/_.|~*–—]+', s)) and len(s) <= 3


# ===== QUESTIONS CAPTURE =====
def extract_sections_questions(pdf_path: str) -> Dict[str, str]:
    """
    Build a mapping of section number -> captured question text (revised criteria):
      - Start capture immediately after the section number: include any text on the header line after the token,
        then continue capturing subsequent lines.
      - Stop when a line is exactly one of the STOP_WORDS (case-insensitive, only the word on the line).
      - Also stop when a new section header is encountered (to avoid mixing sections).
      - Do not add duplicate statements/lines.
      - Skip page-number lines.
    """
    section_to_question: Dict[str, str] = {}
    try:
        reader = PdfReader(pdf_path)
    except Exception as e:
        warn(f"Failed to open PDF '{pdf_path}': {e}")
        return section_to_question

    current_section: Optional[str] = None
    lines_seen: set = set()
    captured_lines: List[str] = []
    capturing: bool = False

    def finalize_current():
        nonlocal current_section, captured_lines, lines_seen, capturing
        if current_section is not None and current_section not in section_to_question:
            text = " ".join(captured_lines).strip()
            section_to_question[current_section] = text
        current_section = None
        captured_lines = []
        lines_seen = set()
        capturing = False

    for page_idx, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception as e:
            warn(f"Failed to extract text from page {page_idx} in '{pdf_path}': {e}")
            text = ""

        for raw_line in text.splitlines():
            line = raw_line.strip()

            # New section header?
            m = SECTION_LINE_RE.match(line)
            if m:
                if capturing:
                    finalize_current()
                current_section = m.group(1)
                captured_lines = []
                lines_seen = set()
                capturing = True
                after = _normalize_line_for_questions((m.group(2) or "").strip())
                if after and not _is_page_number_line(after):
                    key = _norm_key(after)
                    if key not in lines_seen:
                        captured_lines.append(after)
                        lines_seen.add(key)
                continue

            if capturing and current_section is not None:
                if STOP_WORD_RE.match(line):
                    finalize_current()
                    continue

                m2 = SECTION_LINE_RE.match(line)
                if m2:
                    finalize_current()
                    # Begin new capture scope on the same line
                    current_section = m2.group(1)
                    captured_lines = []
                    lines_seen = set()
                    capturing = True
                    after2 = _normalize_line_for_questions((m2.group(2) or "").strip())
                    if after2 and not _is_page_number_line(after2):
                        key2 = _norm_key(after2)
                        if key2 not in lines_seen:
                            captured_lines.append(after2)
                            lines_seen.add(key2)
                    continue

                norm = _normalize_line_for_questions(line)
                if norm and not _is_page_number_line(norm):
                    keyn = _norm_key(norm)
                    if keyn not in lines_seen:
                        captured_lines.append(norm)
                        lines_seen.add(keyn)

    if capturing and current_section is not None:
        finalize_current()

    return section_to_question


# ===== RESPONSE PARAGRAPHS (vendor keyword found) =====
def extract_response_vendor_paragraphs(pdf_path: str) -> Dict[str, Dict[str, List[str]]]:
    """
    Build a mapping:
        section_number -> {"Blis": [para1, para2...], "Vistar": [para1, ...]}
    Rules:
      - Start response capture after a line that is exactly one of RESPONSE_START_WORDS.
      - Stop response capture at the next section header.
      - Within the response block, split into paragraphs by blank lines (and single-word headings).
      - If a paragraph contains 'Blis' or 'Vistar' (whole-word), capture the entire paragraph.
      - Skip obvious page-number lines (including N/M formats).
      - Remove duplicate lines within a paragraph (case-insensitive, whitespace-normalized), preserving order.
      - Dedup paragraphs per vendor (case-insensitive, whitespace-normalized).
      - Preserve bullets/line breaks in the saved paragraph text.
    """
    result: Dict[str, Dict[str, List[str]]] = {}
    try:
        reader = PdfReader(pdf_path)
    except Exception as e:
        warn(f"Failed to open PDF '{pdf_path}': {e}")
        return result

    current_section: Optional[str] = None
    in_response: bool = False
    paragraph_lines: List[str] = []
    seen_line_keys_in_para: set = set()
    last_line_key: Optional[str] = None

    def ensure_section_key(sec: str) -> None:
        if sec not in result:
            result[sec] = {"Blis": [], "Vistar": []}

    def paragraph_text_dedup() -> str:
        """Return paragraph text with internal duplicate lines removed, keeping original order."""
        out_lines: List[str] = []
        seen_keys: set = set()
        for ln in paragraph_lines:
            k = _norm_key(ln)
            if k and k not in seen_keys:
                out_lines.append(ln)
                seen_keys.add(k)
        return "\n".join(out_lines).strip()

    def normalize_para_for_dedup(t: str) -> str:
        return " ".join(t.split()).strip().lower()

    # Track dedup sets per section/vendor
    dedup_sets: Dict[Tuple[str, str], set] = {}

    def add_para_if_contains_vendor(sec: str, text: str) -> None:
        if not text:
            return
        has_blis = bool(BLIS_WORD_RE.search(text))
        has_vistar = bool(VISTAR_WORD_RE.search(text))
        if not (has_blis or has_vistar):
            return
        ensure_section_key(sec)
        norm_key = normalize_para_for_dedup(text)
        if has_blis:
            key = (sec, "Blis")
            dedup_sets.setdefault(key, set())
            if norm_key not in dedup_sets[key]:
                result[sec]["Blis"].append(text)
                dedup_sets[key].add(norm_key)
        if has_vistar:
            key = (sec, "Vistar")
            dedup_sets.setdefault(key, set())
            if norm_key not in dedup_sets[key]:
                result[sec]["Vistar"].append(text)
                dedup_sets[key].add(norm_key)

    def reset_paragraph_state() -> None:
        nonlocal paragraph_lines, seen_line_keys_in_para, last_line_key
        paragraph_lines = []
        seen_line_keys_in_para = set()
        last_line_key = None

    def finalize_paragraph() -> None:
        if paragraph_lines and current_section and in_response:
            text = paragraph_text_dedup()
            add_para_if_contains_vendor(current_section, text)
        reset_paragraph_state()

    for page_idx, page in enumerate(reader.pages):
        try:
            text = page.extract_text() or ""
        except Exception as e:
            warn(f"Failed to extract text from page {page_idx} in '{pdf_path}': {e}")
            text = ""

        for raw_line in text.splitlines():
            line = raw_line.rstrip()

            # Section header?
            m = SECTION_LINE_RE.match(line.strip())
            if m:
                # stop any ongoing response block at the boundary
                if in_response:
                    finalize_paragraph()
                    in_response = False
                current_section = m.group(1).strip()
                continue

            if current_section is None:
                # ignore preface/cover for response capture
                continue

            # Response start line?
            if RESPONSE_START_RE.match(line.strip()):
                # starting a response block
                finalize_paragraph()
                in_response = True
                continue

            if not in_response:
                # outside response block -> ignore
                continue

            # inside response block
            # boundaries & filters
            if _is_page_number_line(line):
                # skip page markers
                continue

            # paragraph separators: blank lines or single-word headings
            if not line.strip():
                finalize_paragraph()
                continue
            if re.match(r'^[A-Za-z][A-Za-z ]*$', line.strip()) and len(line.strip().split()) == 1:
                # a single word like "Comments", "Notes" acts as a separator between paragraphs
                finalize_paragraph()
                continue
            if _is_punctuation_only(line):
                finalize_paragraph()
                continue

            # accumulate line (keep bullets and formatting), with duplicate-line suppression
            norm_line = _normalize_line_for_response(line)
            key = _norm_key(norm_line)
            if key and key != last_line_key and key not in seen_line_keys_in_para:
                paragraph_lines.append(norm_line)
                seen_line_keys_in_para.add(key)
                last_line_key = key

    # finalize at EOF
    if in_response:
        finalize_paragraph()

    return result


# ===== OCCURRENCE SCAN (vendor mentions + questions mapping) =====
def parse_pdf_occurrences(pdf_path: str, section_to_question: Dict[str, str]) -> List[Tuple[str, str, str]]:
    """
    Second pass: scan PDF and return occurrences:
        [(vendor, found_in_display, question_text), ...]
    Rules:
      - Mentions before the first section -> Found in = "Cover Page", Questions = "Cover Page".
      - Mentions within a section -> Found in = section number, Questions = section_to_question[section].
      - Whole-word matches only for 'Blis' or 'Vistar'.
    """
    occurrences: List[Tuple[str, str, str]] = []
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

        for raw_line in text.splitlines():
            line = raw_line.strip()

            # Section change?
            m = SECTION_LINE_RE.match(line)
            if m:
                current_section_num = m.group(1).strip()
                first_section_seen = True
                continue

            # Whole-word checks
            has_blis = bool(BLIS_WORD_RE.search(line))
            has_vistar = bool(VISTAR_WORD_RE.search(line))

            if not first_section_seen:
                # Cover Page mentions
                if has_blis:
                    occurrences.append(("Blis", "Cover Page", "Cover Page"))
                if has_vistar:
                    occurrences.append(("Vistar", "Cover Page", "Cover Page"))
            else:
                if current_section_num:
                    qtext = section_to_question.get(current_section_num, "")
                    if has_blis:
                        occurrences.append(("Blis", current_section_num, qtext))
                    if has_vistar:
                        occurrences.append(("Vistar", current_section_num, qtext))

    return occurrences


# ===== FILENAME VENDOR DETECTION =====
def detect_vendors_in_filename(filename: str) -> List[str]:
    """
    Detect vendor names present in the filename (case-insensitive).
    Returns any of ['Blis', 'Vistar'].
    Only matches whole words, so 'published.pdf' will NOT match 'Blis'.
    """
    lower = filename.lower()
    vendors: List[str] = []
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
    if not os.path.exists(EXTRACT_PATH):
        log(f"Creating new extract workbook at {EXTRACT_PATH}")
        df_new = pd.DataFrame(columns=master_columns)
        with pd.ExcelWriter(EXTRACT_PATH, engine="openpyxl", mode="w") as writer:
            df_new.to_excel(writer, sheet_name=EXTRACT_SHEET, index=False)

    df = pd.read_excel(EXTRACT_PATH, sheet_name=EXTRACT_SHEET, engine="openpyxl")

    for col in master_columns:
        if col not in df.columns:
            df[col] = pd.NA

    df = df[[*master_columns, *[c for c in df.columns if c not in master_columns]]]
    return df


def enforce_vendor_foundin_questions_response_source_at_PQRST(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure Column P (16th) = 'Vendor', Column Q (17th) = 'Found in',
           Column R (18th) = 'Questions', Column S (19th) = 'SourceFileName',
           Column T (20th) = 'Response Keyword Found'.
    Preserves existing columns/order; pads with blank columns if fewer than 20 columns.
    """
    for col in ["Vendor", "Found in", "Questions", "SourceFileName", "Response Keyword Found"]:
        if col not in df.columns:
            df[col] = pd.NA

    cols = list(df.columns)
    base_cols = [c for c in cols if
                 c not in ("Vendor", "Found in", "Questions", "SourceFileName", "Response Keyword Found")]

    while len(base_cols) < 15:
        pad_name = f"_Pad_{len(base_cols) + 1}"
        if pad_name not in df.columns:
            df[pad_name] = pd.NA
        base_cols.append(pad_name)
    base_cols.insert(15, "Vendor")

    while len(base_cols) < 16:
        pad_name = f"_Pad_{len(base_cols) + 1}"
        if pad_name not in df.columns:
            df[pad_name] = pd.NA
        base_cols.append(pad_name)
    base_cols.insert(16, "Found in")

    while len(base_cols) < 17:
        pad_name = f"_Pad_{len(base_cols) + 1}"
        if pad_name not in df.columns:
            df[pad_name] = pd.NA
        base_cols.append(pad_name)
    base_cols.insert(17, "Questions")

    while len(base_cols) < 18:
        pad_name = f"_Pad_{len(base_cols) + 1}"
        if pad_name not in df.columns:
            df[pad_name] = pd.NA
        base_cols.append(pad_name)
    base_cols.insert(18, "SourceFileName")

    while len(base_cols) < 19:
        pad_name = f"_Pad_{len(base_cols) + 1}"
        if pad_name not in df.columns:
            df[pad_name] = pd.NA
        base_cols.append(pad_name)
    base_cols.insert(19, "Response Keyword Found")

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

    # Prepare extract DF and enforce columns
    extract_df = ensure_extract_headers(master_columns)
    extract_df = enforce_vendor_foundin_questions_response_source_at_PQRST(extract_df)

    # Collect PDFs by ID
    id_to_pdfs = collect_pdf_matches(PDF_FOLDER)
    rows_appended = 0

    # Process each ID
    for nid, master_row in lookup.items():
        pdfs = id_to_pdfs.get(nid, [])
        if not pdfs:
            continue

        occurrences: List[
            Tuple[str, str, str, str, str]] = []  # (Vendor, Found in, SourceFileName, Questions, ResponseParagraphs)

        for pdf_path in pdfs:
            base = os.path.basename(pdf_path)

            # Filename occurrences
            vendors_in_name = detect_vendors_in_filename(base)
            for v in vendors_in_name:
                occurrences.append((v, "Filename", base, "Filename", pd.NA))

            # Content occurrences
            section_q = extract_sections_questions(pdf_path)
            response_map = extract_response_vendor_paragraphs(pdf_path)  # section -> {vendor: [paras...]}

            raw_occ = parse_pdf_occurrences(pdf_path, section_q)  # List[(vendor, found_in, question)]
            seen_pairs = set()
            for v, fin, q in raw_occ:
                key = (v, fin)
                if key not in seen_pairs:
                    seen_pairs.add(key)
                    # Prepare response paragraphs if found_in is a section
                    resp_text = pd.NA
                    if fin not in ("Cover Page", "Filename"):
                        paras = response_map.get(fin, {}).get(v, [])
                        if paras:
                            # Join paragraphs with blank line to preserve separation
                            resp_text = "\n\n".join(paras)
                    occurrences.append((v, fin, base, q, resp_text))

        if not occurrences:
            new_row = {c: pd.NA for c in extract_df.columns}
            for col in master_columns:
                new_row[col] = master_row.get(col, pd.NA)
            new_row["Vendor"] = "Other"
            new_row["Found in"] = "Not Applicable"
            new_row["Questions"] = "Not Applicable"
            new_row["SourceFileName"] = os.path.basename(pdfs[0]) if pdfs else pd.NA
            new_row["Response Keyword Found"] = pd.NA
            extract_df = pd.concat([extract_df, pd.DataFrame([new_row])], ignore_index=True)
            rows_appended += 1
        else:
            # Add each occurrence as its own row
            for vendor, found_in_display, source_file, question_text, response_text in occurrences:
                new_row = {c: pd.NA for c in extract_df.columns}
                for col in master_columns:
                    new_row[col] = master_row.get(col, pd.NA)
                new_row["Vendor"] = vendor
                new_row["Found in"] = found_in_display
                if found_in_display in ("Cover Page", "Filename"):
                    new_row["Questions"] = found_in_display
                    new_row["Response Keyword Found"] = pd.NA
                else:
                    new_row["Questions"] = (question_text or "").strip() or pd.NA
                    # Preserve bullets/line breaks; Excel will display '\n' as new lines in the cell
                    new_row["Response Keyword Found"] = response_text
                new_row["SourceFileName"] = source_file

                extract_df = pd.concat([extract_df, pd.DataFrame([new_row])], ignore_index=True)
                rows_appended += 1

    # Save
    extract_df = enforce_vendor_foundin_questions_response_source_at_PQRST(extract_df)
    save_extract_df(extract_df)
    log(f"Appended {rows_appended} row(s) into '{EXTRACT_SHEET}'.")
    log("Completed.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e: warn(f"Script failed: {e}")
