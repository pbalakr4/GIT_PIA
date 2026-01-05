
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Delete rows containing specific numeric IDs from multiple Excel files (all sheets),
then delete PDFs whose filenames end with _<ID>.pdf from provided folders.
Print detailed summaries for both operations.

Optimized for:
- IDs are numeric only.
- 'ID' is the header column in every sheet (usually first column); falls back to first column if missing.
"""

import os
import re
import sys
import argparse
from datetime import datetime
from typing import List, Dict, Set, Tuple, Optional

import pandas as pd


# -----------------------------
# Configuration (defaults)
# -----------------------------
DEFAULT_EXCEL_PATHS = [
    r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidated_master.xlsx",
    r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\T-Ads - Privacy & CyberSecurity\Privacy\PIA Files\T-Ads PIAs Automation\Master.xlsx",
    r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Extract.xlsx",
]

DEFAULT_PDF_FOLDERS = [
    r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\T-Ads - Privacy & CyberSecurity\Privacy\PIA Files\T-Ads PIAs Automation\PIAs All Up",
    r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidatedpdfs",
]

# Example IDs (comma-separated). You can override via CLI: --ids "12345,67892"
DEFAULT_IDS_CSV = "66038,43094,67118,66837,50388"

# Safety switches
DEFAULT_DRY_RUN = False       # If True, shows what would be deleted without changing files
DEFAULT_MAKE_BACKUP = True    # If True, makes a timestamped backup copy before overwriting Excel


# -----------------------------
# Helpers
# -----------------------------

def parse_ids_numeric(ids_csv: str) -> Set[int]:
    """
    Parse comma-separated IDs into a set of integers.
    Ignores blanks and non-numeric tokens.
    """
    ids: Set[int] = set()
    for token in ids_csv.split(","):
        token = token.strip()
        if not token:
            continue
        try:
            ids.add(int(float(token)))  # allows "12345.0" to be treated as 12345
        except ValueError:
            # Non-numeric token; skip
            pass
    return ids


def backup_excel(excel_path: str) -> str:
    """
    Create a timestamped backup copy of the Excel file in the same directory.
    """
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base = os.path.basename(excel_path)
    name, ext = os.path.splitext(base)
    backup_name = f"{name}_backup_{ts}{ext}"
    backup_path = os.path.join(os.path.dirname(excel_path), backup_name)
    import shutil
    shutil.copy2(excel_path, backup_path)
    return backup_path


def find_id_column(df: pd.DataFrame) -> Optional[str]:
    """
    Determine the 'ID' column name.
    - Prefer exact match 'ID' (case-insensitive).
    - If not found, prefer the first column if it looks numeric-heavy or if user says it's 99% first column.
    - Return None if dataframe has no columns.
    """
    if df is None or df.empty and len(df.columns) == 0:
        return None

    # Try case-insensitive match for 'ID'
    for col in df.columns:
        if str(col).strip().lower() == "id":
            return col

    # Fall back to the first column
    if len(df.columns) > 0:
        return df.columns[0]

    return None


def coerce_series_to_int(series: pd.Series) -> pd.Series:
    """
    Coerce a pandas Series to integers where possible.
    - Numeric strings or floats like '12345' or 12345.0 -> 12345
    - Non-numeric cells -> NaN (as pandas NA), so they won't match target IDs.
    """
    # Convert to numeric (errors='coerce' turns non-numeric into NaN), then drop decimals if integral
    s = pd.to_numeric(series, errors='coerce')
    # convert floats like 12345.0 to int; keep NaN
    # where not null and float is integer -> cast to int
    if s.dtype.kind in {'f', 'i'}:
        # Create integer series where possible
        return s.apply(lambda x: int(x) if pd.notna(x) and float(x).is_integer() else pd.NA)
    return s  # unexpected type, but still numeric-coerced series


def process_excel_file(
    excel_path: str,
    target_ids: Set[int],
    dry_run: bool = False,
    make_backup: bool = True,
) -> Tuple[List[str], Set[int]]:
    """
    Read an Excel file (all sheets), remove rows where ID column equals one of target_ids.
    Return:
      - summary_lines: list of strings about findings and deletions
      - removed_ids: set of IDs that were actually found (for later PDF deletion)
    """
    summary_lines: List[str] = []
    removed_ids: Set[int] = set()

    if not os.path.isfile(excel_path):
        summary_lines.append(f"[SKIP] Excel not found: {excel_path}")
        return summary_lines, removed_ids

    try:
        sheets_dict: Dict[str, pd.DataFrame] = pd.read_excel(
            excel_path, sheet_name=None, engine="openpyxl"
        )
    except Exception as e:
        summary_lines.append(f"[ERROR] Failed to read '{excel_path}': {e}")
        return summary_lines, removed_ids

    cleaned_sheets: Dict[str, pd.DataFrame] = {}
    excel_name = os.path.basename(excel_path)

    for sheet_name, df in sheets_dict.items():
        if df is None or (df.empty and len(df.columns) == 0):
            cleaned_sheets[sheet_name] = df
            continue

        id_col = find_id_column(df)
        if id_col is None:
            # No columns? keep as is
            cleaned_sheets[sheet_name] = df
            summary_lines.append(f"Excel {excel_name}, Sheet '{sheet_name}': [INFO] No columns found; skipped.")
            continue

        # Coerce ID column to integers (where possible)
        coerced_id_series = coerce_series_to_int(df[id_col])

        # Build mask for rows to delete: ID in target_ids
        row_has_id = coerced_id_series.apply(lambda x: pd.notna(x) and int(x) in target_ids)

        # Count per-ID occurrences in this sheet (only in ID column)
        id_counts: Dict[int, int] = {}
        if len(target_ids) > 0:
            # Efficient counting
            # Drop NA, cast to int, then count
            existing_ids = coerced_id_series.dropna().astype(int)
            for tid in target_ids:
                id_counts[tid] = int((existing_ids == tid).sum())

        # Log summary per ID
        for tid, count in id_counts.items():
            if count > 0:
                removed_ids.add(tid)
                summary_lines.append(
                    f"Excel {excel_name}, Sheet '{sheet_name}', ID {tid} found in {count} row{'s' if count != 1 else ''}, "
                    f"All {count} row{'s' if count != 1 else ''} {'would be deleted' if dry_run else 'deleted'}."
                )

        # Apply deletion
        if row_has_id.any():
            if dry_run:
                cleaned_sheets[sheet_name] = df.copy()
            else:
                cleaned_sheets[sheet_name] = df.loc[~row_has_id].copy()
        else:
            cleaned_sheets[sheet_name] = df

    # Determine if any sheet changed
    any_deletion = any(
        len(sheets_dict[name]) != len(cleaned_sheets[name]) for name in sheets_dict.keys()
    )

    if any_deletion and not dry_run:
        try:
            if make_backup:
                backup_path = backup_excel(excel_path)
                summary_lines.append(f"[BACKUP] Created backup: {backup_path}")
            with pd.ExcelWriter(excel_path, engine="openpyxl", mode="w") as writer:
                for sheet_name, cleaned_df in cleaned_sheets.items():
                    cleaned_df.to_excel(writer, sheet_name=sheet_name, index=False)
            summary_lines.append(f"[UPDATED] Saved cleaned workbook: {excel_path}")
        except Exception as e:
            summary_lines.append(f"[ERROR] Failed to save cleaned workbook '{excel_path}': {e}")
    elif any_deletion and dry_run:
        summary_lines.append(f"[DRY-RUN] No changes written for: {excel_path}")
    else:
        summary_lines.append(f"[NO-CHANGE] No matching IDs found in: {excel_name}")

    return summary_lines, removed_ids


def extract_id_from_pdf_filename(filename: str) -> Optional[int]:
    """
    Extract trailing numeric ID from filenames ending with _<ID>.pdf
    E.g., "MyDoc_12345.pdf" -> 12345 (int)
    Returns None if no match.
    """
    m = re.search(r"_(\d+)\.pdf$", filename, flags=re.IGNORECASE)
    if m:
        try:
            return int(m.group(1))
        except ValueError:
            return None
    return None


def process_pdf_folders(
    folders: List[str],
    removed_ids: Set[int],
    dry_run: bool = False,
) -> List[str]:
    """
    Delete PDFs whose filenames end with _<ID>.pdf in provided folders, for IDs in removed_ids.
    Return summary lines of deletions (or would-be deletions in dry-run).
    """
    summary_lines: List[str] = []
    canonical_ids: Set[int] = {int(x) for x in removed_ids if x is not None}

    for folder in folders:
        if not os.path.isdir(folder):
            summary_lines.append(f"[SKIP] Folder not found: {folder}")
            continue

        deleted_count = 0
        for entry in os.listdir(folder):
            full_path = os.path.join(folder, entry)
            if not os.path.isfile(full_path):
                continue
            if not entry.lower().endswith(".pdf"):
                continue

            id_in_name = extract_id_from_pdf_filename(entry)
            if id_in_name is not None and id_in_name in canonical_ids:
                if dry_run:
                    summary_lines.append(f"[DRY-RUN] Would delete: {full_path}")
                else:
                    try:
                        os.remove(full_path)
                        deleted_count += 1
                        summary_lines.append(f"[DELETED] {full_path}")
                    except Exception as e:
                        summary_lines.append(f"[ERROR] Could not delete '{full_path}': {e}")

        if deleted_count == 0 and not dry_run:
            summary_lines.append(f"[INFO] No matching PDFs deleted in: {folder}")
        elif dry_run:
            summary_lines.append(f"[DRY-RUN] Completed scan for: {folder}")

    return summary_lines


# -----------------------------
# Main / CLI
# -----------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Delete rows with numeric IDs from multiple Excel files and PDFs ending with _<ID>.pdf."
    )
    parser.add_argument(
        "--excel-paths",
        nargs="*",
        default=DEFAULT_EXCEL_PATHS,
        help="Paths to Excel files (space-separated)."
    )
    parser.add_argument(
        "--pdf-folders",
        nargs="*",
        default=DEFAULT_PDF_FOLDERS,
        help="Folders containing PDFs to clean (space-separated)."
    )
    parser.add_argument(
        "--ids",
        type=str,
        default=DEFAULT_IDS_CSV,
        help="Comma-separated list of numeric IDs, e.g., '12345,67892'."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        default=DEFAULT_DRY_RUN,
        help="Show what would be deleted without changing files."
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Do not create a backup before overwriting Excel."
    )

    args = parser.parse_args()

    target_ids = parse_ids_numeric(args.ids)
    make_backup = not args.no_backup

    print("\n=== START: Excel row deletions ===")
    all_summary_lines: List[str] = []
    all_removed_ids: Set[int] = set()

    for excel in args.excel_paths:
        lines, removed_ids = process_excel_file(
            excel_path=excel,
            target_ids=target_ids,
            dry_run=args.dry_run,
            make_backup=make_backup,
        )
        all_summary_lines.extend(lines)
        all_removed_ids.update(removed_ids)

    for line in all_summary_lines:
        print(line)

    print("\n=== START: PDF deletions ===")
    pdf_summary = process_pdf_folders(
        folders=args.pdf_folders,
        removed_ids=all_removed_ids,
        dry_run=args.dry_run,
    )
    for line in pdf_summary:
        print(line)

    print("\n=== COMPLETE ===")
    print(f"IDs processed: {sorted(list(target_ids))}")
    print(f"IDs removed from Excel (triggered PDF scan): {sorted(list(all_removed_ids))}")


if __name__ == "__main__":
    main()
