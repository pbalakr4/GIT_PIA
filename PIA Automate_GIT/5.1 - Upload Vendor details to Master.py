
#!/usr/bin/env python
from openpyxl import load_workbook

EXTRACT_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Extract.xlsx"
MASTER_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\T-Ads - Privacy & CyberSecurity\Privacy\PIA Files\T-Ads PIAs Automation\Master.xlsx"

VENDOR_SRC_SHEET = "Vendor Extraction"
VENDOR_DST_SHEET = "Vendor Details"


def copy_vendor_details():
    # Load workbooks
    src_wb = load_workbook(EXTRACT_PATH, data_only=True)
    dst_wb = load_workbook(MASTER_PATH)

    # Validate source sheet exists
    if VENDOR_SRC_SHEET not in src_wb.sheetnames:
        print(f"[ERROR] Source sheet '{VENDOR_SRC_SHEET}' not found in Extract file.")
        return

    src_ws = src_wb[VENDOR_SRC_SHEET]

    # Get or create destination sheet
    if VENDOR_DST_SHEET in dst_wb.sheetnames:
        dst_ws = dst_wb[VENDOR_DST_SHEET]
        # Clear all rows including headers
        if dst_ws.max_row > 0:
            dst_ws.delete_rows(1, dst_ws.max_row)
    else:
        dst_ws = dst_wb.create_sheet(title=VENDOR_DST_SHEET)

    # Copy all rows including headers
    rows_copied = 0
    for row in src_ws.iter_rows(values_only=True):
        dst_ws.append(row)
        rows_copied += 1

    # Save changes
    dst_wb.save(MASTER_PATH)
    print(f"[SUCCESS] Copied {rows_copied} rows from '{VENDOR_SRC_SHEET}' to '{VENDOR_DST_SHEET}' (including headers).")


def main():
    copy_vendor_details()


if __name__ == "__main__":
    main()
