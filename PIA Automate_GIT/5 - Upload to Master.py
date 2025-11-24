
import win32com.client as win32

# Paths
EXTRACT_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Extract.xlsx"
MASTER_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\T-Ads - Privacy & CyberSecurity\Privacy\PIA Files\T-Ads PIAs Automation\Master.xlsx"

# Sheet names
SRC_SHEET_1 = "Raw Extract"
DST_SHEET_1 = "Master"
SRC_SHEET_2 = "ID to PD Mapping"
DST_SHEET_2 = "keyword to ID mapped"

def main():
    # Start Excel
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False  # Set to True if you want to see Excel during execution

    try:
        # Open workbooks
        src_wb = excel.Workbooks.Open(EXTRACT_PATH)
        dst_wb = excel.Workbooks.Open(MASTER_PATH)

        src_ws1 = src_wb.Sheets(SRC_SHEET_1)
        src_ws2 = src_wb.Sheets(SRC_SHEET_2)
        dst_ws1 = dst_wb.Sheets(DST_SHEET_1)
        dst_ws2 = dst_wb.Sheets(DST_SHEET_2)

        # Find last row in source and destination
        last_src1 = src_ws1.Cells(src_ws1.Rows.Count, 1).End(-4162).Row  # xlUp = -4162
        last_dst1 = dst_ws1.Cells(dst_ws1.Rows.Count, 1).End(-4162).Row

        last_src2 = src_ws2.Cells(src_ws2.Rows.Count, 1).End(-4162).Row
        last_dst2 = dst_ws2.Cells(dst_ws2.Rows.Count, 1).End(-4162).Row

        # Copy data excluding header
        if last_src1 > 1:
            src_ws1.Range(f"A2:A{last_src1}").EntireRow.Copy()
            dst_ws1.Range(f"A{last_dst1 + 1}").PasteSpecial(Paste=-4163)  # xlPasteValues = -4163

        if last_src2 > 1:
            src_ws2.Range(f"A2:A{last_src2}").EntireRow.Copy()
            dst_ws2.Range(f"A{last_dst2 + 1}").PasteSpecial(Paste=-4163)

        # Save destination workbook
        dst_wb.Save()

        # Print summary
        print(f"Rows copied to '{DST_SHEET_1}': {last_src1 - 1}")
        print(f"Rows copied to '{DST_SHEET_2}': {last_src2 - 1}")

    finally:
        # Close workbooks and quit Excel
        src_wb.Close(False)
        dst_wb.Close(True)
        excel.Quit()

if __name__ == "__main__":
    main()
