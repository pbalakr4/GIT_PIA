
import os
import pandas as pd
from datetime import datetime

def sync_and_update_master_detailed(source_folder, consolidated_master_path, sheet_name="All up", log_dir="C:/Users/PBalakr4/OneDrive - T-Mobile USA/Documents/PIA Automate/Logs"):
    # Validate paths
    if not os.path.exists(source_folder):
        print(f"Source folder '{source_folder}' does not exist.")
        return
    if not os.path.exists(consolidated_master_path):
        print(f"Consolidated master file '{consolidated_master_path}' does not exist.")
        return

    # Ensure log directory exists
    os.makedirs(log_dir, exist_ok=True)

    # Generate log file name based on script name and current timestamp
    script_name = os.path.splitext(os.path.basename(__file__))[0] if '__file__' in globals() else 'Consolidated_MasterExcel_Create'
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    log_file_name = f"{script_name}_{timestamp}.txt"
    log_file_path = os.path.join(log_dir, log_file_name)

    # Find the first Excel file in the source folder
    excel_files = [f for f in os.listdir(source_folder) if f.lower().endswith('.xlsx')]
    if not excel_files:
        print("No Excel file found in source folder.")
        return

    source_excel_path = os.path.join(source_folder, excel_files[0])
    print(f"Reading source Excel file: {source_excel_path}")

    try:
        source_df = pd.read_excel(source_excel_path, engine='openpyxl')
    except Exception as e:
        print(f"Error reading source Excel file: {e}")
        return

    # Load master sheet or initialize empty DataFrame
    try:
        master_df = pd.read_excel(consolidated_master_path, sheet_name=sheet_name, engine='openpyxl')
    except Exception:
        master_df = pd.DataFrame()

    # If master sheet is empty, copy source data and exit
    if master_df.empty:
        with pd.ExcelWriter(consolidated_master_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            source_df.to_excel(writer, sheet_name=sheet_name, index=False)
        log_text = f"[{datetime.now()}] Master sheet '{sheet_name}' was empty. Added all {len(source_df)} rows from source.\n"
        with open(log_file_path, "w", encoding="utf-8") as log_file:
            log_file.write(log_text)
        print(f"Log saved to: {log_file_path}\n")
        print(log_text)
        return

    # Assume first column is the unique identifier (ID)
    id_col = source_df.columns[0]

    updated_changes = []  # For logging updates
    added_changes = []    # For logging additions

    # Convert IDs to string for comparison
    master_df[id_col] = master_df[id_col].astype(str)
    source_df[id_col] = source_df[id_col].astype(str)

    # Create a copy of master for updates
    updated_master_df = master_df.copy()

    for _, src_row in source_df.iterrows():
        src_id = src_row[id_col]
        if src_id in updated_master_df[id_col].values:
            # Check for differences column by column
            idx = updated_master_df[updated_master_df[id_col] == src_id].index[0]
            for col in source_df.columns:
                old_val = updated_master_df.at[idx, col]
                new_val = src_row[col]
                if pd.notna(new_val) and old_val != new_val:
                    updated_master_df.at[idx, col] = new_val
                    updated_changes.append({
                        "ID": src_id,
                        "Column": col,
                        "Old Value": old_val,
                        "New Value": new_val
                    })
        else:
            # Add new row
            updated_master_df = pd.concat([updated_master_df, pd.DataFrame([src_row])], ignore_index=True)
            added_changes.append(src_row.to_dict())

    # Save updated master sheet
    try:
        with pd.ExcelWriter(consolidated_master_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            updated_master_df.to_excel(writer, sheet_name=sheet_name, index=False)

        # Prepare log content
        log_lines = [f"[{datetime.now()}] Process completed successfully."]
        if updated_changes:
            log_lines.append("Rows Updated:")
            log_lines.append(pd.DataFrame(updated_changes).to_string(index=False))
        else:
            log_lines.append("No rows were updated.")

        if added_changes:
            log_lines.append("Rows Added:")
            log_lines.append(pd.DataFrame(added_changes).to_string(index=False))
        else:
            log_lines.append("No rows were added.")

        log_lines.append(f"Total IDs updated: {len(set([c['ID'] for c in updated_changes]))}")
        log_lines.append(f"Total IDs added: {len(added_changes)}")
        log_text = "\n".join(log_lines) + "\n\n"

        # Write to new log file
        with open(log_file_path, "w", encoding="utf-8") as log_file:
            log_file.write(log_text)

        # Print to console
        print(f"Log saved to: {log_file_path}\n")
        print(log_text)

    except Exception as e:
        print(f"Error saving updated master file: {e}")

if __name__ == "__main__":
    # ✅ Hardcoded paths and sheet name
    source_folder = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Nov 2025"
    consolidated_master_path = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidated_master.xlsx"
    sheet_name = "All up"
    log_dir = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Logs"

    sync_and_update_master_detailed(source_folder, consolidated_master_path, sheet_name, log_dir)
