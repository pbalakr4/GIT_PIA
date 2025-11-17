
import os
import pandas as pd

def sync_and_update_master_detailed(source_folder, consolidated_master_path):
    # Validate paths
    if not os.path.exists(source_folder):
        print(f"Source folder '{source_folder}' does not exist.")
        return
    if not os.path.exists(consolidated_master_path):
        print(f"Consolidated master file '{consolidated_master_path}' does not exist.")
        return

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

    # Load master file or initialize empty DataFrame
    try:
        master_df = pd.read_excel(consolidated_master_path, engine='openpyxl')
    except Exception:
        master_df = pd.DataFrame()

    # If master file is empty, copy source data and exit
    if master_df.empty:
        source_df.to_excel(consolidated_master_path, index=False, engine='openpyxl')
        print(f"Master was empty. Added all {len(source_df)} rows from source.")
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

    # Save updated master file
    try:
        updated_master_df.to_excel(consolidated_master_path, index=False, engine='openpyxl')
        print("\nProcess completed successfully.")

        # Display updates in tabular format
        if updated_changes:
            print("\nRows Updated:")
            print(pd.DataFrame(updated_changes))
        else:
            print("\nNo rows were updated.")

        if added_changes:
            print("\nRows Added:")
            print(pd.DataFrame(added_changes))
        else:
            print("\nNo rows were added.")

        print(f"\nTotal IDs updated: {len(set([c['ID'] for c in updated_changes]))}")
        print(f"Total IDs added: {len(added_changes)}")

    except Exception as e:
        print(f"Error saving updated master file: {e}")

if __name__ == "__main__":
    # ✅ Hardcoded paths
    source_folder = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Nov 2025"
    consolidated_master_path = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Consolidated_master.xlsx"

    sync_and_update_master_detailed(source_folder, consolidated_master_path)
