import pandas as pd

# Define the file path
EXTRACT_PATH = r"C:\Users\PBalakr4\OneDrive - T-Mobile USA\Documents\PIA Automate\Extract.xlsx"

# Load sheets into DataFrames
data_identifiers_df = pd.read_excel(EXTRACT_PATH, sheet_name="Data Identifiers")
raw_extract_df = pd.read_excel(EXTRACT_PATH, sheet_name="Raw Extract")

# Prepare the output DataFrame for "ID to PD Mapping"
id_to_pd_mapping_cols = list(raw_extract_df.columns) + list(data_identifiers_df.columns[:3])
id_to_pd_mapping_df = pd.DataFrame(columns=id_to_pd_mapping_cols)

# Step 1: Match keywords and copy rows
for _, keyword_row in data_identifiers_df.iterrows():
    keyword = str(keyword_row["Keywords"]).strip()
    # Find rows in Raw Extract where "What Personal Data is involved" contains the keyword (literal match)
    matched_rows = raw_extract_df[
        raw_extract_df["What Personal Data is involved"].str.contains(keyword, case=False, na=False, regex=False)
    ]

    for _, matched_row in matched_rows.iterrows():
        # Extract all columns from Raw Extract
        raw_data = matched_row.tolist()
        # Extract columns 1-3 from Data Identifiers
        identifier_data = keyword_row.iloc[:3].tolist()
        # Combine and append
        id_to_pd_mapping_df.loc[len(id_to_pd_mapping_df)] = raw_data + identifier_data

# Step 2: Check IDs that were not matched
matched_ids = set(id_to_pd_mapping_df.iloc[:, 0])  # IDs already in mapping
for _, raw_row in raw_extract_df.iterrows():
    raw_id = raw_row.iloc[0]
    if raw_id not in matched_ids:
        # Add all columns from Raw Extract
        raw_data = raw_row.tolist()
        # Add "No Keyword found" in column after Raw Extract columns, and empty for next two
        id_to_pd_mapping_df.loc[len(id_to_pd_mapping_df)] = raw_data + ["No Keyword found", "", ""]

# Write back to Excel
with pd.ExcelWriter(EXTRACT_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
    id_to_pd_mapping_df.to_excel(writer, sheet_name="ID to PD Mapping", index=False)

print("Process completed successfully!")
