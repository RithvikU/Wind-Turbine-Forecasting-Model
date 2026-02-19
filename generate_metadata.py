import pandas as pd

# === Load Excel file ===
file_path = r"C:\Users\rithv\OneDrive\MyData\College\Projects\Wind Turbine Research\Final Database1.xlsx"

# Read all sheets
xls = pd.read_excel(file_path, sheet_name=None)

# Collect metadata
metadata = []
for sheet_name, df in xls.items():
    print(f"Processing sheet: {sheet_name}")
    row_count = df.shape[0]
    col_count = df.shape[1]
    non_null_count = df.count().sum()

    metadata.append({
        "Sheet Name": sheet_name,
        "Row Count": row_count,
        "Column Count": col_count,
        "Non-Null Cells": non_null_count
    })

# Convert to DataFrame
meta_df = pd.DataFrame(metadata)

# Save metadata summary
output_file = "metadata_summary_all_sheets.xlsx"
meta_df.to_excel(output_file, index=False)

print(f"\nMetadata summary saved as: {output_file}")
print(meta_df)
