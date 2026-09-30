import pandas as pd

# Input and output files
input_file = input("Enter the input Excel file name (e.g., file.xlsx): ").strip()
output_file = input("Enter the output Excel file name (e.g., file.xlsx): ").strip()

# Read all sheets
sheets = pd.read_excel(input_file, sheet_name=None)

# Combine all sheets
dataframes = []

for sheet_name, df in sheets.items():
    df["Station"] = sheet_name
    dataframes.append(df)

combined_data = pd.concat(dataframes, ignore_index=True)

# Name of the date/timestamp column
date_column = "Date"

# Convert the column to datetime
combined_data[date_column] = pd.to_datetime(
    combined_data[date_column],
    errors="coerce"
)

# Sort newest/latest first
combined_data = combined_data.sort_values(
    by=date_column,
    ascending=False
)

# Save the final file
combined_data.to_excel(
    output_file,
    index=False,
    engine="openpyxl"
)

print(f"Done! Combined and sorted file saved as: {output_file}")
