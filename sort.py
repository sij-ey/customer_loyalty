import os
import pandas as pd

print("--- Name Sorter & Sheet Generator ---")

# Ask the user for inputs dynamically
input_file = input("Enter the input Excel file name (e.g., file.xlsx): ").strip()
sheet_name = input("Enter the name you want for the new sheet: ").strip()

output_file = 'transformed_output.xlsx'
sort_column = 'Name' # The column used for sorting

try:
    print(f"\nReading '{input_file}'...")
    df = pd.read_excel(input_file)

    # 1. Sort the data
    df_sorted = df.sort_values(by=sort_column, ascending=True)

    # Check if the output file already exists to decide writer mode ('a' to append, 'w' to create fresh)
    writer_mode = 'a' if os.path.exists(output_file) else 'w'

    print(f"Saving sorted data to '{output_file}' (Sheet: {sheet_name})...")
    
    # 2. Save into the Excel file
    with pd.ExcelWriter(output_file, engine='openpyxl', mode=writer_mode, if_sheet_exists='replace') as writer:
        
        # Form 1: The full dataset, cleanly sorted
        df_sorted.to_excel(writer, sheet_name=sheet_name, index=False)

        # Form 2: Conditional split (Optional - change 'CategoryColumn' if needed)
        category_col = 'CategoryColumn'  

        if category_col in df.columns:
            for category, group in df.groupby(category_col):
                cat_sheet_name = str(category)[:30].replace('/', '-')
                group.to_excel(writer, sheet_name=cat_sheet_name, index=False)
                print(f" -> Added category sheet: {cat_sheet_name}")

    print(f"\nSuccess! Sheet '{sheet_name}' was successfully added to '{output_file}'.")

except FileNotFoundError:
    print(f"\nError: Could not find '{input_file}'. Make sure it is in your folder and spelled correctly.")
except KeyError as e:
    print(f"\nColumn Error: {e}. Check if your column name matches your Excel file exactly.")
except Exception as e:
    print(f"\nAn error occurred: {e}")