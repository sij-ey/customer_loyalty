import os
import pandas as pd

print("--- Name Sorter & Sheet Generator ---")

# Ask the user for inputs dynamically
input_file = input("Enter the input Excel file name (e.g., file.xlsx): ").strip()
sheet_name = input("Enter the name you want for the new sheet: ").strip()

output_file = "transformed_output.xlsx"
sort_column = "Name"

try:
    print(f"\nReading '{input_file}'...")
    df = pd.read_excel(input_file)

    # 1. Sort the data
    df_sorted = df.sort_values(by=sort_column, ascending=True)

    # 2. Determine whether to create a new workbook or append to an existing one
    file_exists = os.path.exists(output_file)

    print(
        f"Saving sorted data to '{output_file}' (Sheet: {sheet_name})...")

    if file_exists:
        # Existing workbook → add new sheets
        writer_mode = "a"

        with pd.ExcelWriter(output_file,engine="openpyxl",mode=writer_mode) as writer:

            # Main sorted sheet
            df_sorted.to_excel(writer,sheet_name=sheet_name,index=False)

            # Optional category sheets
            category_col = "CategoryColumn"

            if category_col in df.columns:
                for category, group in df.groupby(category_col):

                    cat_sheet_name = str(category)[:30].replace("/", "-")

                    # Skip if the category sheet already exists
                    if cat_sheet_name in writer.book.sheetnames:
                        print(f" -> Skipping existing sheet: {cat_sheet_name}")
                        continue

                    group.to_excel(writer,sheet_name=cat_sheet_name,index=False)

                    print(f" -> Added category sheet: {cat_sheet_name}")

    else:
        # File doesn't exist → create a new workbook
        writer_mode = "w"

        with pd.ExcelWriter(output_file,engine="openpyxl",mode=writer_mode) as writer:

            # Main sorted sheet
            df_sorted.to_excel(writer,sheet_name=sheet_name,index=False)

            # Optional category sheets
            category_col = "CategoryColumn"

            if category_col in df.columns:
                for category, group in df.groupby(category_col):

                    cat_sheet_name = str(category)[:30].replace("/", "-")

                    group.to_excel(writer,sheet_name=cat_sheet_name,index=False)

                    print(f" -> Added category sheet: {cat_sheet_name}")

    print(f"\nSuccess! Sheet '{sheet_name}' was successfully added to '{output_file}'.")

except FileNotFoundError:
    print(f"\nError: Could not find '{input_file}'. Make sure it is in your folder and spelled correctly.")

except KeyError as e:
    print(f"\nColumn Error: {e}. Check if your column name matches your Excel file exactly.")

except ValueError as e:
    print(f"\nExcel Error: {e}")

except Exception as e:
    print(f"\nAn error occurred: {e}")
