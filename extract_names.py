import pandas as pd

print("--- Multi-Sheet Names Extractor ---")

input_file = input("Enter the input Excel file name (e.g., file.xlsx): ").strip()
output_file = 'names_output.xlsx'

try:
    print(f"\nReading all sheets from '{input_file}'...")
    # sheet_name=None reads ALL sheets into a dictionary: {sheet_name: DataFrame}
    sheets_dict = pd.read_excel(input_file, sheet_name=None)
    
    combined_data = []

    for sheet_name, df in sheets_dict.items():
        # Check if the sheet has a 'Name' column
        if 'Name' in df.columns:
            # Create a clean subset for this sheet
            temp_df = pd.DataFrame()
            temp_df['Name'] = df['Name']
            
            # Automatically find the phone column (checks common variations)
            phone_col_found = None
            for col in ['Phone', 'Phone Number', 'Mobile', 'Phone_Number', 'Telephone']:
                if col in df.columns:
                    phone_col_found = col
                    break
            
            # Assign phone data if found, otherwise leave blank
            if phone_col_found:
                temp_df['Phone'] = df[phone_col_found]
            else:
                temp_df['Phone'] = "N/A"  # Fallback if no phone column matches
                
            # Add the sheet name as the 'Station'
            temp_df['Station'] = sheet_name
            
            combined_data.append(temp_df)
            print(f" -> Processed sheet: {sheet_name}")

    if not combined_data:
        print("\nError: None of the sheets contained a 'Name' column.")
    else:
        # Combine all sheets into one master dataframe
        master_df = pd.concat(combined_data, ignore_index=True)
        
        # Remove duplicates based on the 'Name' column (keeps the first occurrence)
        unique_df = master_df.drop_duplicates(subset=['Name'], keep='first')
        
        # Save to the new output file
        unique_df.to_excel(output_file, index=False)
        
        print(f"\nSuccess! Filtered and deduplicated data saved to '{output_file}'.")
        print(f"Total unique records saved: {len(unique_df)}")

except FileNotFoundError:
    print(f"\nError: Could not find '{input_file}'. Make sure it is in your folder.")
except Exception as e:
    print(f"\nAn error occurred: {e}")