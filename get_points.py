import pandas as pd
from pathlib import Path

OUTPUT_FILE = "names_point.xlsx"

NAME_COLUMNS = [
    "name",
    "full name",
    "fullname",
    "customer name",
    "member name",
]

PHONE_COLUMNS = [
    "phone",
    "phone number",
    "mobile",
    "mobile number",
    "phone_number",
    "telephone",
    "tel",
]

POINTS_COLUMNS = [
    "points",
    "point",
    "points earned",
    "points_earned",
    "earned points",
    "score",
    "total points",
]

AMOUNT_COLUMNS = [
    "amount",
    "amount paid",
    "Amount (KES)",
    "payment amount",
    "transaction amount",
    "sale amount",
    "purchase amount",
    "value",
]

DATE_COLUMNS = [
    "date",
    "points date",
    "date earned",
    "earned date",
    "transaction date",
    "transaction_date",
]


def normalize_column_name(column):
    return (
        str(column)
        .strip()
        .lower()
        .replace("_", " ")
        .replace("-", " ")
    )


def find_column(df, possible_names):

    normalized_columns = {
        normalize_column_name(col): col
        for col in df.columns
    }

    for name in possible_names:

        normalized_name = normalize_column_name(name)

        if normalized_name in normalized_columns:
            return normalized_columns[normalized_name]

    return None


def clean_name(series):

    return (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )


def clean_phone(series):

    return (
        series
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )


def clean_number(series):

    return pd.to_numeric(
        series,
        errors="coerce"
    )


def clean_date(series):

    # Keep the date value from the input file
    return series


input_file = input("\nEnter the input Excel file name (e.g. file.xlsx): ").strip()

input_path = Path(input_file)

if not input_path.exists():

    print("\nERROR: File not found.")

    print(f"Make sure '{input_file}' exists in the same folder.")

    exit()


zero_sheet_name = input("\nEnter the name for the zero-amount sheet (e.g. Zero Amount): ").strip()

if not zero_sheet_name:

    zero_sheet_name = "Zero Amount"

# Excel sheet names cannot exceed 31 characters
if len(zero_sheet_name) > 31:

    print("\nSheet name is longer than 31 characters.")

    print("Excel only allows a maximum of 31 characters.")

    exit()


# ============================================================
# READ EXCEL
# ============================================================

try:

    print(f"\nReading Excel file: {input_file}")

    sheets = pd.read_excel(input_path,sheet_name=None)

    print(f"Found {len(sheets)} sheet(s).")

except Exception as e:

    print("\nERROR while reading Excel:")
    print(e)

    exit()


# ============================================================
# DATA STORAGE
# ============================================================

all_records = []

zero_amount_records = []

total_rows_read = 0
total_zero_amount = 0
total_invalid_amount = 0


# ============================================================
# PROCESS EACH SHEET
# ============================================================

for sheet_name, df in sheets.items():

    print("\n" + "-" * 65)
    print(f"PROCESSING SHEET: {sheet_name}")
    print("-" * 65)

    # Remove completely empty rows
    df = df.dropna(how="all")

    if df.empty:

        print("Skipped: Empty sheet.")
        continue

    total_rows_read += len(df)

    # --------------------------------------------------------
    # FIND COLUMNS
    # --------------------------------------------------------

    name_col = find_column(df,NAME_COLUMNS)

    phone_col = find_column(df,PHONE_COLUMNS)

    points_col = find_column(df,POINTS_COLUMNS)

    amount_col = find_column(df,AMOUNT_COLUMNS)

    date_col = find_column(df,DATE_COLUMNS)


    if name_col is None:

        print("Skipped: No Name column found.")

        continue

    if amount_col is None:

        print("Skipped: No Amount column found.")

        continue

    print(f"Name   : {name_col}")

    print(f"Phone  : {phone_col if phone_col else 'Not found'}")

    print(f"Points : {points_col if points_col else 'Not found'}")

    print(f"Amount : {amount_col}")

    print(f"Date   : {date_col if date_col else 'Not found'}")

    # --------------------------------------------------------
    # CREATE TEMP DATAFRAME
    # --------------------------------------------------------

    temp = pd.DataFrame()

    # Name
    temp["Name"] = clean_name(df[name_col])

    # Phone
    if phone_col:

        temp["Phone"] = clean_phone(df[phone_col])

    else:

        temp["Phone"] = ""

    # Points
    if points_col:

        temp["Points"] = clean_number(df[points_col])

    else:

        temp["Points"] = pd.NA

    # Amount
    temp["Amount"] = clean_number(df[amount_col])

    # Date
    if date_col:

        temp["Date"] = clean_date(df[date_col])

    else:

        temp["Date"] = ""

    # Station
    temp["Station"] = sheet_name

    # --------------------------------------------------------
    # REMOVE ROWS WITHOUT NAME
    # --------------------------------------------------------

    temp = temp[
        temp["Name"].str.strip() != ""
    ]

    # REMOVE INVALID AMOUNTS
    # Blank/non-numeric amounts are not treated as
    # zero. They are excluded from both output groups.

    invalid_amount_mask = (temp["Amount"].isna())

    invalid_count = invalid_amount_mask.sum()

    if invalid_count > 0:

        total_invalid_amount += invalid_count

        print(f"Excluded {invalid_count} row(s) with invalid/blank Amount.")

        temp = temp[~invalid_amount_mask]

    # SEPARATE ZERO AMOUNT RECORDS
    zero_mask = (temp["Amount"] == 0)

    zero_records = temp[zero_mask].copy()

    normal_records = temp[temp["Amount"] > 0].copy()

    # SAVE ZERO-AMOUNT RECORDS FOR SECOND SHEET

    if not zero_records.empty:

        zero_amount_records.append(zero_records)

        total_zero_amount += len(zero_records)

        print(f"Zero amount records: {len(zero_records)}")

    # SAVE NORMAL RECORDS FOR MAIN SHEET

    if not normal_records.empty:

        all_records.append(normal_records)

        print(f"Records kept for main sheet: {len(normal_records)}")


# CREATE MAIN DATAFRAME

if all_records:

    master = pd.concat(all_records,ignore_index=True)

else:

    master = pd.DataFrame(
        columns=[
            "Name",
            "Phone",
            "Points",
            "Amount",
            "Date",
            "Station"
        ]
    )

# CREATE ZERO-AMOUNT DATAFRAME

if zero_amount_records:

    zero_master = pd.concat(zero_amount_records,ignore_index=True)

else:

    zero_master = pd.DataFrame(
        columns=[
            "Name",
            "Phone",
            "Points",
            "Amount",
            "Date",
            "Station"
        ]
    )

# REMOVE EXACT DUPLICATES

master = master.drop_duplicates(keep="first")

zero_master = zero_master.drop_duplicates(keep="first")

# SET COLUMN ORDER

column_order = [
    "Date",
    "Name",
    "Phone",
    "Station",
    "Amount",
    "Points"
]

master = master[column_order]

zero_master = zero_master[column_order]

# SORT DATA

master = master.sort_values(by=["Name", "Date"],na_position="last").reset_index(drop=True)

zero_master = zero_master.sort_values(by=["Name", "Date"],na_position="last").reset_index(drop=True)

# SAVE BOTH SHEETS TO THE SAME EXCEL FILE

try:

    with pd.ExcelWriter(OUTPUT_FILE,engine="openpyxl") as writer:

        # Main records
        master.to_excel(writer,index=False,sheet_name="All Records")

        # Zero amount records
        zero_master.to_excel(writer,index=False,sheet_name=zero_sheet_name)

    print(f"\nSuccess! Excel file saved as:\n{Path(OUTPUT_FILE).resolve()}")

except PermissionError:

    print(f"\nERROR: Could not save '{OUTPUT_FILE}'.")

    print("The file may currently be open in Excel.")

    print("Close the file and run the script again.")

    exit()

except Exception as e:

    print("\nERROR while saving output file:")

    print(e)

    exit()


# FINAL SUMMARY

total_main_records = len(master)

total_zero_records = len(zero_master)

unique_names = master["Name"].nunique()

total_points = pd.to_numeric(master["Points"],errors="coerce").sum()

total_amount = pd.to_numeric(master["Amount"],errors="coerce").sum()


print("\n" + "=" * 65)
print("                    COMPLETE")
print("=" * 65)

print(f"Input file              : {input_file}")

print(f"Output file             : {OUTPUT_FILE}")

print(f"Main records            : {total_main_records}")

print(f"Zero amount records     : {total_zero_records}")

print(f"Unique names            : {unique_names}")

print(f"Total qualifying points : {total_points:g}")

print(f"Total qualifying amount : {total_amount:g}")

print(f"Zero amount sheet       : {zero_sheet_name}")

print("=" * 65)

print("\nDone!")
