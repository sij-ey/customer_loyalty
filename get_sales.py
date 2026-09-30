import pandas as pd


# ============================================================
# FILE SETTINGS
# ============================================================

main_file = input(
    "Enter the main Excel file name (e.g., file.xlsx): "
).strip()

comparison_file = input(
    "Enter the comparison Excel file name (e.g., file.xlsx): "
).strip()

output_file = input(
    "Enter the output Excel file name (e.g., file.xlsx): "
).strip()


# ============================================================
# COLUMN MAPPING
# ============================================================
#
# LEFT side  = column name in MAIN file
# RIGHT side = corresponding column name in COMPARISON file
#
# CHANGE THESE TO MATCH YOUR ACTUAL FILES.
#

column_mapping = {
    "Date": "Date",
    "Station":"Station",
    "Customer": "Name",
    "Customer phone": "Phone",
    "Amount paid": "Amount(KES)"
}


# ============================================================
# MAIN MATCHING COLUMN
# ============================================================

# Column from MAIN file used for matching
main_date_column = "Date"

# Column from COMPARISON file used for matching
comparison_date_column = "Date"


# ============================================================
# READ FILES
# ============================================================

main_df = pd.read_excel(main_file)

comparison_df = pd.read_excel(comparison_file)


# ============================================================
# CHECK MATCHING COLUMNS
# ============================================================

if main_date_column not in main_df.columns:
    raise ValueError(
        f"'{main_date_column}' was not found in the MAIN file.\n"
        f"Available columns: {list(main_df.columns)}"
    )

if comparison_date_column not in comparison_df.columns:
    raise ValueError(
        f"'{comparison_date_column}' was not found in the "
        f"COMPARISON file.\n"
        f"Available columns: {list(comparison_df.columns)}"
    )


# ============================================================
# CONVERT DATE/TIME TO DATETIME
# ============================================================

main_df["_match_date"] = pd.to_datetime(
    main_df[main_date_column],
    errors="coerce"
)

comparison_df["_match_date"] = pd.to_datetime(
    comparison_df[comparison_date_column],
    errors="coerce"
)


# ============================================================
# GET DATES FROM MAIN FILE
# ============================================================

main_dates = set(
    main_df["_match_date"].dropna()
)


# ============================================================
# FIND RECORDS EXISTING ONLY IN COMPARISON FILE
# ============================================================

missing_in_main = comparison_df[
    ~comparison_df["_match_date"].isin(main_dates)
].copy()


# ============================================================
# CONVERT COMPARISON COLUMNS TO MAIN COLUMN NAMES
# ============================================================

# Start an empty DataFrame with exactly the same
# columns as the MAIN file.

new_records = pd.DataFrame(
    columns=main_df.columns
)


# ============================================================
# COPY DATA FROM COMPARISON INTO MAIN STRUCTURE
# ============================================================

for main_column, comparison_column in column_mapping.items():

    if main_column not in main_df.columns:
        print(
            f"WARNING: '{main_column}' does not exist "
            f"in the main file."
        )
        continue

    if comparison_column not in comparison_df.columns:
        print(
            f"WARNING: '{comparison_column}' does not exist "
            f"in the comparison file."
        )
        continue

    new_records[main_column] = (
        missing_in_main[comparison_column].values
    )


# ============================================================
# HANDLE THE TEMPORARY MATCHING COLUMN
# ============================================================

new_records["_match_date"] = (
    missing_in_main["_match_date"].values
)


# ============================================================
# COMBINE MAIN + NEW RECORDS
# ============================================================

# Remove temporary column from main records
main_output = main_df.drop(
    columns=["_match_date"]
).copy()

# Remove temporary column from new records
new_records = new_records.drop(
    columns=["_match_date"]
).copy()


# Ensure new records have exactly the same
# column order as the main file.

new_records = new_records[
    main_output.columns
]


# Combine
all_records = pd.concat(
    [
        main_output,
        new_records
    ],
    ignore_index=True
)


# ============================================================
# MAKE PHONE COLUMN TEXT
# ============================================================

phone_column = "Phone"


def format_phone(phone):

    if pd.isna(phone):
        return ""

    if isinstance(phone, float):

        if phone.is_integer():
            return str(int(phone))

    return str(phone).strip()


if phone_column in all_records.columns:

    all_records[phone_column] = (
        all_records[phone_column]
        .apply(format_phone)
    )


# ============================================================
# PREPARE MISSING-IN-MAIN SHEET
# ============================================================

# Rename comparison columns to MAIN column names.

missing_output = pd.DataFrame(
    columns=main_output.columns
)

for main_column, comparison_column in column_mapping.items():

    if (
        main_column in missing_output.columns
        and comparison_column in missing_in_main.columns
    ):

        missing_output[main_column] = (
            missing_in_main[comparison_column].values
        )


# Format phone
if phone_column in missing_output.columns:

    missing_output[phone_column] = (
        missing_output[phone_column]
        .apply(format_phone)
    )


# ============================================================
# CREATE OUTPUT EXCEL FILE
# ============================================================

with pd.ExcelWriter(
    output_file,
    engine="openpyxl"
) as writer:

    # --------------------------------------------------------
    # SHEET 1
    # --------------------------------------------------------

    all_records.to_excel(
        writer,
        sheet_name="All Records",
        index=False
    )

    # --------------------------------------------------------
    # SHEET 2
    # --------------------------------------------------------

    missing_output.to_excel(
        writer,
        sheet_name="Missing in Main",
        index=False
    )


    # ========================================================
    # FORMAT PHONE COLUMN AS TEXT
    # ========================================================

    if phone_column in all_records.columns:

        ws1 = writer.sheets["All Records"]

        phone_col_number = (
            all_records.columns.get_loc(phone_column) + 1
        )

        for row in range(2, ws1.max_row + 1):

            ws1.cell(
                row=row,
                column=phone_col_number
            ).number_format = "@"


    if phone_column in missing_output.columns:

        ws2 = writer.sheets["Missing in Main"]

        phone_col_number = (
            missing_output.columns.get_loc(phone_column) + 1
        )

        for row in range(2, ws2.max_row + 1):

            ws2.cell(
                row=row,
                column=phone_col_number
            ).number_format = "@"


# ============================================================
# SUMMARY
# ============================================================

print()
print("========================================")
print("COMPARISON COMPLETED")
print("========================================")

print(
    f"Records in main file:             "
    f"{len(main_output)}"
)

print(
    f"New records from comparison:      "
    f"{len(new_records)}"
)

print(
    f"Total records in All Records:     "
    f"{len(all_records)}"
)

print(
    f"Records missing from main:        "
    f"{len(missing_output)}"
)

print()
print(f"Output file: {output_file}")
