import os
import glob
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# Constants and configurations
ACCOUNTS = ["wea", "tang", "sim", "td"]
ACCOUNT_TYPES = ["cheq", "bills", "credit", "savings"]
MONTHS = [f"{i:02}" for i in range(1, 13)]
FOLDER_PATH = "./statements"
CURRENT_YEAR = datetime.now().year
REQUIRED_COLUMNS = ['Date', 'Transaction', 'Description', 'Amount']

# Column mappings
COLUMN_MAPPING = {
    'wea': {'date': 'Date', 'transaction': 'Transaction', 'description': 'Description', 'amount': 'Amount'},
    'tang': {'transaction date': 'Date', 'date': 'Date', 'transaction': 'Transaction', 'name': 'Description', 'amount': 'Amount'},
    'sim': {'date': 'Date', 'transaction': 'Description', 'funds out': 'Amount1', 'funds in': 'Amount2'}
}

def is_valid_year(year):
    return year.isdigit() and 1900 <= int(year) <= CURRENT_YEAR

def is_valid_month(month):
    return month in MONTHS

def apply_column_mapping(df, account):
    if account in COLUMN_MAPPING:
        mapping = COLUMN_MAPPING[account]
        df.columns = df.columns.str.lower()
        df.rename(columns={k.lower(): v for k, v in mapping.items()}, inplace=True)
        df = df[[col for col in df.columns if col in mapping.values()]]
        for col in mapping.values():
            if col not in df.columns:
                df[col] = pd.NA
    return df

def validate_and_reorder_columns(df):
    missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_columns:
        print(f"Warning: Missing columns: {missing_columns}")
    df = df.reindex(columns=REQUIRED_COLUMNS, fill_value=pd.NA)
    return df

def adjust_columns(df, column_names, header_skip=False):
    if header_skip:
        df = df.iloc[1:].copy()
    df.columns = column_names
    df['Transaction'] = ""
    df['Amount1'] = df['Amount1'].apply(lambda x: abs(float(x)) if pd.notnull(x) else 0)
    df['Amount2'] = df['Amount2'].apply(lambda x: -abs(float(x)) if pd.notnull(x) else 0)
    df['Amount'] = df['Amount1'].fillna(0) + df['Amount2'].fillna(0)
    df.drop(columns=['Amount1', 'Amount2'], inplace=True)
    return df

def process_file(file_path, account, account_type):
    if account == "td":
        df = pd.read_csv(file_path, header=None)
        if df.shape[1] < 5:
            raise ValueError("TD account statement does not have at least 5 columns.")
        df = adjust_columns(df, ['Date', 'Description', 'Amount1', 'Amount2', 'Balance'], header_skip=False)
    elif account == "sim":
        df = pd.read_csv(file_path)
        if df.shape[1] != 4:
            raise ValueError("SIM account statement does not have exactly 4 columns.")
        df = adjust_columns(df, ['Date', 'Description', 'Amount1', 'Amount2'], header_skip=True)
    else:
        df = pd.read_csv(file_path)
    df = apply_column_mapping(df, account)
    df = validate_and_reorder_columns(df)
    return df

def find_and_compile_csv_files():
    all_data = []
    found_files = []
    invalid_files = []
    non_csv_files = []

    all_files = glob.glob(os.path.join(FOLDER_PATH, "*"))
    for file_path in all_files:
        file_name = os.path.basename(file_path)
        if not file_name.endswith(".csv"):
            non_csv_files.append(file_name)
            continue

        parts = file_name.split('_')
        if len(parts) != 4:
            invalid_files.append(file_name)
            continue

        year, month, account, account_type_with_ext = parts
        account_type = account_type_with_ext.replace(".csv", "")
        if not (is_valid_year(year) and is_valid_month(month) and account in ACCOUNTS and account_type in ACCOUNT_TYPES):
            invalid_files.append(file_name)
            continue

        try:
            df = process_file(file_path, account, account_type)
            df['Year'] = year
            df['Account'] = account
            df['Type'] = account_type
            df['Amount'] = pd.to_numeric(df['Amount'], errors='coerce').round(2)
            df['Date'] = pd.to_datetime(df['Date'], errors='coerce')
            df = df[['Date', 'Year', 'Account', 'Type', 'Transaction', 'Description', 'Amount']]
            all_data.append(df)
            found_files.append(file_name)
        except Exception as e:
            print(f"Error processing {file_name}: {e}")
            invalid_files.append(file_name)

    if not all_data:
        print("No valid data found. Script ended.")
        return

    combined_df = pd.concat(all_data, ignore_index=True)
    combined_df.sort_values(by='Date', inplace=True)

    output_path = os.path.join(".", "combined_statements.xlsx")
    combined_df.to_excel(output_path, index=False, sheet_name='Sheet1')

    workbook = load_workbook(output_path)
    worksheet = workbook.active
    table = Table(displayName="CombinedDataTable", ref=f"A1:G{len(combined_df) + 1}")
    style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False, showLastColumn=False,
                           showRowStripes=True, showColumnStripes=True)
    table.tableStyleInfo = style
    worksheet.add_table(table)
    workbook.save(output_path)

    print(f"\nCombined data written to {output_path} successfully.")

print("Running the CSV compilation script...")
find_and_compile_csv_files()
print("Script execution finished.")