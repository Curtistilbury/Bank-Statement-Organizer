import os
import glob
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

# Lists of possible values
accounts = ["wea", "tan", "sim", "td"]
types = ["cheq", "bills", "credit"]
months = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]

# The folder where the CSV files are stored, within the Repl environment
folder_path = "./statements"
# Define the current year for validation
current_year = datetime.now().year

# Function to validate the year and month from the filename
def is_valid_year(year):
    return year.isdigit() and 1900 <= int(year) <= current_year

def is_valid_month(month):
    return month in months

def find_and_compile_csv_files():
    all_data = []  # This list will store DataFrames for each valid file

    found_files = []  # List to keep track of found files
    invalid_files = []  # List to keep track of invalid files
    non_csv_files = []  # List to keep track of non-CSV files

    print("Starting to search for files...")

    try:
        # Look for all files in the folder
        all_files = glob.glob(os.path.join(folder_path, "*"))

        for file_path in all_files:
            file_name = os.path.basename(file_path)

            # Skip non-CSV files
            if not file_name.endswith(".csv"):
                non_csv_files.append(file_name)
                continue

            # Check the filename structure
            parts = file_name.split('_')
            if len(parts) == 4:
                year, month, account, account_type_with_ext = parts
                account_type = account_type_with_ext.replace(".csv", "")

                # Validate extracted parts
                if is_valid_year(year) and is_valid_month(month) and account in accounts and account_type in types:
                    found_files.append(file_name)

                    # Read the CSV file into a DataFrame
                    df = pd.read_csv(file_path)

                    # Add metadata columns
                    df['date'] = pd.to_datetime(df['date'], errors='coerce')  # Convert to datetime
                    df['year'] = year
                    df['account'] = account
                    df['type'] = account_type

                    # Ensure necessary columns are present
                    necessary_columns = ['transaction', 'description', 'amount', 'balance']
                    for col in necessary_columns:
                        if col not in df.columns:
                            df[col] = pd.NA

                    # Format Amount and Balance columns as numeric with two decimal places
                    df['amount'] = pd.to_numeric(df['amount'], errors='coerce').round(2)
                    df['balance'] = pd.to_numeric(df['balance'], errors='coerce').round(2)
                    
                    # Map lowercase columns to capitalized columns
                    df.rename(columns={
                        'date': 'Date',
                        'year': 'Year',
                        'account': 'Account',
                        'type': 'Type',
                        'transaction': 'Transaction',
                        'description': 'Description',
                        'amount': 'Amount',
                        'balance': 'Balance'
                    }, inplace=True)

                    # Order columns as specified
                    df = df[['Date', 'Year', 'Account', 'Type', 'Transaction', 'Description', 'Amount', 'Balance']]

                    # Add this DataFrame to the list of all data
                    all_data.append(df)
                else:
                    invalid_files.append(file_name)
            else:
                invalid_files.append(file_name)

        print("Finished searching for files.")
        print("Valid files found and processed:")
        for file in found_files:
            print(f"- {file}")

        print("\nInvalid files (incorrect structure):")
        for file in invalid_files:
            print(f"- {file}")

        print("\nNon-CSV files (skipped):")
        for file in non_csv_files:
            print(f"- {file}")

        if not all_data:
            print("\nNo data found to combine. Script ended because no data was collected.")
            return

        # Combine all the DataFrames in the list into one large DataFrame
        combined_df = pd.concat(all_data, ignore_index=True)
        print("\nData successfully combined.")

        # Define the path for the output Excel file outside the statements folder
        output_path = os.path.join(".", "combined_statements.xlsx")

        # Write the combined DataFrame to an Excel file
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            combined_df.to_excel(writer, index=False, sheet_name='Sheet1')

            # Load the workbook and select the active worksheet
            writer.save()
            workbook = load_workbook(output_path)
            worksheet = workbook.active

            # Define the range and create a table
            table = Table(displayName="CombinedDataTable", ref=f"A1:H{len(combined_df) + 1}")

            # Add a default style with stripes
            style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                                   showLastColumn=False, showRowStripes=True, showColumnStripes=True)
            table.tableStyleInfo = style
            worksheet.add_table(table)

            # Save the workbook
            workbook.save(output_path)

        print(f"Attempted to write combined data to {output_path}")

        if os.path.exists(output_path):
            print(f"\nCombined data written to {output_path} successfully.")
        else:
            print(f"\nFailed to write combined data to {output_path}. Script ended because writing to Excel failed.")
            return

    except Exception as e:
        print(f"An error occurred: {e}")
        print("Script ended due to an error.")

# Run the function to find, read, and compile CSV file data
print("Running the CSV compilation script...")
find_and_compile_csv_files()
print("Script execution finished.")