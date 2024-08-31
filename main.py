import os
import glob
import pandas as pd

# Lists of possible values
accounts = ["wea", "tan", "sim", "td"]
types = ["cheq", "bills", "credit"]
months = ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]

# The folder where the CSV files are stored, within the Repl environment
folder_path = "./statements"

# Function to find CSV files and compile their data
def find_and_compile_csv_files():
    all_data = []  # This list will store DataFrames for each file

    found_files = []  # List to keep track of found files
    not_found_patterns = []  # List to keep track of patterns with no matching files

    print("Starting to search for files...")

    try:
        # Step 1: Loop through possible values to find files
        for account in accounts:
            for account_type in types:
                for month in months:
                    # Construct the pattern to match filenames with any year
                    pattern = f"{folder_path}/*_{month}_{account}_{account_type}.csv"
                    matching_files = glob.glob(pattern)

                    # If files are found
                    if matching_files:
                        for file_path in matching_files:
                            file_name = os.path.basename(file_path)
                            found_files.append(file_name)

                            # Extract the year from the filename (first part before the first "_")
                            year = file_name.split('_')[0]

                            # Read the CSV file into a DataFrame
                            df = pd.read_csv(file_path)

                            # Add columns for the metadata (year, month, account, type)
                            df['year'] = year
                            df['month'] = month
                            df['account'] = account
                            df['type'] = account_type

                            # Add this DataFrame to the list of all data
                            all_data.append(df)
                    else:
                        not_found_patterns.append(pattern)
        
        print("Finished searching for files.")
        print("Files found:")
        for file in found_files:
            print(f"- {file}")

        print("\nPatterns with no matching files:")
        for pattern in not_found_patterns:
            print(f"- {pattern}")

        # Check if any data was collected
        if not all_data:
            print("No data found to combine.")
            print("Script ended because no data was collected.")
            return

        # Step 2: Combine all the DataFrames in the list into one large DataFrame
        combined_df = pd.concat(all_data, ignore_index=True)
        print("Data successfully combined.")

        # Step 3: Define the path for the output Excel file
        output_path = os.path.join(folder_path, "combined_statements.xlsx")

        # Write the combined DataFrame to an Excel file
        combined_df.to_excel(output_path, index=False)
        print(f"Attempted to write combined data to {output_path}")

        # Check if the Excel file was created successfully and print the result
        if os.path.exists(output_path):
            print(f"\nCombined data written to {output_path} successfully.")
        else:
            print(f"\nFailed to write combined data to {output_path}.")
            print("Script ended because writing to Excel failed.")
            return

    except Exception as e:
        print(f"An error occurred: {e}")
        print("Script ended due to an error.")

# Run the function to find, read, and compile CSV file data
print("Running the CSV compilation script...")
find_and_compile_csv_files()
print("Script execution finished.")