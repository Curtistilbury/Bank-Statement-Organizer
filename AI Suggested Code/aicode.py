import csv
from openpyxl import Workbook

# Open the CSV file
def read_csv(file_path):
    data = []
    with open(file_path, mode='r') as file:
        reader = csv.DictReader(file)
        for row in reader:
            data.append(row)
    return data

def compile_data(file_paths):
    all_data = []
    for file_path in file_paths:
        data = read_csv(file_path)
        all_data.extend(data)
    return all_data

def detect_transfers(transactions):
    transfers = []
    non_transfers = []

    for t in transactions:
        description = t['description'].lower()
        if 'Transfer in' or 'Transfer out' in description:
            transfers.append(t)
        else:
            non_transfers.append(t)

    return transfers, non_transfers

def match_transfers(transfers):
    transfer_dict = {}
    matched_transfers = []

    for t in transfers:
        if t['amount'] not in transfer_dict:
            transfer_dict[t['amount']] = t
        else:
            matched_transfers.append((transfer_dict[t['amount']], t))
            del transfer_dict[t['amount']]

    return matched_transfers

def reconcile_data(transactions, matched_transfers):
    # Mark reconciled transfers
    reconciled_transactions = []
    matched_ids = set()

    for t in transactions:
        if t in [pair[0] for pair in matched_transfers] or t in [pair[1] for pair in matched_transfers]:
            matched_ids.add(t['description'])
        else:
            reconciled_transactions.append(t)

    return reconciled_transactions

def write_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "transactions"

    headers = ['date', 'description', 'amount', 'account type']
    ws.append(headers)

    for row in data:
        ws.append([row['date'], row['description'], row['amount'], row['transaction']])

    wb.save(output_file)

def write_budget_comparison(comparison_data, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Budget Comparison"

    headers = ['Category', 'Planned Amount', 'Spent Amount', 'Difference']
    ws.append(headers)

    for row in comparison_data:
        ws.append(row)

    wb.save(output_file)



# File paths to your CSV files
file_paths = ['chequing.csv', 'chequing_bills.csv']
output_file = 'budget_tracker.xlsx'

# Compile and reconcile data
compiled_data = compile_data(file_paths)
transfers, non_transfers = detect_transfers(compiled_data)
matched_transfers = match_transfers(transfers)
reconciled_data = reconcile_data(non_transfers, matched_transfers)
write_to_excel(reconciled_data, output_file)

print(compiled_data)

if (__name__ == '__main__'): pass