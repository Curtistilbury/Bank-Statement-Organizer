import csv
from openpyxl import Workbook

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
        description = t['Description'].lower()
        if 'transfer' in description:
            transfers.append(t)
        else:
            non_transfers.append(t)

    return transfers, non_transfers

def match_transfers(transfers):
    transfer_dict = {}
    matched_transfers = []

    for t in transfers:
        if t['Amount'] not in transfer_dict:
            transfer_dict[t['Amount']] = t
        else:
            matched_transfers.append((transfer_dict[t['Amount']], t))
            del transfer_dict[t['Amount']]

    return matched_transfers

def reconcile_data(transactions, matched_transfers):
    # Mark reconciled transfers
    reconciled_transactions = []
    matched_ids = set()

    for t in transactions:
        if t in [pair[0] for pair in matched_transfers] or t in [pair[1] for pair in matched_transfers]:
            matched_ids.add(t['Description'])
        else:
            reconciled_transactions.append(t)

    return reconciled_transactions

def write_to_excel(data, output_file):
    wb = Workbook()
    ws = wb.active
    ws.title = "Transactions"

    headers = ['Date', 'Description', 'Amount', 'Account Type']
    ws.append(headers)

    for row in data:
        ws.append([row['Date'], row['Description'], row['Amount'], row['Account Type']])

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

def budget_comparison(transactions, budget):
    category_spent = {}

    for t in transactions:
        category = t['Account Type']
        amount = float(t['Amount'])
        if category in category_spent:
            category_spent[category] += amount
        else:
            category_spent[category] = amount

    comparison = []
    for category, planned_amount in budget.items():
        spent = category_spent.get(category, 0)
        difference = planned_amount - spent
        comparison.append([category, planned_amount, spent, difference])

    return comparison

# File paths to your CSV files
file_paths = ['chequing.csv', 'savings.csv', 'credit_card.csv', 'chequing_bills.csv']
output_file = 'budget_tracker.xlsx'

# Compile and reconcile data
compiled_data = compile_data(file_paths)
transfers, non_transfers = detect_transfers(compiled_data)
matched_transfers = match_transfers(transfers)
reconciled_data = reconcile_data(non_transfers, matched_transfers)
write_to_excel(reconciled_data, output_file)

# Budget data (example)
planned_budget = {
    'Food': 300,
    'Rent': 1200,
    'Utilities': 150
}

# Compare and write budget comparison
comparison_data = budget_comparison(reconciled_data, planned_budget)
write_budget_comparison(comparison_data, output_file)
