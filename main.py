
import csv
import xlsxwriter

# A class to define a transaction object
class Entry:
    def __init__(self, date, type, desc, amt, bal) -> None:
        self.date = date
        self.type = type
        self.description = desc
        self.amount = amt
        self.balance = bal

    def __str__(self) -> str:
        return f"date={self.date}&type={self.type}&description={self.description}&amount={self.amount}&balance={self.balance}"

headers = None
transactions = list()

# Import data from the csv file
with open("data.csv") as file:
    reader = csv.DictReader(file, delimiter=',')
    headers = reader.fieldnames
    for row in reader:
        transactions.append(
            Entry(
                row["date"],
                row["transaction"],
                row["description"],
                row["amount"],
                row["balance"],
            )
        )

# Set up Excel Workbook
workbook = xlsxwriter.Workbook("out.xlsx")
sheet = workbook.add_worksheet("Data")

x = 0
for header in headers:
    sheet.write(0, x, header.capitalize())
    x = x + 1

row = 1
numTransactions = 0
for trans in transactions:
    sheet.write(row, 0, trans.date)
    sheet.write(row, 1, trans.type)
    sheet.write(row, 2, trans.description)
    sheet.write(row, 3, trans.amount)
    sheet.write(row, 4, trans.balance)
    row = row + 1
    numTransactions = numTransactions + 1

print("Number of transactions processed: " + str(numTransactions))

workbook.close()
