import random
from datetime import date, timedelta
from openpyxl import Workbook

random.seed(7)

customers = [
    "Musterfirma Nord GmbH",
    "Beispiel Bau GmbH",
    "Test Logistik GmbH",
    "Demo Catering GmbH",
    "Muster Medien GmbH",
]

today = date.today()

wb = Workbook()
ws = wb.active
ws.append(["Rechnungsnummer", "Kunde", "Rechnungsdatum", "Fälligkeitsdatum", "Betrag (EUR)"])

number = 1001
overdue = 0

for i in range(25):
    customer = random.choice(customers)
    invoice_date = today - timedelta(days=random.randint(5, 90))
    due_date = invoice_date + timedelta(days=14)
    amount = round(random.uniform(80, 2500), 2)

    ws.append([number, customer, invoice_date, due_date, amount])
    number = number + 1

    if due_date < today:
        overdue = overdue + 1

for row in ws.iter_rows(min_row=2, min_col=3, max_col=4):
    for cell in row:
        cell.number_format = "DD.MM.YYYY"

for row in ws.iter_rows(min_row=2, min_col=5, max_col=5):
    for cell in row:
        cell.number_format = "0.00"

ws.column_dimensions["A"].width = 18
ws.column_dimensions["B"].width = 26
ws.column_dimensions["C"].width = 16
ws.column_dimensions["D"].width = 18
ws.column_dimensions["E"].width = 14

wb.save("offene_posten.xlsx")
print("Created: offene_posten.xlsx")
print("Invoices:", ws.max_row - 1)
print("Overdue:", overdue)