from datetime import date
from openpyxl import load_workbook

wb = load_workbook("offene_posten.xlsx")
ws = wb.active

today = date.today()
overdue_list = []

for row in ws.iter_rows(min_row=2, values_only=True):
    number = row[0]
    customer = row[1]
    due = row[3].date()
    amount = row[4]

    days_late = (today - due).days

    if days_late > 0:
        overdue_list.append(
            {
                "number": number,
                "customer": customer,
                "days_late": days_late,
                "amount": amount,
                "due": due,
            }
        )

print("Überfällige Rechnungen:", len(overdue_list))
print()

for item in overdue_list:
    print(f"{item['number']}  {item['customer']}  {item['days_late']} Tage  {item['amount']:.2f} EUR")


    by_customer = {}

for item in overdue_list:
    name = item["customer"]

    if name not in by_customer:
        by_customer[name] = {"count": 0, "total": 0, "max_days": 0}

    by_customer[name]["count"] = by_customer[name]["count"] + 1
    by_customer[name]["total"] = by_customer[name]["total"] + item["amount"]

    if item["days_late"] > by_customer[name]["max_days"]:
        by_customer[name]["max_days"] = item["days_late"]

print()
print("Offen pro Kunde:")

for name in sorted(by_customer, key=lambda n: by_customer[n]["total"], reverse=True):
    data = by_customer[name]
    print(f"{name}: {data['count']} Rechnungen, {data['total']:.2f} EUR, längste Verspätung {data['max_days']} Tage")


    from datetime import timedelta
from fpdf import FPDF
from pathlib import Path

my_name = "Ihr Name / Ihre Firma"
my_address = "Musterstrasse 1, 12345 Musterstadt"
my_iban = "IBAN: DE00 0000 0000 0000 0000 00"
pay_days = 7


def line(pdf, text, size=11, bold=False, height=6):
    pdf.set_font("Helvetica", "B" if bold else "", size)
    pdf.cell(0, height, text, new_x="LMARGIN", new_y="NEXT")


def table_row(pdf, a, b, c, d, bold=False):
    pdf.set_font("Helvetica", "B" if bold else "", 10)
    pdf.cell(35, 8, a, border=1)
    pdf.cell(40, 8, b, border=1)
    pdf.cell(45, 8, c, border=1, align="R")
    pdf.cell(40, 8, d, border=1, align="R", new_x="LMARGIN", new_y="NEXT")


def make_letter(customer, items):
    pdf = FPDF()
    pdf.add_page()

    line(pdf, my_name)
    line(pdf, my_address)
    pdf.ln(10)
    line(pdf, customer)
    pdf.ln(8)
    line(pdf, f"Datum: {today.strftime('%d.%m.%Y')}")
    pdf.ln(4)
    line(pdf, "Zahlungserinnerung", 16, True, 10)
    pdf.ln(4)

    pay_until = today + timedelta(days=pay_days)
    text = (
        "Sehr geehrte Damen und Herren,\n\n"
        "bei der Durchsicht unserer Buchhaltung haben wir festgestellt, "
        "dass die folgenden Rechnungen noch nicht beglichen sind. "
        "Wir bitten Sie, den offenen Betrag bis zum "
        f"{pay_until.strftime('%d.%m.%Y')} auf das unten genannte Konto zu überweisen."
    )
    pdf.set_font("Helvetica", size=11)
    pdf.multi_cell(0, 6, text, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(4)

    table_row(pdf, "Rechnung", "Fällig am", "Tage überfällig", "Betrag (EUR)", bold=True)

    sum_open = 0
    for item in sorted(items, key=lambda i: i["days_late"], reverse=True):
        table_row(
            pdf,
            str(item["number"]),
            item["due"].strftime("%d.%m.%Y"),
            str(item["days_late"]),
            f"{item['amount']:.2f}",
        )
        sum_open = sum_open + item["amount"]

    table_row(pdf, "", "", "Summe", f"{sum_open:.2f}", bold=True)

    pdf.ln(6)
    line(pdf, my_iban)
    pdf.ln(4)
    pdf.set_font("Helvetica", size=11)
    pdf.multi_cell(
        0,
        6,
        "Sollte sich Ihre Zahlung mit diesem Schreiben überschnitten haben, "
        "betrachten Sie es bitte als gegenstandslos.\n\n"
        "Mit freundlichen Grüßen\n\n" + my_name,
        new_x="LMARGIN",
        new_y="NEXT",
    )

    folder = Path("Zahlungserinnerungen")
    folder.mkdir(exist_ok=True)

    file_name = "Zahlungserinnerung_" + customer.replace(" ", "_") + ".pdf"
    path = folder / file_name
    pdf.output(str(path))
    return str(path)


print()
print("Zahlungserinnerungen werden erstellt:")

for customer in sorted(by_customer, key=lambda n: by_customer[n]["total"], reverse=True):
    items = [i for i in overdue_list if i["customer"] == customer]
    print("Created:", make_letter(customer, items))


    from openpyxl import Workbook
from openpyxl.styles import Font

summary = Workbook()
ss = summary.active
ss.title = "Übersicht"

ss.append(["Kunde", "Offene Rechnungen", "Offener Betrag (EUR)", "Längste Verspätung (Tage)"])

for customer in sorted(by_customer, key=lambda n: by_customer[n]["total"], reverse=True):
    data = by_customer[customer]
    ss.append([customer, data["count"], data["total"], data["max_days"]])

last = ss.max_row
ss.append(["Gesamt", f"=SUM(B2:B{last})", f"=SUM(C2:C{last})", f"=MAX(D2:D{last})"])

for cell in ss[1]:
    cell.font = Font(bold=True)

for cell in ss[ss.max_row]:
    cell.font = Font(bold=True)

for row in ss.iter_rows(min_row=2, min_col=3, max_col=3):
    for cell in row:
        cell.number_format = "0.00"

ss.column_dimensions["A"].width = 28
ss.column_dimensions["B"].width = 20
ss.column_dimensions["C"].width = 22
ss.column_dimensions["D"].width = 26

summary.save("Mahnungsuebersicht.xlsx")
print("Created: Mahnungsuebersicht.xlsx")