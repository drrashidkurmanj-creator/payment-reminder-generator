# Payment Reminder Generator (Zahlungserinnerung)

A small Python tool for businesses: it reads an Excel list of **open invoices** (*Offene Posten*), finds the **overdue** ones and automatically creates a **German payment reminder letter (PDF) for each customer**, plus an **Excel summary** for the manager.

## What it does

- Reads open invoices from an Excel file
- Finds invoices past their due date and counts the days overdue
- Adds up the open amount for each customer
- Creates one PDF reminder letter (*Zahlungserinnerung*) per customer, listing all their overdue invoices
- Creates an Excel overview (`Mahnungsuebersicht.xlsx`) with totals per customer

## Example output

- Sample input: [offene_posten.xlsx](offene_posten.xlsx)
- Sample letter: [Zahlungserinnerung_Musterfirma_Nord_GmbH.pdf](Zahlungserinnerung_Musterfirma_Nord_GmbH.pdf)
- Sample overview: [Mahnungsuebersicht.xlsx](Mahnungsuebersicht.xlsx)

All customers and amounts are **made-up sample data**. The sender name, address and IBAN in the letters are placeholders.

## How to use

1. Install Python 3 and the libraries:

   ```
   pip install openpyxl fpdf2
   ```

2. Create sample data (optional):

   ```
   python offene_posten_daten.py
   ```

3. Open `zahlungserinnerung.py` and change the sender data at the top of the letter section (name, address, IBAN, payment deadline in days).
4. Run:

   ```
   python zahlungserinnerung.py
   ```

5. The letters are saved in a folder called `Zahlungserinnerungen`. The overview is saved as `Mahnungsuebersicht.xlsx`.

## Input format

The Excel file needs these columns, in this order:

`Rechnungsnummer | Kunde | Rechnungsdatum | Fälligkeitsdatum | Betrag (EUR)`

## Customization

Every company keeps its open invoices a little differently. The tool can be adapted to your own Excel file, your letter text and your payment deadlines.

## Important note

This is a demo tool. The wording of reminder letters, any fees or interest, and legal deadlines are the responsibility of the company and its legal or tax advisor. The tool does not replace legal or tax advice.

## Tech

Python, openpyxl, fpdf2

---

## Auf Deutsch

Ein kleines Python-Programm für Unternehmen: Es liest eine Excel-Liste **offener Posten**, findet die **überfälligen** Rechnungen und erstellt automatisch für jeden Kunden eine **Zahlungserinnerung als PDF** sowie eine **Excel-Übersicht** für die Geschäftsführung. Alle Beispieldaten sind frei erfunden. Der Wortlaut der Schreiben sowie Mahngebühren oder Zinsen liegen in der Verantwortung des Unternehmens bzw. seiner rechtlichen und steuerlichen Beratung.
