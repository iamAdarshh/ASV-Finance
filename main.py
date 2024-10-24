import PyPDF2
import pandas as pd
from datetime import datetime

# Update the path to use the uploaded file
excel_path = 'data/New ASV Finance 2024.xlsx'
template_path = 'data/refund-form.pdf'
output_path = 'data/filled/Filled_Kostenrueckerstattung_Formular.pdf'


def fill_refund_form(data, output_location):
    with open(template_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        writer = PyPDF2.PdfWriter()

        # Get the form fields
        form_fields = reader.get_fields()

        # Print all field names to help with debugging or to see the field names
        for field in form_fields:
            print(f"Field name: {field}")

        # Fill the form fields with the provided data
        for field_name, field_value in data.items():
            if field_name in form_fields:
                form_fields[field_name].update({
                    '/V': field_value,  # Set the value for the field
                })

        # Add all the pages from the original PDF to the writer
        for page_num in range(len(reader.pages)):
            writer.add_page(reader.pages[page_num])

        # Update the form field data in the output PDF
        writer.update_page_form_field_values(writer.pages[0], data)

        # Save the filled PDF to the output path
        with open(output_location, 'wb') as output_pdf:
            writer.write(output_pdf)

    print(f"Filled form saved to {output_location}")


if __name__ == '__main__':
    # Load all the relevant sheets
    event_df = pd.read_excel(excel_path, sheet_name='Event')
    purchase_df = pd.read_excel(excel_path, sheet_name='Purchase')
    purchase_item_df = pd.read_excel(excel_path, sheet_name='PurchaseItem')

    current_date = datetime.now()
    # Format the date as "day.month.year"
    formatted_date = current_date.strftime("%d.%m.%Y")

    # Iterate through each purchase entry
    for index, purchase_row in purchase_df.iterrows():
        print(f"Filling for Purchase ID: {purchase_row['Id']}")

        # Find the event linked to this purchase
        event_row = event_df[event_df['Id'] == purchase_row['EventId']].iloc[0]

        # Filter purchase items based on the PurchaseId
        items = purchase_item_df[purchase_item_df['PurchaseId'] == purchase_row['Id']]

        if items.empty:
            print(f"No items found for PurchaseId: {purchase_row['Id']}")
            continue  # Continue to process the next row

        # Prepare the data to fill in the PDF
        form_data = {
            'Name': "Lakshay Khanna",  # From the Purchase sheet
            'Mail': 'vorstand@asv.uni-paderborn.de',  # Placeholder email
            'Funktion': event_row['Name'],  # Event name from the Event sheet
            'Kontoinhaber': purchase_row['Name'] if pd.notna(purchase_row['Name']) else '',
            'IBAN': purchase_row['IBan'] if pd.notna(purchase_row['IBan']) else '',
            'BIC': purchase_row['BIC'] if pd.notna(purchase_row['BIC']) else '',
            'Ort Datum': f'Paderborn, {formatted_date}',
        }

        total = 0
        count = 0
        for item_index, item_row in items.iterrows():
            # Populate the items fields
            form_data[f'Beleg{count + 1}'] = str(item_row['BillNumber'])
            form_data[f'Begründung{count + 1}'] = item_row['Item']  # Item name
            form_data[f'Betrag{count + 1}'] = f"{item_row['Amount']:.2f}"  # Amount as string

            total += item_row['Amount']
            count += 1

        form_data['Erstattungsbetrag'] = f"{total:.2f}"

        # Define the output PDF path with unique naming
        output_pdf_path = f"data/filled_refund_form_{purchase_row['Id']}.pdf"

        # Fill the refund form
        fill_refund_form(form_data, output_pdf_path)

    print("All forms filled successfully.")
