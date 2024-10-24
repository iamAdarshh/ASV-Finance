# Refund Form Automation Script

This script automates filling in a PDF refund form using data from an Excel file. It reads data from multiple sheets, processes purchase details, fills out the PDF, and generates a completed refund form for each purchase entry. The filled PDF forms are saved as non-editable files, except for the signature field.

## Features

- Reads data from multiple Excel sheets: **Event**, **Purchase**, and **PurchaseItem**.
- Fills out predefined fields in a PDF refund form template.
- Supports multiple purchase items per refund.
- Automatically calculates the total refund amount.
- Makes the filled PDF non-editable, except for the signature field.
- Saves the filled forms with unique names for each purchase entry.

## Requirements

- Python 3.x
- `PyPDF2` for PDF manipulation
- `pandas` for reading and processing Excel files

### Install required libraries:

```bash
pip install -r requirements.txt
```

## How to Use

1. **Prepare the Excel file**:  
   A sample Excel file named **"New ASV Finance 2024_Sample.xlsx"** is provided to help you set up your data. You can use this as a reference to create your own Excel sheet named **"New ASV Finance 2024.xlsx"**. Ensure the following sheets are present:
   - **Event**: Contains information about events.
   - **Purchase**: Contains purchase-related details such as the purchaser’s name, IBAN, BIC, etc.
   - **PurchaseItem**: Contains the items related to each purchase (e.g., Bill Number, Amount, etc.).

2. **Prepare the PDF template**:  
   Make sure you have a PDF form template (`refund-form.pdf`) that has predefined fields that match the field names used in the script.

3. **Modify paths**:  
   Ensure the paths to the Excel file and the PDF form template are correct:
   - `excel_path`: Path to your Excel file (e.g., `"New ASV Finance 2024.xlsx"`).
   - `template_path`: Path to the PDF form template.
   - `output_path`: Directory where filled forms will be saved.
  
4. **Run the script**:  
   Simply run the script with:

   ```bash
   python main.py
   ```

   This will generate filled and non-editable refund forms for each purchase entry and save them in the output directory. The forms will be saved with unique filenames based on the purchase IDs.

### Example Output:
   For a purchase with `Purchase ID = 123`, the filled PDF will be saved as:

   ```
   data/filled_refund_form_123.pdf
   ```

5. **Signature**:  
   After generating the PDFs, only the signature field will remain editable. You can sign the forms manually or using a digital signature tool.

## Script Breakdown

- **`fill_refund_form(data, output_location)`**:  
   This function takes the form data and fills the PDF template using `PyPDF2`. It then flattens the fields to make them non-editable, except for the signature field.

- **Excel Sheet Structure**:
   - **Event Sheet**: Contains event details. The event name will be used to fill the "Funktion" field in the PDF.
   - **Purchase Sheet**: Contains purchase details such as the purchaser's name, IBAN, BIC, etc.
   - **PurchaseItem Sheet**: Contains individual item details (e.g., Bill Number, Amount) related to each purchase.

- **Data Mapping**:  
   The script extracts specific fields from the Excel sheets and maps them to the PDF form fields:
   - `Name`: Purchaser's name
   - `IBAN`: Purchaser's IBAN (left blank if not available)
   - `BIC`: Purchaser's BIC (left blank if not available)
   - `BelegX`, `BegründungX`, `BetragX`: The bill number, item name, and amount for each purchase item, where `X` is the item number.
   - `Erstattungsbetrag`: The total refund amount calculated from the purchase items.

## Customization

- **Adjusting the PDF fields**:  
   If your PDF form fields have different names, modify the `form_data` dictionary in the script to match the correct field names.

- **Signature Field**:  
   If you need to modify how the signature field behaves, you can manually adjust the form template or handle it separately using tools like `reportlab`.

## Troubleshooting

- **Field Names Mismatch**:  
   Ensure that the field names in the PDF template match the ones expected by the script. You can print the field names with this section in the script:
   ```python
   for field in form_fields:
       print(f"Field name: {field}")
   ```

- **No items found**:  
   If the script prints "No items found for PurchaseId", check that the `PurchaseItem` sheet contains matching `PurchaseId` entries for the `Purchase` sheet.

## Contributing
In case of any changes or improvements, feel free to make changes and raise a PR! Contributions are always welcome to improve the project.

## License

This project is licensed under the MIT License.

---
