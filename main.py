import tkinter as tk
from tkinter import filedialog
from bs4 import BeautifulSoup
import openpyxl
import os


def select_file():
    file_path = filedialog.askopenfilename(filetypes=[("HTML files", "*.html")])
    if file_path:
        process_file(file_path)


def parse_html(file_path):
    with open(file_path, 'r', encoding='utf-8') as file:
        soup = BeautifulSoup(file, 'html.parser')

    # Define the keys to extract
    keys_to_extract = [
        'ACM version', 'ACM diagnosis version', 'ACM VIN', 'ACM serial number',
        'ACM hardware part number', 'ACM certification', 'ACM hardware version'
    ]

    # Extract the values
    extracted_values = {key: None for key in keys_to_extract}

    rows = soup.find_all('tr')
    for row in rows:
        cells = row.find_all('td')
        if len(cells) == 2:
            key = cells[0].get_text(strip=True)
            if key in extracted_values:
                extracted_values[key] = cells[1].get_text(strip=True)

    # Debug print to check extracted data
    for key, value in extracted_values.items():
        print(f"{key}: {value}")

    return extracted_values


def update_excel(extracted_values, excel_path):
    if not os.path.isfile(excel_path):
        print(f"Excel file not found: {excel_path}")
        return

    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active  # Or specify the sheet name: wb['SheetName']

    # Header mapping: Map the keys to the actual column headers in the Excel file
    header_mapping = {
        'ACM version': 'Version',
        'ACM diagnosis version': 'Diagnosis Version',
        'ACM VIN': 'Vin',
        'ACM serial number': 'Serial Number',
        'ACM hardware part number': 'Part Number',
        'ACM certification': 'Certification',
        'ACM hardware version': 'Hardware Version'
    }

    # Identify the header row
    header_row_index = None
    for row in ws.iter_rows(min_row=1, max_row=10):  # Adjust range if headers are farther down
        headers = {cell.value: cell.column for cell in row if cell.value}
        print(f"Headers found in row {row[0].row}: {headers}")  # Debug print
        if set(header_mapping.values()).issubset(headers.keys()):
            header_row_index = row[0].row
            break

    if not header_row_index:
        print("Header row not found in the Excel sheet")
        return

    # Find the correct column indices based on headers
    headers = {cell.value: cell.column for cell in ws[header_row_index]}
    print(f"Headers and their columns: {headers}")  # Debug print

    # Check if all required columns are present
    for key in extracted_values.keys():
        if header_mapping[key] not in headers:
            print(f"Column for '{key}' not found in the Excel sheet")
            return

    # Find the lowest ID number with an otherwise empty row
    target_row = None
    for row in ws.iter_rows(min_row=header_row_index + 1):
        id_cell = row[0]  # Assuming ID Number is in the first column
        if id_cell.value is not None:
            # Check if the rest of the cells in this row are empty
            if all(cell.value is None for cell in row if cell.column != 1):  # Exclude ID Number column
                target_row = id_cell.row
                break

    if target_row is None:
        print("No suitable row found for updating")
        return

    # Insert values into the identified row
    for key, value in extracted_values.items():
        ws.cell(row=target_row, column=headers[header_mapping[key]], value=value)

    wb.save(excel_path)
    wb.close()
    print(f"Excel file '{excel_path}' updated successfully.")


def process_file(file_path):
    extracted_values = parse_html(file_path)
    excel_path = '/Users/finleybrown/Documents/Work Docs CVE/Copies/Unit Assessment ACM--PC copy.xlsx'  # Update this with your actual file path
    update_excel(extracted_values, excel_path)


root = tk.Tk()
root.title("HTML to Excel")

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(padx=10, pady=10)

select_button = tk.Button(frame, text="Select HTML File", command=select_file)
select_button.pack()

root.mainloop()


