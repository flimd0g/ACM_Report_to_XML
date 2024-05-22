HTML to Excel Data Extraction and Insertion.

This program is specifically designed for helping to keep track of custom ACM files but could be easily adjusted for a similar use-case.

Overview

This project provides a Python application that extracts specific data points from an HTML file and inserts the data into a specified Excel file. The application includes a graphical user interface (GUI) to facilitate file selection.

Features

- Extracts specified data points from an HTML file.
- Inserts the extracted data into specified columns in an Excel file.
- Finds the lowest ID number with an otherwise empty row in the Excel sheet and populates that row with the extracted data.
- Provides a simple GUI for file selection.

Requirements

- Python 3.7 or later
- Required Python packages:
  - tkinter
  - beautifulsoup4
  - openpyxl

Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/html-to-excel.git
   cd html-to-excel
   ```

2. **Install required packages**:
   ```bash
   pip install beautifulsoup4 openpyxl
   ```

Usage

1. **Run the application**:
   ```bash
   python main.py
   ```

2. **Select an HTML file**:
   - Click the "Select HTML File" button.
   - Choose the HTML file from which you want to extract data.

3. **Check the console output**:
   - The console will display debug information about the extracted data and the headers found in the Excel file.
   - Ensure that the headers in the Excel file match the expected headers.

Configuration

- **Excel File Path**:
  - Update the `excel_path` variable in the `process_file` function to the correct path of your Excel file.

  ```python
  def process_file(file_path):
      extracted_values = parse_html(file_path)
      excel_path = '/path/to/your/excel/file.xlsx'  # Update this with your actual file path
      update_excel(extracted_values, excel_path)
  ```

- **Keys to Extract**:
  - Modify the `keys_to_extract` list in the `parse_html` function to include the data points you want to extract from the HTML file.

  ```python
  keys_to_extract = [
      'ACM version', 'ACM diagnosis version', 'ACM VIN', 'ACM serial number',
      'ACM hardware part number', 'ACM certification', 'ACM hardware version'
  ]
  ```

- **Header Mapping**:
  - Ensure the `header_mapping` dictionary in the `update_excel` function matches the actual column headers in your Excel file.

  ```python
  header_mapping = {
      'ACM version': 'Version',
      'ACM diagnosis version': 'Diagnosis Version',
      'ACM VIN': 'Vin',
      'ACM serial number': 'Serial Number',
      'ACM hardware part number': 'Part Number',
      'ACM certification': 'Certification',
      'ACM hardware version': 'Hardware Version'
  }
  ```

Example

### HTML File Sample
```html
<tr><td>ACM version</td><td>5.57.0.0</td></tr>
<tr><td>ACM diagnosis version</td><td>000E21</td></tr>
<tr><td>ACM VIN</td><td>WDB9634232L972473</td></tr>
<tr><td>ACM serial number</td><td>01A27D29</td></tr>
<tr><td>ACM hardware part number</td><td>0004464354002</td></tr>
<tr><td>ACM certification</td><td>OM471-6-1-A-08</td></tr>
<tr><td>ACM hardware version</td><td>14/11.00</td></tr>
```

Excel File Sample
| ID Number | Hardware Class | Part Number | Denoting Numbers | Label Variant | Hardware Version | Certification | Version | Diagnosis Version | Serial Number | Vin |
|-----------|----------------|-------------|------------------|---------------|------------------|---------------|---------|--------------------|----------------|-----|
| 1         |                |             |                  |               |                  |               |         |                    |                |     |

Debugging

- If the application outputs "Header row not found in the Excel sheet", ensure that the column headers in the Excel file match the keys in the `header_mapping` dictionary.
- Use the debug prints in the console to verify that the headers are correctly identified.

License

This project is licensed under the MIT License - see the LICENSE file for details.

Acknowledgments

- [BeautifulSoup](https://www.crummy.com/software/BeautifulSoup/)
- [openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [Tkinter](https://docs.python.org/3/library/tkinter.html)
