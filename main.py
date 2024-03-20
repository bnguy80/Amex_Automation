import pdfplumber
import pdf2image
import re
import dateparser
from datetime import datetime
import xlwings as xw
import pytesseract
import nbformat
from nbconvert.preprocessors import ExecutePreprocessor

# https://tesseract-ocr.github.io/tessdoc/Installation.html
# https://pypi.org/project/pytesseract/
# https://pypi.org/project/pdf2image/

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'
poppler_path = "C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin"


def list_files_macro(wb):
    macro_vba = wb.app.macro('ListFilesInSpecificFolder')
    macro_vba()  # calling VBA function


def extract_pdf_invoice_totals(wb, ws):
    patterns = [
        r"Grand Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total \(in USD\)\s+\$(\d[\d,]*\.\d{2})",
        r"Total amount due(?: \(USD\))?:?\s+\$?\S?(\d[\d,]*\.\d{2})",
        r"Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total:\s+(\d[\d,]*\.\d{2})(?:\s+USD)?",
        r"New charges\s+\$(\d[\d,]*\.\d{2})",
        r"Invoice Total\s+\$(\d[\d,]*\.\d{2})",
    ]

    # Calculate the row number where the file paths end
    last_row = ws.range('B' + str(ws.cells.last_cell.row)).end('up').row

    # Iterate over the PDF file paths listed in the Excel sheet
    for i in range(8, last_row + 1):
        pdf_path = ws.range(f"B{i}").value
        total_found = False

        # Open the PDF and search for the patterns
        with pdfplumber.open(pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)
            for pattern in patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    # Write the matched pattern's value to the Excel sheet
                    total = float(match.group(1).replace(',', ''))
                    ws.range(f"C{i}").value = total
                    total_found = True
                    break

        if not total_found:
            ws.range(f"C{i}").value = "Total Not Found"

    # Save the workbook
    wb.save()


def extract_ocr_invoice_totals(wb, ws):
    patterns = [
        r"Grand Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total amount due(?: \(USD\))?:?\s+\$?\S?(\d[\d,]*\.\d{2})",
        r"Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total:\s+(\d[\d,]*\.\d{2})(?:\s+USD)?",
        r"New charges\s+\$(\d[\d,]*\.\d{2})",
        r"Invoice Total\s+\$(\d[\d,]*\.\d{2})",
        r"Billing Date\s+([A-z]+ \d{1,2}, \d{4})",
    ]

    # Calculate the row number where the file paths end
    last_row = ws.range('B' + str(ws.cells.last_cell.row)).end('up').row

    # Iterate over the PDF file paths listed in the Excel sheet
    for i in range(8, last_row + 1):
        total_cell_value = ws.range(f"C{i}").value
        pdf_path = ws.range(f"B{i}").value

        # Check if the "Total" cell value is "Total Not Found"
        if total_cell_value == "Total Not Found":
            # Perform OCR on the PDF
            images = pdf2image.convert_from_path(
                pdf_path,
                poppler_path=poppler_path
            )

            for image in images:
                ocr_text = pytesseract.image_to_string(image)
                for pattern in patterns:
                    match = re.search(pattern, ocr_text, re.IGNORECASE)
                    if match:
                        total = float(match.group(1).replace(',', ''))
                        ws.range(f"C{i}").value = total
                        break
                if match:
                    break

    # Save the workbook
    wb.save()


def extract_pdf_invoice_dates(wb, ws, start_date, end_date):
    date_col = 'E'  # Assuming column E is where you want to put the dates

    for i in range(8, ws.range('B' + str(ws.cells.last_cell.row)).end('up').row + 1):
        pdf_path = ws.range(f"B{i}").value
        date_found = False

        with pdfplumber.open(pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)
            dates = re.findall(r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}', text) + \
                    re.findall(r'\d{1,2}[\/-][A-Za-z]{3}[\/-]\d{2,4}', text) + \
                    re.findall(r'[A-Za-z]{3}\.?\s\d{1,2},\s\d{4}', text) + \
                    re.findall(r'[A-Za-z]+ \d{1,2}, \d{4}', text)
            for date_text in dates:
                parsed_date = dateparser.parse(date_text)
                if parsed_date and start_date <= parsed_date <= end_date:
                    formatted_date = parsed_date.strftime('%Y-%m-%d')
                    ws.range(f"{date_col}{i}").value = formatted_date
                    date_found = True
                    break

        if not date_found:
            ws.range(f"{date_col}{i}").value = "Date Not Found"

    wb.save()


def extract_ocr_invoice_dates(wb, ws, start_date, end_date):
    date_col = 'E'  # Assuming column E is for dates

    for i in range(8, ws.range('B' + str(ws.cells.last_cell.row)).end('up').row + 1):
        if ws.range(f"{date_col}{i}").value == "Date Not Found":
            pdf_path = ws.range(f"B{i}").value
            images = pdf2image.convert_from_path(
                pdf_path,
                poppler_path=poppler_path)
            date_found = False

            for image in images:
                text = pytesseract.image_to_string(image)
                # Simplify to search for any dates within the given range
                dates = re.findall(r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}', text) + \
                        re.findall(r'\d{1,2}[\/-][A-Za-z]{3}[\/-]\d{2,4}', text) + \
                        re.findall(r'[A-Za-z]{3}\.?\s\d{1,2},\s\d{4}', text) + \
                        re.findall(r'[A-Za-z]+ \d{1,2}, \d{4}', text)

                for date_text in dates:
                    parsed_date = dateparser.parse(date_text)
                    if parsed_date and start_date <= parsed_date <= end_date:
                        formatted_date = parsed_date.strftime('%Y-%m-%d')
                        ws.range(f"{date_col}{i}").value = formatted_date
                        date_found = True
                        break
                if date_found:
                    break

            if not date_found:
                ws.range(f"{date_col}{i}").value = "Date Not Found"

    wb.save()


def get_vendors(wb, ws, ws2):
    vendors = ws2.range('A2:A' + str(ws2.cells.last_cell.row)).value
    file_name_col = 'A'
    vendor_col = 'D'
    total_col = 'C'

    for row in range(8, ws.range(f'{file_name_col}' + str(ws.cells.last_cell.row)).end('up').row + 1):
        file_name = ws.range(f"{file_name_col}{row}").value.lower()
        total_value = ws.range(f"{total_col}{row}").value

        matched_vendor = None
        for vendor in vendors:
            if vendor is None:
                continue
            lower_vendor = vendor.lower()
            # handle "new" as a special case
            if 'newrelic' in file_name and lower_vendor == 'new':
                matched_vendor = 'NEW'
                break
            # handle "MSFT" as a special case
            elif 'msft' in file_name and lower_vendor == 'microsoft' and total_value == 144:
                matched_vendor = 'MSFT'
                break
            # handle "COMPTIA" as a special case
            elif 'comptia' in file_name:
                matched_vendor = 'COMPTIA'
                break
            # handle other vendors
            elif lower_vendor in file_name and lower_vendor not in ('new', 'microsoft'):
                matched_vendor = vendor
                break

        ws.range(f"{vendor_col}{row}").value = matched_vendor

    wb.save()


def run_notebook(notebook_path):
    # Load the notebook
    with open(notebook_path, 'r', encoding='utf-8') as f:
        nb = nbformat.read(f, as_version=4)

    # Set up a notebook processor and execute the notebook
    ep = ExecutePreprocessor(timeout=600, kernel_name='python3')
    ep.preprocess(nb, {'metadata': {'path': './'}})  # Adjust the path as necessary

    print("Notebook executed successfully.")


def main():
    # Load the Excel workbook and select the sheet
    workbook_path = 'G:/B_Amex/Template.xlsm'
    notebook_path = 'C:/Users/brand/IdeaProjects/Invoice_Reading/data_manipulation.ipynb'
    wb = xw.Book(workbook_path)
    ws = wb.sheets['File Name']
    ws1 = wb.sheets['Transaction Details 2']
    ws2 = wb.sheets['Xlookup table']

    # Define the date range
    start_date = datetime(2024, 1, 21)
    end_date = datetime(2024, 2, 21)

    # list_files_macro(wb)
    # extract_pdf_invoice_totals(wb, ws)
    # extract_ocr_invoice_totals(wb, ws)
    # get_vendors(wb, ws, ws2)
    # extract_pdf_invoice_dates(wb, ws, start_date, end_date)
    # extract_ocr_invoice_dates(wb, ws, start_date, end_date)
    # run_notebook(notebook_path)


if __name__ == "__main__":
    main()
