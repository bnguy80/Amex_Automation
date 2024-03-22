import pdfplumber
import pdf2image
import re
import dateparser
from datetime import datetime
import xlwings as xw
import pytesseract
import pandas as pd


class PDF:
    # Patterns are static, define them at the class level
    total_patterns = [
        r"Grand Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total amount due(?: \(USD\))?:?\s+\$?\S?(\d[\d,]*\.\d{2})",
        r"Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total:\s+(\d[\d,]*\.\d{2})(?:\s+USD)?",
        r"New charges\s+\$(\d[\d,]*\.\d{2})",
        r"Invoice Total\s+\$(\d[\d,]*\.\d{2})",
        r"Billing Date\s+([A-z]+ \d{1,2}, \d{4})",
    ]

    date_patterns = [
        r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}',
        r'\d{1,2}[\/-][A-Za-z]{3}[\/-]\d{2,4}',
        r'[A-Za-z]{3}\.?\s\d{1,2},\s\d{4}',
        r'[A-Za-z]+ \d{1,2}, \d{4}'
    ]

    poppler_path = "C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin"

    fall_back_total = float(666.66)

    fallback_date = datetime(1999, 1, 1)  # Using a datetime object for comparison

    def __init__(self, pdf_path):
        self.pdf_name = None
        self.pdf_path = pdf_path
        self.total = None
        self.date = None
        self.invoice_number = None
        self.vendor = None

    def extract_pdf_invoice_total(self):
        match_found = False  # Flag to indicate a match was found

        with pdfplumber.open(self.pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)
            for pattern in self.total_patterns:
                match = re.search(pattern, text, re.IGNORECASE)
                if match:
                    self.total = float(match.group(1).replace(',', ''))
                    match_found = True
                    break
        if not match_found:
            self.total = self.fall_back_total

    def extract_ocr_invoice_total(self):

        images = pdf2image.convert_from_path(self.pdf_path, poppler_path=self.poppler_path)
        match_found = False  # Flag to indicate match was found

        for image in images:
            ocr_text = pytesseract.image_to_string(image)
            for pattern in self.total_patterns:
                match = re.search(pattern, ocr_text, re.IGNORECASE)
                if match:
                    # Extract the total and convert to float
                    self.total = float(match.group(1).replace(',', ''))
                    match_found = True
                    break  # Exit the inner loop after finding the first match
            if match_found:
                break  # Exit the outer loop after finding the first match

        if not match_found:
            # When no match was found
            self.total = self.fall_back_total

    def process_totals(self):
        # Calls the date extraction methods in the correct sequence
        self.extract_pdf_invoice_total()  # Tries to parse regular pdf first
        if self.total == self.fall_back_total:
            self.extract_ocr_invoice_total()  # Fall back to find the total

    def extract_pdf_invoice_date(self, start_date, end_date):

        start_date = dateparser.parse(start_date)
        end_date = dateparser.parse(end_date)
        match_found = False

        with pdfplumber.open(self.pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)
            for pattern in self.date_patterns:
                dates = re.findall(pattern, text)
                for date_text in dates:
                    parsed_date = dateparser.parse(date_text)
                    if parsed_date and start_date <= parsed_date <= end_date:
                        formatted_date = parsed_date.strftime('%Y-%m-%d')
                        self.date = formatted_date
                        match_found = True
                        break
                if match_found:
                    break
        if not match_found:
            self.date = self.fallback_date

    def extract_ocr_invoice_date(self, start_date, end_date):

        start_date = dateparser.parse(start_date)
        end_date = dateparser.parse(end_date)
        match_found = False

        images = pdf2image.convert_from_path(self.pdf_path, poppler_path=self.poppler_path)

        for image in images:
            ocr_text = pytesseract.image_to_string(image)
            for pattern in self.date_patterns:
                dates = re.findall(pattern, ocr_text)
                for date_text in dates:
                    parsed_date = dateparser.parse(date_text)
                    if parsed_date and start_date <= parsed_date <= end_date:
                        formatted_date = parsed_date.strftime('%Y-%m-%d')
                        self.date = formatted_date
                        match_found = True
                        break
                if match_found:
                    break
            if match_found:
                break
        if not match_found:
            self.date = self.fallback_date

    def process_dates(self, start_date, end_date):
        # Calls the date extraction methods in the correct sequence
        self.extract_pdf_invoice_date(start_date, end_date)
        if self.date == self.fallback_date:
            self.extract_ocr_invoice_date(start_date, end_date)  # Fall back to find the date

    def process_vendor(self, vendors):
        lower_file_name = self.pdf_name.lower()
        matched_vendor = None

        for vendor in vendors:
            if vendor is None:
                continue

            lower_vendor = vendor.lower()
            if 'newrelic' in lower_file_name and lower_vendor == 'new':
                matched_vendor = 'NEW'
            elif 'msft' in lower_file_name and lower_vendor == 'microsoft':
                matched_vendor = 'MSFT'
            elif lower_vendor in lower_file_name and lower_vendor not in ('new', 'microsoft'):
                matched_vendor = vendor

        self.vendor = matched_vendor if matched_vendor else 'Unknown'


class PDFCollection:
    def __init__(self):
        self.pdfs_dataframe = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor', 'Date'])
        self.vendors = []

    def add_pdf(self, pdf_path, pdf_name, start_date, end_date):
        # Create a PDF object and extract data
        pdf = PDF(pdf_path)
        pdf.pdf_name = pdf_name

        # Directly invoke processing methods with appropriate dates
        pdf.process_totals()
        pdf.process_dates(start_date, end_date)
        pdf.process_vendor(self.vendors)

        # Add a new row to the DataFrame
        new_row = {
            'File Name': pdf.pdf_name,
            'File Path': pdf.pdf_path,
            'Amount': pdf.total,
            'Vendor': pdf.vendor,
            'Date': pdf.date.strftime('%Y-%m-%d') if isinstance(pdf.date, datetime) else pdf.date
        }

        # Append the new row to the DataFrame
        self.pdfs_dataframe = self.pdfs_dataframe.append(new_row, ignore_index=True)

    def populate_pdf_dataframe_from_worksheet(self, worksheet, start_date, end_date):
        data_df = worksheet.read_data_as_dataframe()

        for _, row in data_df.iterrows():
            self.add_pdf(row['File Name'], row['File'], start_date, end_date)

    def populate_vendors_from_vendor_worksheet(self, vendor_worksheet):
        self.vendors.clear()  # Clear existing vendors to avoid duplication
        vendors_range = vendor_worksheet.range('A7:A' + str(vendor_worksheet.cells.last_cell.row)).value
        if isinstance(vendors_range, list):
            # Flatten the list if it's a list of tuples (happens with single column ranges)
            self.vendors = [vendor[0] for vendor in vendors_range if vendor[0] is not None]
        else:
            # In case it's a single value (only one vendor in the list)
            self.vendors = [vendors_range] if vendors_range else []

    def remove_pdf(self, pdf_name):
        # Find the index of rows where 'File Path' matches pdf_path
        rows_to_drop = self.pdfs_dataframe[self.pdfs_dataframe['File Name'] == pdf_name].index

        # Drop these rows from the DataFrame
        if not rows_to_drop.empty:
            self.pdfs_dataframe = self.pdfs_dataframe.drop(index=rows_to_drop)
        else:
            print(f"No PDF found with path: {pdf_name}")

    def remove_all_pdf(self):
        self.pdfs_dataframe = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor', 'Date'])

    def get_pdf_dataframe(self):
        return self.pdfs_dataframe


class Worksheet:
    def __init__(self, name, sheet):
        self.name = name
        self.sheet = sheet
        self.worksheet_dataframe = pd.DataFrame()

    def read_data_as_dataframe(self):
        # Use xlwings to read data into a DataFrame, header True to interpret first row as column headers for the dataframe and index True to include the correct index from the sheet to dataframe
        self.worksheet_dataframe = self.sheet.range('A7').options(pd.DataFrame, expand='table', header=True,
                                                                  index=True).value

    def update_data_from_dataframe_to_sheet(self, dataframe):
        # Use xlwings to write DataFrame data back to the sheet
        pass


class Workbook:
    def __init__(self, workbook_path=None):

        if workbook_path is None:
            self.wb = xw.Book()
            self.workbook_name = self.wb.name
        else:
            self.wb = xw.Book(workbook_path)
            self.workbook_name = self.wb.name

        # Automatically add all existing worksheets
        for sheet in self.wb.sheets:
            self.worksheets[sheet.name] = Worksheet(sheet.name, sheet)

    def add_worksheet(self, worksheet_name):
        if worksheet_name not in self.worksheets:
            sheet = self.wb.sheets.add(worksheet_name)
            self.worksheets[worksheet_name] = Worksheet(worksheet_name, sheet)
        else:
            print(f"Worksheet '{worksheet_name}' already exists.")

    def remove_worksheet(self, worksheet_name):
        if worksheet_name in self.worksheets:
            self.wb.sheets[worksheet_name].delete()
            del self.worksheets[worksheet_name]
        else:
            print(f"Worksheet '{worksheet_name}' not found.")

    def remove_all_worksheet(self):
        self.worksheets = {}

    def get_worksheet(self, worksheet_name):
        return self.worksheets[worksheet_name]

    def get_all_worksheets(self):
        return self.worksheets

    def call_macro_workbook(self, macro_name):
        macro_vba = self.wb.app.macro(macro_name)
        macro_vba()

    def save(self, save_path=None):
        if save_path:  # Save to a specific path when specified, or just save at the current location
            self.wb.save(save_path)
        else:
            self.wb.save()


class ExcelManipulation:
    pass


class AutomationController:
    vendor_worksheet_name = 'Xlookup table'  # Make sure this is correct

    def __init__(self, start_date, end_date):
        self.workbooks_dict = {}
        self.pdf_collection = PDFCollection()
        self.start_date = start_date
        self.end_date = end_date
        self.manipulation = ExcelManipulation()

    # def perform_task(self, workbook_name, worksheet_name):
    #
    # def update_data_across_workbooks(self, source_workbook_name, target_workbook_name, criteria):

    def open_workbook(self, path, workbook_name):
        #  Opens an Excel workbook from the specified path and initializes all its worksheets. With workbook_name as the key in the workbook dictionary.
        workbook = Workbook(path)
        self.workbooks_dict[workbook_name] = workbook

    def get_workbooks(self):
        return self.workbooks_dict

    def save_workbook(self, workbook_name, save_path=None):
        if workbook_name in self.workbooks_dict:
            workbook = self.workbooks_dict[workbook_name]
            workbook.save(save_path)
        else:
            print(f"Workbook '{workbook_name}' not found.")

    def update_worksheet_from_pdf_collection(self, workbook_name, worksheet_name):
        workbook = self.workbooks_dict.get(workbook_name)
        if workbook is None:
            print(f"Workbook '{workbook_name}' not found")
            return
        worksheet = workbook.get_worksheet(worksheet_name)
        if worksheet is None:
            print(f"Worksheet '{worksheet_name}' not found in Workbook '{workbook_name}'.")
            return

        # This step is to get the vendor worksheet to be able to get vendors for pdfs
        vendor_worksheet = workbook.get_worksheet(self.vendor_worksheet_name)
        self.pdf_collection.populate_vendors_from_vendor_worksheet(vendor_worksheet)

        # This step populates the PDFCollection directly from the worksheet
        self.pdf_collection.populate_pdf_dataframe_from_worksheet(worksheet, self.start_date, self.end_date)

        pdf_data = self.pdf_collection.get_pdf_dataframe()

        worksheet.update_data_from_dataframe_to_sheet(pdf_data)
