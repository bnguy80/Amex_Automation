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
        self.data = {}

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

    # def extract_vendor(self):

    def extract_data(self):
        # Placeholder: Implement PDF data extraction logic here
        self.data = {"total": self.total, "date": self.date, "invoice_number": self.invoice_number,
                     "vendor": self.vendor}

    def get_data(self):
        return self.data


class PDFCollection:
    def __init__(self):
        self.pdfs = {}

    def add_pdf(self, pdf_path):
        pdf = PDF(pdf_path)
        pdf.extract_data()  # Assume PDFs are processed upon addition
        self.pdfs[pdf_path] = pdf

    def remove_pdf(self, pdf_path):
        del self.pdfs[pdf_path]

    def remove_all_pdfs(self):
        self.pdfs = {}

    # def extract_totals_from_all_pdfs(self):
    #
    # def extract_dates_from_all_pdfs(self):

    def aggregate_data_for_worksheet_update(self):
        aggregated_data = {}
        # Logic to aggregate data from PDFs
        for pdf_path, pdf in self.pdfs.items():
            data = pdf.get_data()
            # Example aggregation logic
            aggregated_data[pdf_path] = data
        return aggregated_data


class Worksheet:
    def __init__(self, name, sheet):
        self.name = name
        self.sheet = sheet
        self.dataframe = pd.DataFrame()

    def read_data_as_dataframe(self):
        # Use xlwings to read data into a DataFrame
        self.dataframe = self.sheet.range('A1').options(pd.DataFrame, expand='table').value

    def write_data_from_dataframe_to_sheet(self, dataframe):
        # Use xlwings to write DataFrame data back to the sheet
        self.sheet.range('A1').value = dataframe

    # def extract_column_data(self, column_letter):

    # def extract_row_data(self, row_letter):


class Workbook:
    def __init__(self, name, workbook_path=None):
        self.workbook_name = name
        self.worksheets = {}
        self.wb = xw.Book(workbook_path)

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
        return self.worksheets.get(worksheet_name)

    def call_macro_workbook(self, macro_name):
        macro_vba = self.wb.app.macro(macro_name)
        macro_vba()

    def save(self, save_path=None):
        if save_path:  # Save to a specific path when specified, or just save at the current location
            self.wb.save(save_path)
        else:
            self.wb.save()


class AutomationController:
    def __init__(self):
        self.workbooks = {}
        self.pdf_collection = PDFCollection()
        self.start_date = None
        self.end_date = None

    # def perform_task(self, workbook_name, worksheet_name):
    #
    # def update_data_across_workbooks(self, source_workbook_name, target_workbook_name, criteria):
    #
    # def gather_pdf_file_paths_and_names(self, workbook_name, worksheet_name):
    #
    # def collect_and_add_pdfs_to_collection(self, workbook_name, worksheet_name):

    def open_workbook(self, path, workbook_name):
        #  Opens an Excel workbook from the specified path and initializes all its worksheets. With workbook_name as the key in the workbook dictionary.
        workbook = Workbook(path)
        self.workbooks[workbook_name] = workbook

    # def save_workbook(self, workbook_name):

    def update_worksheet_from_pdf_collection(self, workbook_name, worksheet_name):
        workbook = self.workbooks[workbook_name]
        worksheet = workbook.get_worksheet(worksheet_name)
        pdf_data = self.pdf_collection.aggregate_data_for_worksheet_update()
        # Assume pdf_data is in a form that can be directly used to update a DataFrame
        worksheet.dataframe.update(pdf_data)
        worksheet.write_data_from_dataframe_to_sheet(worksheet.dataframe)

    # def extract_pdf_totals(self):
    #
    # def extract_pdf_dates(self):
