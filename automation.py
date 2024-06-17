import os
import re
from datetime import datetime
from itertools import combinations
import dateparser
import numpy as np
import pandas as pd
import pdf2image
import pdfplumber
import pytesseract
import xlwings as xw
from tabulate import tabulate
from tqdm import tqdm

# from invoice2data import extract_data
# from invoice2data.extract.loader import read_templates

# https://tesseract-ocr.github.io/tessdoc/Installation.html # How-to download tesseract for OCR on local machine 6/16/2024
# https://pypi.org/project/pytesseract/
# https://pypi.org/project/pdf2image/
# https://poppler.freedesktop.org/

# C:/Users/bnguyen/PycharmProjects/Tesseract-OCR/tesseract.exe -Truth
# C:/Program Files/Tesseract-OCR/tesseract.exe -computer
pytesseract.pytesseract.tesseract_cmd = "C:/Program Files/Tesseract-OCR/tesseract.exe"  # Explicitly set the ocr tesseract.exe path, need to also install it locally 6/15/2024

# C:/Users/bnguyen/PycharmProjects/poppler-24.02.0/Library/bin -Truth
# C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin -computer
poppler_path = "C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin"  # Need to locally install it


# # THIS IS FOR INVOICE2DATA USAGE 6/15/2024
# # Set the path for pdftotext directly in the script
# # C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin -computer
# pdftotext_path = "C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin/"
# os.environ['PATH'] += os.pathsep + pdftotext_path


class PDF:
    # Static fallback patterns for pdfplumber and OCR
    total_patterns = [
        r"Grand Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total amount due(?: \(USD\))?:?\s+\$?\S?(\d[\d,]*\.\d{2})",
        r"Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total\s+\(in USD\)\s*:? ?\$?(\d[\d,]*\.\d{2})"
        r"Total:\s+(\d[\d,]*\.\d{2})(?:\s+USD)?",
        r"New charges\s+\$(\d[\d,]*\.\d{2})",
        r"Invoice Total\s+\$(\d[\d,]*\.\d{2})",
        r"Billing Date\s+([A-z]+ \d{1,2}, \d{4})",
    ]

    # Static fallback patterns for pdfplumber and OCR
    date_patterns = [
        r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}',
        r'\d{1,2}[\/-][A-Za-z]{3}[\/-]\d{2,4}',
        r'[A-Za-z]{3}\.?\s\d{1,2},\s\d{4}',
        r'[A-Za-z]+ \d{1,2}, \d{4}'
    ]

    # Template for vendor-specific patterns to extract total amounts and dates from identified invoice PDFs 6/16/2024.
    vendor_patterns = {
        # Comcast Business Internet template 6/16/2024
        'Thanks for choosing Comcast Business!': {
            'date': [
                r'\d+\s+([A-Z][a-z]{2} \d{1,2}, \d{4})',
                # Add other date patterns for Comcast
            ],
            'total': [
                r'Regular monthly charges\s+\$([\d\.,]+)',
                # Add other total patterns for Comcast
            ]
        },
        # Comcast Cable template 6/16/2024
        'Comcast Business Cable': {
            'date': [
                r'(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})'
            ],
            'total': [
                r'Total Amount Due(?: \(USD\))?:?\s+\$?\S?(\d[\d,]*\.\d{2})'
            ]
        },
        'adobe': {
            'date': [
                r'\d{1,2}[\/-][A-Za-z]{3}[\/-]\d{2,4}'
            ],
            'total': [
                r'Grand Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        # Amazon invoices can only use OCR
        'amazon': {
            'date': [
                r'[A-Za-z]+ \d{1,2}, \d{4}'
            ],
            'total': [
                r'Grand Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        # Apple
        'Apple Store for Business': {
            'date': [
                r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        'calendy': {
            # Special case, cannot find it via OCR, but still try
            'date': [
                r'[A-Za-z]{3}\.?\s\d{1,2},\s\d{4}'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        'cbt': {
            'date': [
                r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}'
            ],
            'total': [
                r'Total\s+\(in USD\)\s*:? ?\$?(\d[\d,]*\.\d{2})'
            ]
        },
        'cloudflare': {
            'date': [
                r'Invoice Date:\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        'comptia': {
            'date': [
                r'Invoice Date:\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        'deft.com': {
            'date': [
                r'Date\s+(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        'dell!': {
            'date': [
                r'Purchased On:\s+([A-Za-z]{3}\.?\s\d{1,2},\s\d{4})'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        # Granite
        'www.granitenet.com': {
            'date': [
                r'INVOICE DATE:\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})'
            ],
            'total': [
                r'TOTAL AMOUNT DUE:\s*\$([\d,]+\.?\d*)'
            ]
        },
        'lastpass': {
            'date': [
                r'Invoice Date:\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        'Microsoft Corporation': {
            'date': [
                r'Due Date:\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})'
            ],
            'total': [
                r'Grand Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        'relic': {
            'date': [
                r'Due Date:\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})'
            ],
            'total': [
                r'Invoice Total\s+\$(\d[\d,]*\.\d{2})'
            ]
        },
        'www.serversupply.com': {
            'date': [
                r'Date:\s*(\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4})'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        # Special case, need OCR but still try
        'chatgpt': {
            'date': [
                r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
    }

    FALL_BACK_TOTAL = float(666.66)  # DO NOT CHANGE 6/16/2024
    FALL_BACK_DATE = datetime(1999, 1, 1)  # DO NOT CHANGE 6/16/2024
    FALL_BACK_VENDOR = 'Unknown'  # DO NOT CHANGE 6/16/2024

    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.pdf_name = None
        self.total = None
        self.date = None
        self.vendor = None

    def identify_vendor(self, text):
        for vendor_identifier, patterns in self.vendor_patterns.items():
            if vendor_identifier in text:  # Using the key directly in the search
                self.vendor = vendor_identifier  # Optionally map to a more readable format if needed
                return patterns
        self.vendor = self.FALL_BACK_VENDOR
        return None

    def extract_pdf_invoice_total(self):
        """
        Extracts the total amount from a PDF invoice using vendor-specific patterns, with a fallback to general patterns.
        """
        with pdfplumber.open(self.pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)

        # Identify vendor and use specific patterns if available
        vendor_info = self.identify_vendor(text)
        if vendor_info and 'total' in vendor_info:
            total_patterns = vendor_info['total']
        else:
            total_patterns = self.total_patterns  # Fallback to general patterns

        # Search for the total using the determined patterns
        for pattern in total_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                self.total = float(match.group(1).replace(',', ''))
                return pattern  # Return the pattern used for matching

        # No match was found for total
        self.total = self.FALL_BACK_TOTAL
        return None

    def extract_ocr_invoice_total(self):
        """
        This method extracts the total amount from an OCR-processed invoice. It uses the pytesseract library to convert images of the invoice to text and searches for patterns that indicate the presence of the total amount. If a match is found, the total amount is extracted and converted to a float.
        :param self: The instance of the class calling the method.
        :return: This method does not return any value. The extracted total amount is stored in the 'total' attribute of the calling instance.
        """
        images = pdf2image.convert_from_path(self.pdf_path, poppler_path=poppler_path)

        for image in images:
            ocr_text = pytesseract.image_to_string(image)
            for pattern in self.total_patterns:
                match = re.search(pattern, ocr_text, re.IGNORECASE)
                if match:
                    # Extract the total and convert to float
                    self.total = float(match.group(1).replace(',', ''))
                    return pattern

        # When no match was found
        self.total = self.FALL_BACK_TOTAL  # set to 666.66 6/15/2024
        return None

    # def extract_total_date_with_invoice2data(self, start_date, end_date):
    #     templates = read_templates('C:/Users/brand/IdeaProjects/Invoice_Reading/invoice2data_templates')
    #     data = extract_data(self.pdf_path, templates=templates)
    #
    #     if data:
    #         self.total = data['amount']
    #         self.date = data['date'].strftime('%Y-%m-%d')
    #     else:
    #         # If the fields could not be found using invoice2data 6/15/2024
    #         self.total = self.fall_back_total
    #         self.date = self.fallback_date

    # def process_pdf_total_date(self, start_date, end_date):
    #     # First attempt to extract using invoice2data
    #     self.extract_total_date_with_invoice2data(start_date, end_date)
    #     # Check if fallback values are used and use OCR if they are
    #     if self.total == self.fall_back_total:
    #         self.extract_ocr_invoice_total()
    #     if self.date == self.fallback_date:
    #         self.extract_ocr_invoice_date(start_date, end_date)

    def extract_pdf_invoice_date(self, start_date, end_date):
        """
        Extract invoice date from PDF file using vendor-specific patterns, with a fallback to general patterns.
        """
        start_date = dateparser.parse(start_date)
        end_date = dateparser.parse(end_date)

        with pdfplumber.open(self.pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)

        # Identify vendor and use specific patterns if available
        vendor_info = self.identify_vendor(text)
        if vendor_info and 'date' in vendor_info:
            date_patterns = vendor_info['date']
        else:
            date_patterns = self.date_patterns  # Fallback to general patterns

        # Search for the date using the determined patterns
        for pattern in date_patterns:
            dates = re.findall(pattern, text)
            for date_text in dates:
                parsed_date = dateparser.parse(date_text)
                if parsed_date and start_date <= parsed_date <= end_date:
                    formatted_date = parsed_date.strftime('%Y-%m-%d')
                    self.date = formatted_date
                    return pattern

        self.date = self.FALL_BACK_DATE
        return None

    def extract_ocr_invoice_date(self, start_date, end_date):

        start_date = dateparser.parse(start_date)
        end_date = dateparser.parse(end_date)

        images = pdf2image.convert_from_path(self.pdf_path, poppler_path=poppler_path)

        for image in images:
            ocr_text = pytesseract.image_to_string(image)
            for pattern in self.date_patterns:
                dates = re.findall(pattern, ocr_text)
                for date_text in dates:
                    parsed_date = dateparser.parse(date_text)
                    if parsed_date and start_date <= parsed_date <= end_date:
                        formatted_date = parsed_date.strftime('%Y-%m-%d')
                        self.date = formatted_date
                        return pattern

        self.date = self.FALL_BACK_DATE
        return None

    def process_pdf_total(self):
        """
        Calls the date extraction methods in the correct sequence first with pdfplumber then if total not found, then use tesseract OCR
        """
        pattern_used_pdf = self.extract_pdf_invoice_total()  # Tries to parse regular pdf first
        if self.total == self.FALL_BACK_TOTAL:  # If the total is set to the fallback after using the pdfplumber function, then use OCR to find the total 6/15/2024
            pattern_used_ocr = self.extract_ocr_invoice_total()  # Fall back to find the total

            # Print the pattern used, if any
            if pattern_used_ocr:
                print(f"Pattern Used to Find Amount (OCR): {pattern_used_ocr}")
            else:
                print("Amount not found in PDF or OCR.")
        elif pattern_used_pdf:
            print(f"Pattern Used to Find Amount (PDF): {pattern_used_pdf}")
        else:
            print("Amount successfully found, but no pattern was identified.")

    def process_pdf_date(self, start_date, end_date):
        # Calls the date extraction methods in the correct sequence
        pattern_used_pdf = self.extract_pdf_invoice_date(start_date, end_date)
        if self.date == self.FALL_BACK_DATE:
            pattern_used_ocr = self.extract_ocr_invoice_date(start_date, end_date)  # Fall back to find the date

            # Print the pattern used, if any
            if pattern_used_ocr:
                print(f"Pattern Used to Find Date (OCR): {pattern_used_ocr}")
            else:
                print("Date not found in PDF or OCR.")
        elif pattern_used_pdf:
            print(f"Pattern Used to Find Date (PDF): {pattern_used_pdf}")
        else:
            print("Date successfully found, but no pattern was identified.")

    def match_pdf_invoice_vendor(self, vendors_list):
        lower_file_name = self.pdf_name.lower()
        matched_vendor = None

        for vendor in vendors_list:
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
    invoice_counter = 0

    def __init__(self):
        self.pdf_collection_dataframe = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor',
                                                              'Date'])  # Note The header names (File Name--> File name) is different compared to Invoice & Transaction Details 2 worksheets 6/15/2024
        self.vendors_list = []

    def remove_pdf(self, pdf_name):
        # Find the index of rows where 'File Path' matches pdf_path
        rows_to_drop = self.pdf_collection_dataframe[self.pdf_collection_dataframe['File Name'] == pdf_name].index

        # Drop these rows from the DataFrame
        if not rows_to_drop.empty:
            self.pdf_collection_dataframe = self.pdf_collection_dataframe.drop(index=rows_to_drop)
        else:
            print(f"No PDF found with path: {pdf_name}")

    def remove_all_pdf(self):
        self.pdf_collection_dataframe = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor', 'Date'])

    def get_pdf_collection_dataframe(self):
        return self.pdf_collection_dataframe

    def reset_counter(self):
        self.invoice_counter = 0

    def create_pdf_instances_set_path_name_total_date_vendor(self, pdf_path, pdf_name, start_date, end_date):

        # Creates a PDF instance and sets the pdf invoice path and name first that is used later for further data extraction 6/15/2024
        pdf = PDF(pdf_path)
        pdf.pdf_name = pdf_name
        # Increment the counter
        self.invoice_counter += 1

        # Directly invoke processing methods to extract date, total, vendor, for each PDF object
        pdf.process_pdf_total()
        pdf.process_pdf_date(start_date, end_date)
        # pdf.process_pdf_total_date(start_date, end_date)
        pdf.match_pdf_invoice_vendor(self.vendors_list)

        # If it exists, update the existing entry
        if pdf.pdf_path in self.pdf_collection_dataframe['File Path'].values:
            # Selects all rows in the dataframe where the condition is "True", filtering rows that have matching "File Path"
            index = self.pdf_collection_dataframe[self.pdf_collection_dataframe['File Path'] == pdf.pdf_path].index[0]
            self.pdf_collection_dataframe.at[index, 'Amount'] = pdf.total
            self.pdf_collection_dataframe.at[index, 'Vendor'] = pdf.vendor
            self.pdf_collection_dataframe.at[index, 'Date'] = pdf.date.strftime('%Y-%m-%d') if isinstance(pdf.date,
                                                                                                          datetime) else pdf.date
        else:
            # If it doesn't exist, append a new dataframe row
            new_row_df = pd.DataFrame([{
                'File Name': pdf.pdf_name,
                'File Path': pdf.pdf_path,
                'Amount': pdf.total,
                'Vendor': pdf.vendor,
                'Date': pdf.date.strftime('%Y-%m-%d') if isinstance(pdf.date, datetime) else pdf.date
            }])

            # Concat a new row of pdf data to pdf_collection_dataframe
            self.pdf_collection_dataframe = pd.concat([self.pdf_collection_dataframe, new_row_df], ignore_index=True)

            # Print to concatenate the details of each pdf 6/15/2024
            print(f"Getting Invoice {self.invoice_counter}:\n"
                  f"File Name: {pdf.pdf_name}\n"
                  f"File Path: {pdf.pdf_path}\n"
                  f"Amount: {pdf.total}\n"
                  f"Vendor: {pdf.vendor}\n"
                  f"Date: {pdf.date.strftime('%Y-%m-%d') if isinstance(pdf.date, datetime) else pdf.date}")
            print("\n")

    def populate_pdf_collections_vendors_list_from_xlookup_worksheet(self, xlookup_table_worksheet):
        self.vendors_list.clear()  # Clear existing vendors to avoid duplication
        sheet = xlookup_table_worksheet.sheet
        vendors_range = sheet.range('A8:A' + str(sheet.cells.last_cell.row)).value

        if isinstance(vendors_range, list):
            # Iterate over the list and stop if a None value is encountered
            for vendor in vendors_range:
                # Check if the vendor is a tuple and has a non-None first element
                if isinstance(vendor, tuple) and vendor[0] is not None:
                    self.vendors_list.append(vendor[0])
                # Check if the vendor is not a tuple and is not None
                elif not isinstance(vendor, tuple) and vendor is not None:
                    self.vendors_list.append(vendor)
                # Break the loop if None is encountered
                else:
                    break
        elif vendors_range is not None:
            # If vendors_range is a single value (string or tuple), turn it into a list
            self.vendors_list = [vendors_range[0]] if isinstance(vendors_range, tuple) else [vendors_range]

    def populate_pdf_collection_dataframe_from_worksheet(self, worksheet, start_date, end_date):
        data_df = worksheet.read_data_as_dataframe()
        # print(data_df.columns)  # Check if 'File Name' and 'File Path' are part of the columns

        # Creating pdf instances; setting the path, name, total, date, vendor for each one. Then add it into the pdf_collection_dataframe 6/16/2024
        for _, row in data_df.iterrows():
            self.create_pdf_instances_set_path_name_total_date_vendor(row['File Path'], row['File Name'], start_date,
                                                                      end_date)


class Worksheet:
    def __init__(self, name, sheet):
        self.name = name
        self.sheet = sheet
        self.worksheet_dataframe = pd.DataFrame()

    def read_data_as_dataframe(self):
        # Use xlwings to read data into a DataFrame, header True to interpret first row as column headers for the dataframe, index=False to make sure the first column is not interpreted as an index column
        dataframe = self.worksheet_dataframe = self.sheet.range('A7').options(pd.DataFrame, expand='table', header=True,
                                                                              index=False).value

        return dataframe

    def update_data_from_dataframe_to_sheet(self, data_df, progress_bar):
        # Assuming 'data_df' is a DataFrame with columns ['File Name', 'File Path', 'Amount', 'Vendor', 'Date']

        # Read the existing data from worksheet into a dataframe
        existing_data_df = self.read_data_as_dataframe()

        # Update the rows in existing_data_df with the data in pdf_data based on matching 'File Path'
        for _, data_row in data_df.iterrows():
            file_path = data_row['File Path']
            mask = existing_data_df['File Path'] == file_path

            # If the file_path exists in the existing_data_df, update the corresponding row
            if mask.any():
                existing_data_df.loc[mask, 'Amount'] = data_row['Amount']
                existing_data_df.loc[mask, 'Vendor'] = data_row['Vendor']
                existing_data_df.loc[mask, 'Date'] = data_row['Date']

            progress_bar.update(1)

        # Write the updated DataFrame back to the Excel sheet
        # This will overwrite the existing data starting from the top-left cell where your data begins (e.g., 'A7')
        self.sheet.range('A7').options(index=False).value = existing_data_df.reset_index(drop=True)


class Workbook:
    def __init__(self, workbook_path=None):
        self.worksheets = {}

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

    def remove_all_worksheets_dict(self):
        self.worksheets = {}

    def get_worksheet(self, worksheet_name):
        return self.worksheets[worksheet_name]

    def get_all_worksheets(self):
        return self.worksheets

    def save(self, save_path=None):
        if save_path:  # Save to a specific path when specified, or just save at the current location
            self.wb.save(save_path)
        else:
            self.wb.save()

    def call_macro_workbook(self, macro_name, macro_parameter_1: None, macro_parameter_2: None):
        macro_vba = self.wb.app.macro(macro_name)
        macro_vba(macro_parameter_1, macro_parameter_2)


class DataManipulation:
    @staticmethod
    def find_matching_transactions(invoice_df, transaction_df):

        # Each transaction is matched with a unique row, "File Name" linked to a unique transaction, each row in invoice_df represents a distinct invoice PDF and only matches with one row in transaction_df
        matched_transactions = set()  # This set will track matched transactions
        matched_invoices = set()  # This set will track matched invoice indices.
        # Iterate over the invoice dataframe
        for index, invoice_row in invoice_df.iterrows():
            DataManipulation.match_transaction(invoice_row, transaction_df, matched_transactions, matched_invoices,
                                               index)

        # After all, invoices have been processed in finding matches, filter and print unmatched matches
        unmatched_invoices = invoice_df.loc[~invoice_df.index.isin(matched_invoices)]
        if not unmatched_invoices.empty:
            print("\n")
            print("Unmatched Invoices:")
            print(tabulate(unmatched_invoices, headers='keys', tablefmt='psql'))

    @staticmethod
    def find_combinations(transactions, target_amount):
        # Only consider combinations of up to 3 transactions to reduce complexity
        for r in range(1, min(4, len(transactions) + 1)):
            for combo in combinations(transactions, r):
                if np.isclose(sum(item['Amount'] for item in combo), target_amount, atol=0.01):
                    return combo
        return None

    @staticmethod
    def match_transaction(invoice_row, transaction_details_df, matched_transactions, matched_invoices, index):

        # Extract relevant info from the invoice_row
        invoice_row_vendor = invoice_row['Vendor']  # Extract the vendor from the invoice row
        invoice_row_total = invoice_row[
            'Amount']  # Extract the Amount from the invoice row to compare against the transaction total
        invoice_row_date = pd.to_datetime(invoice_row['Date'])  # Ensure datetime format
        invoice_row_file_name = invoice_row['File Name']  # Extrac the file name from the invoice row
        invoice_row_file_path = invoice_row['File Path']  # Extract the file path from the invoice row

        # Ensure "File path" column exists in the DataFrame
        if 'File path' not in transaction_details_df.columns:
            transaction_details_df['File path'] = None

        # Basic filtering by vendor and not matched yet 6/15/2024
        potential_matches_candidates = transaction_details_df[
            (transaction_details_df['Vendor'].str.contains(invoice_row_vendor, case=False, na=False)) &
            (~transaction_details_df.index.isin(matched_transactions))
            ]

        # Strategy 1: Exact match on Amount and Date
        exact_matches = potential_matches_candidates[
            (potential_matches_candidates['Amount'] == invoice_row_total) &
            (pd.to_datetime(potential_matches_candidates['Date'], errors='coerce') == invoice_row_date)
            ]

        if not exact_matches.empty:
            first_match_index = exact_matches.iloc[0].name
            transaction_details_df.at[first_match_index, 'File name'] = invoice_row_file_name
            transaction_details_df.at[first_match_index, 'Column1'] = "Amount & Date Match"
            transaction_details_df.at[first_match_index, 'File path'] = invoice_row_file_path
            matched_transactions.add(first_match_index)
            matched_invoices.add(index)
            print(f"Match found for Strategy 1 in transaction with id {first_match_index}!")
            return  # Match found and processed, return early

        # Strategy 2: Match by Vendor and Amount, ignoring date
        if invoice_row_date is not pd.NaT:  # If date is a valid datetime
            non_date_matches = potential_matches_candidates[
                (potential_matches_candidates['Amount'] == invoice_row_total) &
                (pd.to_datetime(potential_matches_candidates['Date'], errors='coerce') != invoice_row_date)
                ]

            # Include detail in "Column1" that user needs to manually check the match to make sure it is indeed the correct invoice even though the date does not match
            if not non_date_matches.empty:
                first_match_index = non_date_matches.iloc[0].name
                transaction_details_df.at[first_match_index, 'File name'] = invoice_row_file_name
                transaction_details_df.at[first_match_index, 'Column1'] = 'Amount & Non-Date Match'
                transaction_details_df.at[first_match_index, 'File path'] = invoice_row_file_path
                matched_transactions.add(first_match_index)
                matched_invoices.add(index)
                print(f"Match found for Strategy 2 in transaction with id {first_match_index}!")
                return  # Match found and processed, return early

        # Strategy 3: Match by vendor, date, and among a subset of transactions that when added together equal the amount from the invoice_row["Amount"]

        # Filter potential_matches by same "Description", "Date", "Vendor" from transaction_details_df
        # Ensure a datetime format for filtering and filter out already matched transactions
        potential_combination_candidates_strategy_3 = transaction_details_df[
            (transaction_details_df['Vendor'] == invoice_row_vendor) &
            (pd.to_datetime(transaction_details_df['Date'], errors='coerce') == invoice_row_date) &
            (~transaction_details_df.index.isin(matched_transactions))
            ].reset_index(drop=False)  # Keep the original index to use it later 6/15/2024

        # Group by 'Description', 'Vendor', and 'Date', then look for matching combinations within each group
        for (_, group) in potential_combination_candidates_strategy_3.groupby(['Description', 'Vendor', 'Date']):
            transactions_to_check = group.to_dict('records')  # Now includes 'index' from DataFrame
            print("Potential Candidates Strategy 3")
            print(tabulate(transactions_to_check, headers='keys', tablefmt='psql'))
            combo = DataManipulation.find_combinations(transactions_to_check, invoice_row_total)
            if combo:
                for transaction in combo:
                    # Use the 'index' key to identify the original row in transaction_details_df
                    idx = transaction['index']  # 'index' is now explicitly included
                    transaction_details_df.at[idx, 'File name'] = invoice_row_file_name
                    transaction_details_df.at[idx, 'Column1'] = 'Combination Amount Match'
                    transaction_details_df.at[idx, 'File path'] = invoice_row_file_path
                    matched_transactions.add(idx)
                    matched_invoices.add(index)
                print(f"Match found for Strategy 3 in transactions with id {[t['index'] for t in combo]}!")
                return  # Match found and processed, return early

        # No match found, set the File name to "None" and consider additional strategies or manual review 6/15/2024
        print(f"No match found for {invoice_row_file_name} with Amount {invoice_row_total}. Consider manual review.")


class AutomationController:
    XLOOKUP_TABLE_WORKSHEET_NAME = "Xlookup table"  # Make sure this is correct, 3/24/24: is correct inside Template.xlsm 6/15/2024
    TEMPLATE_WORKBOOK_NAME = "Template.xlsm"  # This is the workbook that we will be storing the intermediary data for matching AMEX Statement transactions and invoices for 6/15/2024.
    TEMPLATE_INVOICES_WORKSHEET_NAME = "Invoices"
    TEMPLATE_TRANSACTION_DETAILS_2_WORKSHEET_NAME = "Transaction Details 2"
    LIST_INVOICE_NAME_AND_PATH_MACRO_NAME = "ListFilesInSpecificFolder"  # Macro name to get invoice pdf file names and file_paths from the invoices folder 6/15/2024

    def __init__(self, amex_path, amex_statement, start_date, end_date, macro_parameter_1=None, macro_parameter_2=None):
        self.workbooks_dict = {}
        self.pdf_collection = PDFCollection()
        self.manipulation = DataManipulation()
        self.amex_path = amex_path  # The directory where the AMEX Statement workbook is located 6/16/2024
        self.amex_statement = amex_statement  # This is the final workbook that the automation will put the data into; sent to Ana 6/15/2024.
        self.start_date = start_date  # Start date of Amex Statement transactions
        self.end_date = end_date  # End date of Amex Statement transactions
        self.macro_parameter_1 = macro_parameter_1
        self.macro_parameter_2 = macro_parameter_2

    def open_workbook(self, path, workbook_name):
        #  Opens an Excel workbook from the specified path and initializes all its worksheets. With workbook_name as the key in the workbook dictionary.
        workbook = Workbook(path)
        self.workbooks_dict[workbook_name] = workbook

    def get_workbook(self, workbook_name):
        return self.workbooks_dict[workbook_name]

    def save_selected_workbook(self, workbook_name, save_path=None):
        if workbook_name in self.workbooks_dict:
            workbook = self.workbooks_dict[workbook_name]
            workbook.save(save_path)
        else:
            print(f"Workbook '{workbook_name}' not found.")

    #
    # def update_data_across_workbooks(self, source_workbook_name, target_workbook_name, criteria):

    @staticmethod
    def duplicate_and_label_rows(df):
        # Find indices of 'CLOUDFLARE' rows to duplicate
        cloudflare_indices = df[df['Vendor'] == 'CLOUDFLARE'].index.tolist()
        offset = 0  # Offset to adjust indices after each insertion of a copy

        for index in cloudflare_indices:
            # Duplicate the row
            duplicated_row = df.loc[index].copy()
            # Insert the duplicated row immediately after the original
            df = pd.concat([df.iloc[:index + 1 + offset], pd.DataFrame([duplicated_row]),
                            df.iloc[index + 1 + offset:]]).reset_index(drop=True)
            offset += 1  # Increment offset for each insertion of a copy

        # Renumber file names starting from index 8
        for i in range(len(df)):
            df.at[i, 'File name'] = f"{8 + i} - {df.loc[i, 'File name']}"
            # Need to also add logic to navigate to file path and rename file if a File path is present 6/16/2024.
        return df

    def update_invoices_worksheet_with_all_extracted_data(self):
        template_workbook = self.workbooks_dict.get(self.TEMPLATE_WORKBOOK_NAME)
        if template_workbook is None:
            print(f"Workbook '{self.TEMPLATE_WORKBOOK_NAME}' not found")
            return  # Returns early because the workbook was not found
        invoice_worksheet = template_workbook.get_worksheet(self.TEMPLATE_INVOICES_WORKSHEET_NAME)
        if invoice_worksheet is None:
            print(
                f"Worksheet '{self.TEMPLATE_INVOICES_WORKSHEET_NAME}' not found in Workbook '{self.TEMPLATE_WORKBOOK_NAME}'.")
            return  # Returns early because the worksheet was not found

        # This step is to get the Xlookup table worksheet to be able to get vendors for pdfs
        xlookup_table_worksheet = template_workbook.get_worksheet(self.XLOOKUP_TABLE_WORKSHEET_NAME)
        self.pdf_collection.populate_pdf_collections_vendors_list_from_xlookup_worksheet(xlookup_table_worksheet)

        # This step populates the pdf_collection_dataframe with all the pdf data 6/16/2024
        self.pdf_collection.populate_pdf_collection_dataframe_from_worksheet(invoice_worksheet, self.start_date, self.end_date)
        pdf_collection_df = self.pdf_collection.get_pdf_collection_dataframe()
        num_updates = len(pdf_collection_df.index)
        progress_bar = tqdm(total=num_updates, desc="Updating Invoices Worksheet from Extracted PDF Data", bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt}')

        # print(pdf_data.columns)

        # Resets the invoice counter
        self.pdf_collection.reset_counter()

        invoice_worksheet.update_data_from_dataframe_to_sheet(pdf_collection_df, progress_bar)  # Updates the Invoice worksheet with all the necessary fields for each pdf invoice 6/15/2024

        progress_bar.close()

    def process_invoices_worksheet(self):
        # Open the template workbook and add all sheets
        template_workbook_path = os.path.join(self.amex_path, self.TEMPLATE_WORKBOOK_NAME)
        self.open_workbook(template_workbook_path, self.TEMPLATE_WORKBOOK_NAME)

        # Get initial invoice names and invoice file paths for "Invoices" worksheet of Template workbook
        template_workbook = self.get_workbook(self.TEMPLATE_WORKBOOK_NAME)
        template_workbook.call_macro_workbook(self.LIST_INVOICE_NAME_AND_PATH_MACRO_NAME, self.macro_parameter_1, self.macro_parameter_2)

        # Update Template workbook "Invoices" worksheet from Invoice PDF Collection DataFrame
        self.update_invoices_worksheet_with_all_extracted_data()

        # Save the changes
        template_workbook.save()

    def process_transaction_details_worksheet(self):
        # Open the template workbook and add all sheets
        template_workbook_path = os.path.join(self.amex_path, self.TEMPLATE_WORKBOOK_NAME)
        self.open_workbook(template_workbook_path, self.TEMPLATE_WORKBOOK_NAME)

        template_workbook = self.get_workbook(self.TEMPLATE_WORKBOOK_NAME)
        invoices_worksheet = template_workbook.get_worksheet(self.TEMPLATE_INVOICES_WORKSHEET_NAME)
        transaction_details_worksheet = template_workbook.get_worksheet(self.TEMPLATE_TRANSACTION_DETAILS_2_WORKSHEET_NAME)

        # Convert the Invoice worksheet into DataFrame
        invoices_worksheet_df = invoices_worksheet.read_data_as_dataframe()
        print("Invoices DataFrame before matching")
        print(tabulate(invoices_worksheet_df, headers='keys', tablefmt='psql'))
        print("\n")

        # Convert Transaction Details 2 worksheet into DataFrame
        transaction_details_worksheet_df = transaction_details_worksheet.read_data_as_dataframe()
        # Print the transaction details DataFrame before matching
        print("Before Matching Transaction Details 2 DataFrame:")
        print(tabulate(transaction_details_worksheet_df, headers='keys', tablefmt='psql'))
        print("\n")

        # Matches invoice files found in "Invoices" worksheet to "Transaction Details 2" worksheet transactions
        # Works through 3 strategies of 1: exact matching between vendor|date|amount 2: match between vendor|amount|non-matching date or 3: target total between subset of transactions that sum to amount of invoice
        DataManipulation.find_matching_transactions(invoices_worksheet_df, transaction_details_worksheet_df)

        # Call function to duplicate Cloudflare rows and update file names starting from index 8
        transaction_details_worksheet_df = self.duplicate_and_label_rows(transaction_details_worksheet_df)
        print("\n")
        print("After Matching Invoice and Transactions Transaction Details 2 DataFrame")
        print(tabulate(transaction_details_worksheet_df, headers='keys', tablefmt='psql'))
        print("\n")


controller = AutomationController("K:/B_Amex", "Amex Corp Feb'24 - Addisu Turi (IT).xlsx", "01/21/2024", "2/21/2024",
                                  r"K:\t3nas\APPS\\",
                                  "[02] Feb 2024")  # Make sure to have "r" and \ at the end to treat as raw string parameter 6/15/2024
controller.process_invoices_worksheet()
controller.process_transaction_details_worksheet()
