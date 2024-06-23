import re
from datetime import datetime
from typing import Union
import dateparser
import pandas as pd
import pdf2image
import pdfplumber
import pytesseract

# from invoice2data import extract_data
# from invoice2data.extract.loader import read_templates

# https://tesseract-ocr.github.io/tessdoc/Installation.html --> Tesseract at UB Mannheim --> https://github.com/UB-Mannheim/tesseract/wiki # How-to download tesseract for OCR on local machine 6/16/2024
# https://pypi.org/project/pytesseract/
# https://pypi.org/project/pdf2image/
# https://github.com/oschwartz10612/poppler-windows/releases


# C:/Users/bnguyen/AppData/Local/Programs/Tesseract-OCR/tesseract.exe -Truth
# C:/Program Files/Tesseract-OCR/tesseract.exe -computer

pytesseract.pytesseract.tesseract_cmd = "C:/Program Files/Tesseract-OCR/tesseract.exe"  # Explicitly set the ocr tesseract.exe path, need to also install it locally 6/15/2024

# C:/Users/bnguyen/AppData/Local/Programs/poppler-24.02.0/Library/bin -Truth
# C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin -computer

poppler_path = "C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin"  # Need to locally install it


# # THIS IS FOR INVOICE2DATA USAGE 6/15/2024
# # Set the path for pdftotext directly in the script
# # C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin -computer

# pdftotext_path = "C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin/"
# os.environ['PATH'] += os.pathsep + pdftotext_path


class PDF:
    # Static fallback patterns for pdfplumber and OCR; DO NOT CHANGE ORDER!
    _TOTAL_PATTERNS = [
        r"Grand Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total amount due(?: \(USD\))?:?\s+\$?\S?(\d[\d,]*\.\d{2})",
        r"Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total\s+\(in USD\)\s*:? ?\$?(\d[\d,]*\.\d{2})"
        r"Total:\s+(\d[\d,]*\.\d{2})(?:\s+USD)?",
        r"New charges\s+\$(\d[\d,]*\.\d{2})",
        r"Invoice Total\s+\$(\d[\d,]*\.\d{2})",
        r"Billing Date\s+([A-z]+ \d{1,2}, \d{4})",
    ]

    # Static fallback patterns for pdfplumber and OCR; DO NOT CHANGE ORDER!
    _DATE_PATTERNS = [
        r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}',
        r'\d{1,2}[\/-][A-Za-z]{3}[\/-]\d{2,4}',
        r'[A-Za-z]{3}\.?\s\d{1,2},\s\d{4}',
        r'[A-Za-z]+ \d{1,2}, \d{4}'
    ]

    # Template for vendor-specific patterns to extract total amounts and dates from identified invoice PDFs 6/16/2024.
    _VENDOR_PATTERNS = {
        # Comcast Business Internet
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
        # Comcast Cable
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

    def _identify_vendor(self, pdf_text: str) -> Union[dict, None]:
        for vendor_identifier, patterns in self._VENDOR_PATTERNS.items():
            if vendor_identifier in pdf_text:  # Using the key directly in the search
                self.vendor = vendor_identifier  # Optionally map to a more readable format if needed
                return patterns
        self.vendor = self.FALL_BACK_VENDOR
        return None  # Use to return prematurely because the vendor was identified 6/15/2024

    def _extract_pdf_invoice_total(self) -> Union[str, None]:
        """
        Extracts the total amount from a PDF invoice using vendor-specific patterns, with a fallback to general patterns.
        """
        with pdfplumber.open(self.pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)

        # Identify vendor and use specific patterns if available
        vendor_info = self._identify_vendor(text)
        if vendor_info and 'total' in vendor_info:
            total_patterns = vendor_info['total']
        else:
            total_patterns = self._TOTAL_PATTERNS  # Fallback to general patterns

        # Search for the total using the determined patterns
        for pattern in total_patterns:
            match = re.search(pattern, text, re.IGNORECASE)
            if match:
                self.total = float(match.group(1).replace(',', ''))
                return pattern  # Return the pattern used for matching

        # No match was found for total
        self.total = self.FALL_BACK_TOTAL
        return None

    def _extract_ocr_invoice_total(self) -> Union[str, None]:
        """
        This method extracts the total amount from an OCR-processed invoice. It uses the pytesseract library to convert images of the invoice to text and searches for patterns that indicate the presence of the total amount. If a match is found, the total amount is extracted and converted to a float.
        :param self: The instance of the class calling the method.
        :return: This method does not return any value. The extracted total amount is stored in the 'total' attribute of the calling instance.
        """
        images = pdf2image.convert_from_path(self.pdf_path,
                                             poppler_path=poppler_path)  # Converts the PDF file into an image 6/18/2024

        for image in images:
            ocr_text = pytesseract.image_to_string(image)
            for pattern in self._TOTAL_PATTERNS:
                match = re.search(pattern, ocr_text, re.IGNORECASE)
                if match:
                    # Extract the total and convert to float
                    self.total = float(match.group(1).replace(',', ''))
                    return pattern

        # When no match was found
        self.total = self.FALL_BACK_TOTAL  # set to 666.66 6/15/2024
        return None

    def _extract_pdf_invoice_date(self, start_date: str, end_date: str) -> Union[str, None]:
        """
        Extract invoice date from PDF file using vendor-specific patterns, with a fallback to general patterns.
        """
        start_date = dateparser.parse(start_date)
        end_date = dateparser.parse(end_date)

        with pdfplumber.open(self.pdf_path) as pdf:
            text = ' '.join(page.extract_text() or '' for page in pdf.pages)

        # Identify vendor and use specific patterns if available
        vendor_info = self._identify_vendor(text)
        if vendor_info and 'date' in vendor_info:
            date_patterns = vendor_info['date']
        else:
            date_patterns = self._DATE_PATTERNS  # Fallback to general patterns

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

    def _extract_ocr_invoice_date(self, start_date: str, end_date: str) -> Union[str, None]:
        start_date = dateparser.parse(start_date)
        end_date = dateparser.parse(end_date)

        images = pdf2image.convert_from_path(self.pdf_path, poppler_path=poppler_path)

        for image in images:
            ocr_text = pytesseract.image_to_string(image)
            for pattern in self._DATE_PATTERNS:
                dates = re.findall(pattern, ocr_text)
                for date_text in dates:
                    parsed_date = dateparser.parse(date_text)
                    if parsed_date and start_date <= parsed_date <= end_date:
                        formatted_date = parsed_date.strftime('%Y-%m-%d')
                        self.date = formatted_date
                        return pattern

        self.date = self.FALL_BACK_DATE
        return None

    def process_pdf_total(self) -> None:
        """
        Calls the date extraction methods in the correct sequence first with pdfplumber then if total not found, then use tesseract OCR
        """
        # Strategy 1: Use pdfplumber
        pattern_used_pdf = self._extract_pdf_invoice_total()  # Tries to parse regular pdf first
        if self.total == self.FALL_BACK_TOTAL:  # If the total is set to the fallback after using the pdfplumber function, then use OCR to find the total 6/15/2024
            # Strategy 2: Use pytesserract
            pattern_used_ocr = self._extract_ocr_invoice_total()  # Fall back to find the total

            # Print the pattern used, if any
            if pattern_used_ocr:
                print(f"Pattern Used to Find Amount (OCR): {pattern_used_ocr}")
            else:
                print("Amount not found in PDF or OCR.")
        elif pattern_used_pdf:
            print(f"Pattern Used to Find Amount (PDF): {pattern_used_pdf}")
        else:
            print("Amount successfully found, but no pattern was identified.")

    def process_pdf_date(self, start_date, end_date) -> None:
        # Strategy 1: Use pdfplumber
        pattern_used_pdf = self._extract_pdf_invoice_date(start_date, end_date)
        if self.date == self.FALL_BACK_DATE:
            # Strategy 2: Use pytesserract
            pattern_used_ocr = self._extract_ocr_invoice_date(start_date, end_date)  # Fall back to find the date

            # Print the pattern used, if any
            if pattern_used_ocr:
                print(f"Pattern Used to Find Date (OCR): {pattern_used_ocr}")
            else:
                print("Date not found in PDF or OCR.")
        elif pattern_used_pdf:
            print(f"Pattern Used to Find Date (PDF): {pattern_used_pdf}")
        else:
            print("Date successfully found, but no pattern was identified.")

    def match_pdf_invoice_vendor(self, vendors_list: list) -> None:
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
        # Note The header names (File Name (Invoices)--> File name (Transaction Details 2) is different compared to Invoice & Transaction Details 2 worksheets 6/15/2024
        self.pdf_collection_dataframe = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor', 'Date'])
        self.vendors_list = []

    def remove_pdf(self, pdf_name: str) -> None:
        # Find the index of rows where 'File Path' matches pdf_path
        rows_to_drop = self.pdf_collection_dataframe[self.pdf_collection_dataframe['File Name'] == pdf_name].index

        # Drop these rows from the DataFrame
        if not rows_to_drop.empty:
            self.pdf_collection_dataframe = self.pdf_collection_dataframe.drop(index=rows_to_drop)
        else:
            print(f"No PDF found with path: {pdf_name}")

    def remove_all_pdf(self) -> None:
        self.pdf_collection_dataframe = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor', 'Date'])

    def get_pdf_collection_dataframe(self):
        return self.pdf_collection_dataframe

    def reset_counter(self):
        self.invoice_counter = 0

    def create_pdf_instances_set_path_name_total_date_vendor(self, pdf_path: str, pdf_name: str, start_date: str,
                                                             end_date: str) -> None:

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

    def populate_pdf_collections_vendors_list_from_xlookup_worksheet(self, xlookup_table_worksheet) -> None:
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

    def populate_pdf_collection_dataframe_from_worksheet(self, worksheet, start_date: str, end_date: str):
        data_df = worksheet.read_data_as_dataframe()
        # print(data_df.columns)  # Check if 'File Name' and 'File Path' are part of the columns

        # Creating pdf instances; setting the path, name, total, date, vendor for each one. Then add it into the pdf_collection_dataframe 6/16/2024
        for _, row in data_df.iterrows():
            self.create_pdf_instances_set_path_name_total_date_vendor(row['File Path'], row['File Name'], start_date, end_date)
