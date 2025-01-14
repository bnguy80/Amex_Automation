import datetime
from abc import abstractmethod, ABC
from typing import Union, List, Protocol

import pdfplumber
import pdf2image
import re
import pytesseract
import dateparser

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

class GeneralPatternProvider(Protocol):

    def get_total_pattern(self) -> List[str]:
        """Returns a list of generic total patterns"""

    def get_date_pattern(self) -> List[str]:
        """Returns a list of generic date patterns"""


class VendorSpecificPatternProvider(Protocol):

    def get_total_pattern(self, pdf_text: str) -> List[str]:
        """Returns the vendor-specific total pattern"""

    def get_date_pattern(self, pdf_text: str) -> List[str]:
        """Returns the vendor-specific date pattern"""


class GeneralPattern:
    # Static fallback patterns for pdfplumber and OCR; DON'T CHANGE ORDER!
    _TOTAL_PATTERNS = [
        r"Grand Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total amount due(?: \(USD\))?:?\s+\$?\S?(\d[\d,]*\.\d{2})",
        r"Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})",
        r"Total\s+\(in USD\)\s*:? ?\$?(\d[\d,]*\.\d{2})"
        r"Total:\s+(\d[\d,]*\.\d{2})(?:\s+USD)?",
        r"New charges\s+\$(\d[\d,]*\.\d{2})",
        r"Invoice Total\s+\$(\d[\d,]*\.\d{2})",
        r"Billing Date\s+([A-z]+ \d{1,2}, \d{4})",
        r"Order total\s+\$(\d[\d,]*\.\d{2})"
    ]

    # Static fallback patterns for pdfplumber and OCR; DON'T CHANGE ORDER!
    _DATE_PATTERNS = [
        r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}',
        r'\d{1,2}[\/-][A-Za-z]{3}[\/-]\d{2,4}',
        r'[A-Za-z]{3}\.?\s\d{1,2},\s\d{4}',
        r'[A-Za-z]+ \d{1,2}, \d{4}'
    ]

    def get_total_pattern(self) -> List[str]:
        return self._TOTAL_PATTERNS

    def get_date_pattern(self) -> List[str]:
        return self._DATE_PATTERNS


class VendorSpecificPattern:
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
            # Special case can't find it via OCR, but still try
            'date': [
                r'[A-Za-z]{3}\.?\s\d{1,2},\s\d{4}'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        # CBT
        'sales@cbtnuggets.com': {
            'date': [
                r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}'
            ],
            'total': [
                r'Total \(in USD\)\s*:?\s*\$?(\d{1,3}(?:,\d{3})*\.\d{2})'
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
        # Special case, need OCR but still try as it is not a PDF but an image 6/24/2024.
        'chatgpt': {
            'date': [
                r'\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}'
            ],
            'total': [
                r'Total(?: \(USD\))?:?\s+\$?(\d[\d,]*\.\d{2})'
            ]
        },
        # CDW 6/24/2024
        'cdw.com': {
            'date': [
                r'Due Date\s*.*\$\d+\.\d+\s+(\d{1,2}/\d{1,2}/\d{4})'
            ],
            'total': [
                r'Amount Due(?: \(USD\))?:?\s+\$?\S?(\d[\d,]*\.\d{2})'
            ]
        },
        # Special case where it doesn't have the correct character mapping, so cid:15...etc. is extracted from pdfplumber instead 6/24/2024
        'EBAY': {
            'date': [
                r'Placed On:\s+([A-Za-z]{3}\.?\s\d{1,2},\s\d{4})'
            ],
            'total': [
                r'Order total\s+\$(\d[\d,]*\.\d{2})'
            ]
        },
        # OTTERAI 6/24/2024
        'otter.ai': {
            'date': [
                r'Date due ([A-Za-z]{3}\.?\s\d{1,2},\s\d{4})',
                r'Date issued ([A-Za-z]{3}\.?\s\d{1,2},\s\d{4})'
            ],
            'total': [
                r'Amount due\s+\$(\d[\d,]*\.\d{2})',
                r'Total refunded without credit note\s+\$(\d[\d,]*\.\d{2})'
            ]
        },
        # SYMPREX 6/24/2024
        'symprex.com': {
            'date': [
                r'Invoice date:\s*(\d{2}-[A-Za-z]{3}-\d{4})'
            ],
            'total': [
                r'Total:\s+([0-9,]+\.\d{2})\s+USD'
            ]
        },
        # GODADDY, need to use OCR,
        # will still populate date and total pattern that would have found a match using pdfplumber 7/15/2024
        'GoDaddy.com': {
            'date': [

            ],
            'total': [

            ]
        },
        'Monday.com': {
            'date': [

            ],
            'total': [

            ]
        }
    }

    def get_total_pattern(self, pdf_text: str) -> List[str]:
        for vendor_identifier, patterns in self._VENDOR_PATTERNS.items():
            if vendor_identifier in pdf_text:
                return patterns['total']
        return []

    def get_date_pattern(self, pdf_text: str) -> List[str]:
        for vendor_identifier, patterns in self._VENDOR_PATTERNS.items():
            if vendor_identifier in pdf_text:
                return patterns['date']
        return []
    

class PDFProcessor(ABC):

    _FALL_BACK_TOTAL = float(666.66)  # DON'T CHANGE 6/16/2024
    _FALL_BACK_DATE = datetime.date(1999, 1, 1)  # DON'T CHANGE 6/16/2024
    _FALL_BACK_VENDOR = 'Unknown'  # DON'T CHANGE 6/16/2024

    def __init__(self, start_date, end_date):
        self._start_date = start_date
        self._end_date = end_date
        self._vendors_list = []

    @abstractmethod
    def extract_total(self, pdf):
        ...

    @abstractmethod
    def extract_date(self, pdf):
        ...

    def get_vendors_from_xlookup_worksheet(self, xlookup_table_worksheet) -> None:

        self._vendors_list.clear()  # Clear existing vendors to avoid duplication
        sheet = xlookup_table_worksheet.sheet
        vendors_range = sheet.range('A8:A' + str(sheet.cells.last_cell.row)).value

        if isinstance(vendors_range, list):
            # Iterate over the list and stop if a None value is encountered
            for vendor in vendors_range:
                # Check if the vendor is a tuple and has a non-None first element
                if isinstance(vendor, tuple) and vendor[0] is not None:
                    self._vendors_list.append(vendor[0])
                # Check if the vendor is not a tuple and is not None
                elif not isinstance(vendor, tuple) and vendor is not None:
                    self._vendors_list.append(vendor)
                # Break the loop if None is encountered
                else:
                    break
        elif vendors_range is not None:
            # If vendors_range is a single value (string or tuple), turn it into a list
            self._vendors_list = [vendors_range[0]] if isinstance(vendors_range, tuple) else [vendors_range]

    def extract_vendor(self, pdf):

        lower_file_name = pdf.pdf_name.lower()
        matched_vendor = None

        for vendor in self._vendors_list:
            if vendor is None:
                continue

            lower_vendor = vendor.lower()
            if 'newrelic' in lower_file_name and lower_vendor == 'new':
                matched_vendor = 'NEW'
            elif 'msft' in lower_file_name and lower_vendor == 'microsoft':
                matched_vendor = 'MSFT'
            elif lower_vendor in lower_file_name and lower_vendor not in ('new', 'microsoft'):
                matched_vendor = vendor

        pdf.vendor = matched_vendor if matched_vendor else self._FALL_BACK_VENDOR


class PDFPlumberProcessor(PDFProcessor):

    def __init__(self, start_date, end_date, vendor_specific_pattern: VendorSpecificPatternProvider, general_pattern: GeneralPatternProvider):
        super().__init__(start_date, end_date)
        self._vendor_specific_pattern = vendor_specific_pattern
        self._general_pattern = general_pattern

    def extract_total(self, pdf):

        try:
            with pdfplumber.open(pdf.pdf_path) as pdf_text:
                text = ' '.join(page.extract_text() or '' for page in pdf_text.pages)

            total_patterns = self._vendor_specific_pattern.get_total_pattern(text)
            if len(total_patterns) == 0:
                total_patterns = self._general_pattern.get_total_pattern()

            # Search for the total using the determined patterns
            for pattern in total_patterns:
                match = re.search(pattern, text, re.IGNORECASE)  # Ignore case sensitivity 6/24/2024
                if match:
                    extracted_value = match.group(1).replace(',', '')
                    pdf.total = extracted_value
                    return pattern  # Return the pattern used for matching
            # No match was found for the total
            pdf.total = self._FALL_BACK_TOTAL
            return None
        except FileNotFoundError as ex:
            raise FileNotFoundError(f"File not found while extracting PDF data: {pdf.pdf_path}") from ex

    def extract_date(self, pdf):

        try:
            start_date = dateparser.parse(self._start_date)
            end_date = dateparser.parse(self._end_date)

            with pdfplumber.open(pdf.pdf_path) as pdf_text:
                text = ' '.join(page.extract_text() or '' for page in pdf_text.pages)

            date_patterns = self._vendor_specific_pattern.get_date_pattern(text)
            if len(date_patterns) == 0:
                date_patterns = self._general_pattern.get_date_pattern()

            for pattern in date_patterns:
                dates = re.findall(pattern, text)
                for date_text in dates:
                    parsed_date = dateparser.parse(date_text)
                    if parsed_date and start_date <= parsed_date <= end_date:
                        pdf.date = parsed_date
                        return pattern
            pdf.date = self._FALL_BACK_DATE
            return None
        except FileNotFoundError as ex:
            raise FileNotFoundError(f"File not found while extracting PDF data: {pdf.pdf_path}") from ex


class PDFOCRProcessor(PDFProcessor):

    def __init__(self, start_date, end_date, general_pattern: GeneralPatternProvider):
        super().__init__(start_date, end_date)
        self._general_pattern = general_pattern

    def extract_total(self, pdf):

        try:
            total_patterns = self._general_pattern.get_total_pattern()
            images = pdf2image.convert_from_path(pdf.pdf_path, poppler_path=poppler_path)

            for image in images:
                ocr_text = pytesseract.image_to_string(image)
                for pattern in total_patterns:
                    match = re.search(pattern, ocr_text, re.IGNORECASE)
                    if match:
                        extracted_value = match.group(1).replace(',', '')
                        pdf.total = extracted_value
                        return pattern
            pdf.total = self._FALL_BACK_TOTAL
            return None
        except FileNotFoundError as ex:
            raise FileNotFoundError(f"File not found while extracting PDF data: {pdf.pdf_path}") from ex

    def extract_date(self, pdf):

        try:
            date_patterns = self._general_pattern.get_date_pattern()
            start_date = dateparser.parse(self._start_date)
            end_date = dateparser.parse(self._end_date)
            images = pdf2image.convert_from_path(pdf.pdf_path, poppler_path=poppler_path)

            for image in images:
                ocr_text = pytesseract.image_to_string(image)
                for pattern in date_patterns:
                    dates = re.findall(pattern, ocr_text)
                    for date_text in dates:
                        parsed_date = dateparser.parse(date_text)
                        if parsed_date and start_date <= parsed_date <= end_date:
                            pdf.date = parsed_date
                            return pattern
            pdf.date = self._FALL_BACK_DATE
            return None
        except FileNotFoundError as ex:
            raise FileNotFoundError(f"File not found while extracting PDF data: {pdf.pdf_path}") from ex
