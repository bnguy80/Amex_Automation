import os
from tqdm import tqdm
import pdfplumber
import pdf2image
import re
import dateparser
from datetime import datetime
import xlwings as xw
import pytesseract
import pandas as pd

pytesseract.pytesseract.tesseract_cmd = "C:/Program Files/Tesseract-OCR/tesseract.exe"  # Explicitly set the ocr tesseract.exe path
poppler_path = "C:/Users/brand/OneDrive/Desktop/poppler-24.02.0/Library/bin"

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

    fall_back_total = float(666.66)

    fallback_date = datetime(1999, 1, 1)  # Using a datetime object for comparison

    def __init__(self, pdf_path):
        self.pdf_name = None
        self.pdf_path = pdf_path
        self.total = None
        self.date = None
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

        images = pdf2image.convert_from_path(self.pdf_path, poppler_path=poppler_path)
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
    invoice_counter = 0

    def __init__(self):
        self.pdfs_dataframe = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor', 'Date'])
        self.vendors = []

    def reset_counter(self):
        self.invoice_counter = 0

    def add_pdf(self, pdf_path, pdf_name, start_date, end_date):
        # Create a PDF object and extract data
        pdf = PDF(pdf_path)
        pdf.pdf_name = pdf_name

        # Increment the counter
        self.invoice_counter += 1

        # Directly invoke processing methods with appropriate dates
        pdf.process_totals()
        pdf.process_dates(start_date, end_date)
        pdf.process_vendor(self.vendors)

        # If it exists, update the existing entry
        if pdf.pdf_path in self.pdfs_dataframe['File Path'].values:
            # Selects all rows in the dataframe where the condition is "True", filtering rows that have matching "File Path"
            index = self.pdfs_dataframe[self.pdfs_dataframe['File Path'] == pdf.pdf_path].index[0]
            self.pdfs_dataframe.at[index, 'Amount'] = pdf.total
            self.pdfs_dataframe.at[index, 'Vendor'] = pdf.vendor
            self.pdfs_dataframe.at[index, 'Date'] = pdf.date.strftime('%Y-%m-%d') if isinstance(pdf.date,
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
            # Concat a new row to the dataframe
            print(f"Getting Invoice: {self.invoice_counter}")
            self.pdfs_dataframe = pd.concat([self.pdfs_dataframe, new_row_df], ignore_index=True)

    def populate_pdf_dataframe_from_worksheet(self, worksheet, start_date, end_date):
        data_df = worksheet.read_data_as_dataframe()
        # print(data_df.columns)  # Check if 'File Name' and 'File Path' are part of the columns

        for _, row in data_df.iterrows():
            self.add_pdf(row['File Path'], row['File Name'], start_date, end_date)

    def populate_vendors_from_vendor_worksheet(self, xlookup_table_worksheet):
        self.vendors.clear()  # Clear existing vendors to avoid duplication
        sheet = xlookup_table_worksheet.sheet
        vendors_range = sheet.range('A8:A' + str(sheet.cells.last_cell.row)).value

        if isinstance(vendors_range, list):
            # Iterate over the list and stop if a None value is encountered
            for vendor in vendors_range:
                # Check if the vendor is a tuple and has a non-None first element
                if isinstance(vendor, tuple) and vendor[0] is not None:
                    self.vendors.append(vendor[0])
                # Check if the vendor is not a tuple and is not None
                elif not isinstance(vendor, tuple) and vendor is not None:
                    self.vendors.append(vendor)
                # Break the loop if None is encountered
                else:
                    break
        elif vendors_range is not None:
            # If vendors_range is a single value (string or tuple), turn it into a list
            self.vendors = [vendors_range[0]] if isinstance(vendors_range, tuple) else [vendors_range]

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
        # Use xlwings to read data into a DataFrame, header True to interpret first row as column headers for the dataframe, index=False to make sure the first column is not interpreted as an index column
        dataframe = self.worksheet_dataframe = self.sheet.range('A7').options(pd.DataFrame, expand='table', header=True,
                                                                              index=False).value

        return dataframe

    def update_data_from_dataframe_to_sheet(self, pdf_data, progress_bar):
        # Assuming 'pdf_data' is a DataFrame with columns ['File Name', 'File Path', 'Amount', 'Vendor', 'Date']

        # Read the existing data from worksheet into a dataframe
        existing_data = self.read_data_as_dataframe()

        # Update the rows in existing_data_df with the data in pdf_data based on matching 'File Path'
        for _, pdf_row in pdf_data.iterrows():
            file_path = pdf_row['File Path']
            mask = existing_data['File Path'] == file_path

            # If the file_path exists in the existing_data_df, update the corresponding row
            if mask.any():
                existing_data.loc[mask, 'Amount'] = pdf_row['Amount']
                existing_data.loc[mask, 'Vendor'] = pdf_row['Vendor']
                existing_data.loc[mask, 'Date'] = pdf_row['Date']

            progress_bar.update(1)

        # Write the updated DataFrame back to the Excel sheet
        # This will overwrite the existing data starting from the top-left cell where your data begins (e.g., 'A7')
        self.sheet.range('A7').options(index=False).value = existing_data.reset_index(drop=True)


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
    xlookup_table_worksheet_name = "Xlookup table"  # Make sure this is correct
    path = "G:/B_Amex"
    amex_workbook_name = "Amex Corp Feb'24 - Addisu Turi (IT) (1).xlsx"
    template_workbook_name = "Template.xlsm"

    def __init__(self, start_date, end_date):
        self.workbooks_dict = {}
        self.pdf_collection = PDFCollection()
        self.start_date = start_date
        self.end_date = end_date
        self.manipulation = ExcelManipulation()

    #
    # def update_data_across_workbooks(self, source_workbook_name, target_workbook_name, criteria):

    def open_workbook(self, path, workbook_name):
        #  Opens an Excel workbook from the specified path and initializes all its worksheets. With workbook_name as the key in the workbook dictionary.
        workbook = Workbook(path)
        self.workbooks_dict[workbook_name] = workbook

    def get_workbook(self, workbook_name):
        return self.workbooks_dict[workbook_name]

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

            # This step is to get the Xlookup table worksheet to be able to get vendors for pdfs
        xlookup_table_worksheet = workbook.get_worksheet(self.xlookup_table_worksheet_name)
        self.pdf_collection.populate_vendors_from_vendor_worksheet(xlookup_table_worksheet)

        # This step populates the PDFCollection directly from the worksheet
        self.pdf_collection.populate_pdf_dataframe_from_worksheet(worksheet, self.start_date, self.end_date)

        pdf_data = self.pdf_collection.get_pdf_dataframe()
        num_updates = len(pdf_data.index)
        progress_bar = tqdm(total=num_updates, desc="Updating Excel Sheet Invoices from Invoice PDF Data",
                            bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt}')

        pdf_data = self.pdf_collection.get_pdf_dataframe()
        # print(pdf_data.columns)

        # Resets the invoice counter
        self.pdf_collection.reset_counter()

        # Updates the Invoice sheet from the gathered pdf data from the invoice folder
        worksheet.update_data_from_dataframe_to_sheet(pdf_data, progress_bar)

        progress_bar.close()

    def process_invoices_worksheet(self):
        # Macro name to get invoice pdf file names and file_paths from the invoices folder, need to adjust monthly 3/23/2024
        list_invoice_name_and_path = "ListFilesInSpecificFolder"

        # Open the template workbook and add all sheets
        template_workbook_path = os.path.join(self.path, self.template_workbook_name)
        self.open_workbook(template_workbook_path, self.template_workbook_name)

        # Get initial invoice names and invoice file paths for "Invoices" worksheet of Template workbook
        template_workbook = self.get_workbook(self.template_workbook_name)
        template_workbook.call_macro_workbook(list_invoice_name_and_path)

        # Update Template workbook "Invoices" worksheet from Invoice PDF data
        self.update_worksheet_from_pdf_collection(self.template_workbook_name, "Invoices")

        # Save the changes
        template_workbook.save()

    def process_transaction_details_worksheet(self):
        template_workbook = self.get_workbook(self.template_workbook_name)

        invoices_worksheet = template_workbook.get_worksheet("Invoices")
        transaction_details_worksheet = template_workbook.get_worksheet("Transactions Details 2")

        # Convert the worksheets into dataframes


controller = AutomationController("01/21/2024", "2/21/2024")
controller.process_invoices_worksheet()
