from datetime import datetime
import pandas as pd

from models.pdf import PDF
from pdf_processor import PDFPlumberProcessor, PDFOCRProcessor


class PDFCollectionManager:
    invoice_counter = 0

    def __init__(self, start_date, end_date):
        # Note The header names (File Name (Invoices)--> File name (Transaction Details 2) is different compared to Invoice & Transaction Details 2 worksheets 6/15/2024
        self.pdf_collection_dataframe = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor', 'Date'])
        self.start_date: str = start_date
        self.end_date: str = end_date
        self.text_processor = PDFPlumberProcessor(start_date, end_date)
        self.ocr_processor = PDFOCRProcessor(start_date, end_date)

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

    @staticmethod
    def _display_processor_used(pattern, data_type, processor_type):
        if pattern:
            print(f'Pattern Used to Find {data_type} ({processor_type}): {pattern}')
        else:
            print(f'{data_type} could not be found')

    def _select_processor_and_extract_total(self, pdf):
        # try:/except: block, use pdfplumber first then except to use ocr 6/28/2024
        pattern_used_pdf = self.text_processor.extract_total(pdf)
        if pdf.total == 666.66:
            pattern_used_ocr = self.ocr_processor.extract_total(pdf)
            self._display_processor_used(pattern_used_ocr, 'Amount', 'OCR')
        elif pattern_used_pdf:
            self._display_processor_used(pattern_used_pdf, 'Amount', 'PDF')

        #     # Print the pattern used, if any
        #     if pattern_used_ocr:
        #         print(f"Pattern Used to Find Amount (OCR): {pattern_used_ocr}")
        #     else:
        #         print("Amount not found in PDF or OCR.")
        # elif pattern_used_pdf:
        #     print(f"Pattern Used to Find Amount (PDF): {pattern_used_pdf}")

    def _select_processor_and_extract_date(self, pdf):
        # try:\except: block, use pdfplumber first then except to use ocr 6/28/2024
        pattern_used_pdf = self.text_processor.extract_date(pdf)
        if pdf.date == datetime(1999, 1, 1).strftime('%Y-%m-%d'):
            pattern_used_ocr = self.ocr_processor.extract_date(pdf)
            self._display_processor_used(pattern_used_ocr, 'Date', 'OCR')
        elif pattern_used_pdf:
            self._display_processor_used(pattern_used_pdf, 'Date', 'PDF')

        #     # Print the pattern used, if any
        #     if pattern_used_ocr:
        #         print(f"Pattern Used to Find Date (OCR): {pattern_used_ocr}")
        #     else:
        #         print("Date not found in PDF or OCR.")
        # elif pattern_used_pdf:
        #     print(f"Pattern Used to Find Date (PDF): {pattern_used_pdf}")

    def _extract_data_add_invoice_to_collection(self, pdf_path: str, pdf_name: str) -> None:

        # Creates a PDF instance and sets the pdf invoice path and name first that is used later for further data extraction 6/15/2024
        pdf = PDF(pdf_path)
        pdf.pdf_name = pdf_name
        # Increment the counter
        self.invoice_counter += 1

        # Directly invoke processing methods to extract date, total, vendor, for each PDF object
        # try:\except: block to log pdf that wasn't successful in extracting data possibly? 6/28/2024
        self._select_processor_and_extract_total(pdf)
        self._select_processor_and_extract_date(pdf)
        self.text_processor.extract_vendor(pdf)

        # If it exists, update the existing entry
        if pdf.pdf_path in self.pdf_collection_dataframe['File Path'].values:
            # Selects all rows in the dataframe where the condition is "True", filtering rows that have matching "File Path"
            index = self.pdf_collection_dataframe[self.pdf_collection_dataframe['File Path'] == pdf.pdf_path].index[0]
            self.pdf_collection_dataframe.at[index, 'Amount'] = pdf.total
            self.pdf_collection_dataframe.at[index, 'Vendor'] = pdf.vendor
            self.pdf_collection_dataframe.at[index, 'Date'] = pdf.date.strftime('%Y-%m-%d') if isinstance(pdf.date, datetime) else pdf.date
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

    def extract_pdf_data_populate_pdm_df(self, invoice_worksheet, xlookup_table_worksheet):
        invoice_df = invoice_worksheet.read_data_as_dataframe()

        # Populate PDFProcessor vendors_list to be able to match for pdf.vendor during data extraction
        self.text_processor.get_vendors_from_xlookup_worksheet(xlookup_table_worksheet)

        # Creating pdf instances; setting the path, name, total, date, vendor for each one. Then add it into the pdf_collection_dataframe 6/16/2024
        for _, row in invoice_df.iterrows():
            self._extract_data_add_invoice_to_collection(row['File Path'], row['File Name'])
