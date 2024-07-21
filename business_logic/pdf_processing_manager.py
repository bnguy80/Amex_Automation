import abc
from datetime import datetime
import pandas as pd

from models.pdf import PDF
from business_logic.pdf_processor import PDFProcessor


class PDFProcessingManager:
    pdf_counter = 0

    def __init__(self, text_processor: PDFProcessor, ocr_processor: PDFProcessor):
        # This mirrors the headers present in the Invoices worksheet of Template – Master.xlsm 7/2/2024
        self.pdf_proc_mng_df: pd.DataFrame = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor', 'Date'])
        self.text_processor: PDFProcessor = text_processor
        self.ocr_processor: PDFProcessor = ocr_processor

    def remove_pdf_proc_mng_df_row(self, pdf_name: str) -> None:
        # Find the index of rows where 'File Path' matches pdf_path
        rows_to_drop = self.pdf_proc_mng_df[self.pdf_proc_mng_df['File Name'] == pdf_name].index

        # Drop these rows from the DataFrame
        if not rows_to_drop.empty:
            self.pdf_proc_mng_df = self.pdf_proc_mng_df.drop(index=rows_to_drop)
        else:
            print(f"No PDF found with path: {pdf_name}")

    def clear_pdf_proc_mng_df(self) -> None:
        self.pdf_proc_mng_df = pd.DataFrame(columns=['File Name', 'File Path', 'Amount', 'Vendor', 'Date'])

    def get_pdf_proc_mng_df(self) -> pd.DataFrame:
        return self.pdf_proc_mng_df

    def _reset_counter(self) -> None:
        self.pdf_counter = 0

    def _add_pdf(self, pdf: PDF) -> None:
        new_row = {
            'File Name': pdf.pdf_name,
            'File Path': pdf.pdf_path,
            'Amount': pdf.total,
            'Vendor': pdf.vendor,
            'Date': pdf.date
        }
        self.pdf_proc_mng_df = pd.concat([self.pdf_proc_mng_df, pd.DataFrame([new_row])], ignore_index=True)

    def _log_pdf_processing_details(self, pdf, pattern_used_pdf_amount, pattern_used_ocr_amount, pattern_used_pdf_date, pattern_used_ocr_date) -> None:

        log_msg = f"Processing PDF {self.pdf_counter}:\n"
        log_msg += f"File Name: {pdf.pdf_name}\n"
        log_msg += f"File Path: {pdf.pdf_path}\n"
        log_msg += f"Amount: {pdf.total}\n"
        log_msg += f"Vendor: {pdf.vendor}\n"
        log_msg += f"Date: {pdf.date.strftime('%Y-%m-%d') if isinstance(pdf.date, datetime) else pdf.date}\n"

        if pattern_used_pdf_amount:
            log_msg += f"Pattern Used to Find Amount (PDF): {pattern_used_pdf_amount}\n"
        if pattern_used_ocr_amount:
            log_msg += f"Pattern Used to Find Amount (OCR): {pattern_used_ocr_amount}\n"
        if pattern_used_pdf_date:
            log_msg += f"Pattern Used to Find Date (PDF): {pattern_used_pdf_date}\n"
        if pattern_used_ocr_date:
            log_msg += f"Pattern Used to Find Date (OCR): {pattern_used_ocr_date}\n"

        print(log_msg)

    def _select_processor_and_extract_total(self, pdf: PDF):
        # try:/except: block, use pdfplumber first then except to use ocr 6/28/2024
        pattern_used_pdf = self.text_processor.extract_total(pdf)
        pattern_used_ocr = None
        if pdf.total == 666.66:
            pattern_used_ocr = self.ocr_processor.extract_total(pdf)
        return pattern_used_pdf, pattern_used_ocr

    def _select_processor_and_extract_date(self, pdf: PDF):
        # try:\except: block, use pdfplumber first then except to use ocr 6/28/2024
        pattern_used_pdf = self.text_processor.extract_date(pdf)
        pattern_used_ocr = None
        if pdf.date == datetime(1999, 1, 1).strftime('%Y-%m-%d'):
            pattern_used_ocr = self.ocr_processor.extract_date(pdf)
        return pattern_used_pdf, pattern_used_ocr

    def _process_pdf(self, pdf_path: str, pdf_name: str) -> None:

        # Creates a PDF instance and sets the pdf invoice path and name first that is used later for further data extraction 6/15/2024
        pdf: PDF = PDF(pdf_path, pdf_name)
        # Increment the counter
        self.pdf_counter += 1

        # Directly invoke processing methods to extract date, total, vendor, for each PDF object
        # try:\except: block to log pdf that wasn't successful in extracting data possibly? 6/28/2024
        pattern_used_pdf_total, pattern_used_ocr_total = self._select_processor_and_extract_total(pdf)
        pattern_used_pdf_date, pattern_used_ocr_date = self._select_processor_and_extract_date(pdf)
        self.text_processor.extract_vendor(pdf)

        self._add_pdf(pdf)
        self._log_pdf_processing_details(pdf, pattern_used_pdf_total, pattern_used_ocr_total, pattern_used_pdf_date, pattern_used_ocr_date)

    def populate_pdf_proc_mng_df(self, invoice_worksheet, xlookup_table_worksheet) -> None:
        invoice_df: pd.DataFrame = invoice_worksheet.read_data_as_dataframe()

        # Populate PDFProcessor vendors_list to be able to match for pdf.vendor during data extraction
        self.text_processor.get_vendors_from_xlookup_worksheet(xlookup_table_worksheet)

        # Creating pdf instances; setting the path, name, total, date, vendor for each one. Then add it into the pdf_collection_dataframe 6/16/2024
        for _, row in invoice_df.iterrows():
            self._process_pdf(row['File Path'], row['File Name'])

        self._reset_counter()
