import os
import time

from business_logic.workbook_manager import TemplateWorkbookManager, AmexWorkbookManager
from business_logic.pdf_processing_manager import PDFProcessingManager
from business_logic.invoice_matching_manager import invoice_matching_manager
from utils.utilities import print_dataframe


class AmexAutomationOrchestrator:
    XLOOKUP_TABLE_WORKSHEET_NAME: str = "Xlookup table"  # Make sure this is correct, 3/24/24: is correct inside Template - Master.xlsm 6/15/2024
    TEMPLATE_WORKBOOK_NAME: str = "Template - Master.xlsm"  # This is the workbook that we will be storing the intermediary data for matching AMEX Statement transactions and invoices for 6/15/2024.
    TEMPLATE_INVOICES_WORKSHEET_NAME: str = "Invoices"
    TEMPLATE_TRANSACTION_DETAILS_2_WORKSHEET_NAME: str = "Transaction Details 2"
    AMEX_TRANSACTION_DETAILS_WORKSHEET_NAME: str = "Transaction Details"
    LIST_INVOICE_NAME_AND_PATH_MACRO_NAME: str = "ListFilesInSpecificFolder"  # Macro name to get invoice pdf file Names and file_paths from the invoices folder 6/15/2024
    RESIZE_TABLE_MACRO_NAME: str = "ResizeTable"

    def __init__(self, amex_path, amex_statement_name, start_date, end_date, macro_parameter_1=None, macro_parameter_2=None):

        self.amex_path = amex_path  # The directory where the AMEX Statement workbook is located 6/16/2024
        self.amex_statement = amex_statement_name  # This is the final workbook that the automation will put the data into; sent to Ana 6/15/2024.
        self.template_workbook_path = os.path.join(self.amex_path, self.TEMPLATE_WORKBOOK_NAME)
        self.amex_workbook_path = str(os.path.join(self.amex_path, self.amex_statement))
        self.macro_parameter_1 = macro_parameter_1
        self.macro_parameter_2 = macro_parameter_2

        # Start date of Amex Statement transactions
        # End date of Amex Statement transactions
        self.pdf_proc_mng = PDFProcessingManager(start_date, end_date)
        self.invoice_matching_manager = invoice_matching_manager  # Using a list of strategies to match invoices to transactions. ONLY ONE INSTANCE 6/22/2024.
        self.template_workbook_manager = TemplateWorkbookManager(self.TEMPLATE_WORKBOOK_NAME, self.template_workbook_path)
        self.amex_workbook_manager = AmexWorkbookManager(self.amex_statement, self.amex_workbook_path)

    def prepare_template_workbook(self):

        amex_statement = self.amex_workbook_manager.get_worksheet(self.AMEX_TRANSACTION_DETAILS_WORKSHEET_NAME)
        amex_statement_df = amex_statement.read_data_as_dataframe()
        # Close the Workbook after getting the data so that there is no confusing with running macros in Template - Master.xlsm 7/7/2024
        self.amex_workbook_manager.workbook.close()

        # Also keep in mind that the CLOUDFLARE transaction will only show the total before the split between MARKETING (.5098039216) and COMMS (.4901960784) 7/7/2024
        # Before updating the worksheet need to clear the contents of the Date, Description, Amount columns from the table first --> VBA macro? 7/7/2024

        transaction_details_worksheet = self.template_workbook_manager.get_worksheet("Transaction Details 2")
        transaction_details_worksheet.update_sheet(amex_statement_df)

        # After updating the worksheet, resize the table 7/7/2024
        self.template_workbook_manager.workbook.call_macro_workbook(self.RESIZE_TABLE_MACRO_NAME)


    def process_invoices_worksheet(self):

        # Get initial invoice names and invoice file paths for the "Invoices" worksheet of Template workbook
        self.template_workbook_manager.workbook.call_macro_workbook(self.LIST_INVOICE_NAME_AND_PATH_MACRO_NAME, self.macro_parameter_1, self.macro_parameter_2)
        invoice_worksheet = self.template_workbook_manager.get_worksheet(self.TEMPLATE_INVOICES_WORKSHEET_NAME)

        # This step is to get the Xlookup table worksheet to be able to get vendors for pdfs
        xlookup_table_worksheet = self.template_workbook_manager.get_worksheet(self.XLOOKUP_TABLE_WORKSHEET_NAME)

        # This step populates the pdf_processing_manager with all the pdf data of the path, name, total, date, vendor 7/2/2024
        self.pdf_proc_mng.populate_pdf_proc_mng_df(invoice_worksheet, xlookup_table_worksheet)
        pdf_proc_mng_df = self.pdf_proc_mng.get_pdf_proc_mng_df()

        # Updates the Invoice worksheet from pdf_proc_mng_df with all required data to begin matching between transaction statements in transaction_details_df 7/2/2024
        invoice_worksheet.update_sheet(pdf_proc_mng_df)

        # Save the changes
        self.template_workbook_manager.workbook.save()

    def process_transaction_details_2_worksheet(self) -> None:

        # Convert the Invoice worksheet into DataFrame
        invoices_worksheet = self.template_workbook_manager.get_worksheet(self.TEMPLATE_INVOICES_WORKSHEET_NAME)
        invoices_worksheet_df = invoices_worksheet.read_data_as_dataframe()
        print_dataframe(invoices_worksheet_df, "Invoices DataFrame Before Matching Process:")

        # Convert Transaction Details 2 worksheet into DataFrame
        transaction_details_worksheet = self.template_workbook_manager.get_worksheet(self.TEMPLATE_TRANSACTION_DETAILS_2_WORKSHEET_NAME)
        transaction_details_worksheet_df = transaction_details_worksheet.read_data_as_dataframe()
        # Print the transaction details DataFrame before matching
        print_dataframe(transaction_details_worksheet_df, "Transaction Details 2 DataFrame Before Matching Process:")

        # Update with the 'File path' column to the end if it isn't already present.
        if 'File Path' not in transaction_details_worksheet_df.columns:
            transaction_details_worksheet_df['File Path'] = ''

        # Sets the preprocessed dataframes to the InvoiceMatchingManager class to do further processing 6/19/2024.
        self.invoice_matching_manager.set_data(invoices_worksheet_df, transaction_details_worksheet_df)

        # Matches invoice files found in "Invoices" worksheet to "Transaction Details 2" worksheet transactions
        # Works through primary strategies of 1: exact matching between vendor|date|amount 2: match between vendor|amount|non-matching date or 3: target total between subset of transactions that sum to amount of invoice
        # Fallback strategy of matching only by vendor after going through primary strategies
        self.invoice_matching_manager.execute_invoice_matching()

        # Sequence the 'File Name' column of invoices that matches were found for transaction_details_df, starting at index 8 6/29/2024
        self.invoice_matching_manager.sequence_file_names()

        # Drop the 'File path' column from the transaction_details_worksheet_df 6/23/2024
        if 'File Path' in transaction_details_worksheet_df.columns:
            transaction_details_worksheet_df = transaction_details_worksheet_df.drop('File Path', axis=1)

        print_dataframe(transaction_details_worksheet_df, "Transaction Details 2 DataFrame After Matching Sequencing File Names:")

        transaction_details_worksheet.update_sheet(transaction_details_worksheet_df)

    def process_amex_transaction_details_worksheet(self) -> None:
        transaction_details_worksheet = self.template_workbook_manager.get_worksheet(self.TEMPLATE_TRANSACTION_DETAILS_2_WORKSHEET_NAME)
        transaction_details_worksheet_df = transaction_details_worksheet.read_data_as_dataframe()

        amex_worksheet = self.amex_workbook_manager.get_worksheet(self.AMEX_TRANSACTION_DETAILS_WORKSHEET_NAME)
        amex_worksheet.update_sheet(transaction_details_worksheet)


# "H:/Amex Automation" Automation Truth--> amex_path
# "C:/Users/brand/IdeaProjects/Amex Automation DATA" -computer
# r"H:\Amex Automation\t3nas\APPS\\" -Truth--> macro_parameter_1
# r"C:\Users\brand\IdeaProjects\Amex Automation DATA\t3nas\APPS\\" -computer
path_truth = "H:/Amex Automation"
path_computer = "C:/Users/brand/IdeaProjects/Amex Automation DATA"
macro_truth = r"H:\Amex Automation\t3nas\APPS\\"
macro_computer = r"C:\Users\brand\IdeaProjects\Amex Automation DATA\t3nas\APPS\\"

# Make sure to have "r" and \ at the end to treat as raw string parameter 6/15/2024
controller = AmexAutomationOrchestrator(path_computer, "Amex Corp Feb'24 - Addisu Turi (IT).xlsx", "01/21/2024", "2/21/2024", macro_computer, "[02] Feb 2024")
# controller.prepare_template_workbook()
# controller.process_invoices_worksheet()
# controller.process_transaction_details_2_worksheet()
# controller.process_amex_transation_details_worksheet()
