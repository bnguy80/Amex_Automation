import os

import pandas as pd
from tabulate import tabulate
from tqdm import tqdm

from workbook_manager import TemplateWorkbookManager
from pdf_collection_manager import PDFCollectionManager
from invoice_transaction_matcher import manipulation


class AutomationController:
    XLOOKUP_TABLE_WORKSHEET_NAME = "Xlookup table"  # Make sure this is correct, 3/24/24: is correct inside Template - Master.xlsm 6/15/2024
    TEMPLATE_WORKBOOK_NAME = "Template - Master.xlsm"  # This is the workbook that we will be storing the intermediary data for matching AMEX Statement transactions and invoices for 6/15/2024.
    TEMPLATE_INVOICES_WORKSHEET_NAME = "Invoices"
    TEMPLATE_TRANSACTION_DETAILS_2_WORKSHEET_NAME = "Transaction Details 2"
    LIST_INVOICE_NAME_AND_PATH_MACRO_NAME = "ListFilesInSpecificFolder"  # Macro name to get invoice pdf file names and file_paths from the invoices folder 6/15/2024

    def __init__(self, amex_path, amex_statement_name, start_date, end_date, macro_parameter_1=None, macro_parameter_2=None):

        self.amex_path = amex_path  # The directory where the AMEX Statement workbook is located 6/16/2024
        self.amex_statement = amex_statement_name  # This is the final workbook that the automation will put the data into; sent to Ana 6/15/2024.

        # Start date of Amex Statement transactions
        # End date of Amex Statement transactions
        self.pdf_collection = PDFCollectionManager(start_date, end_date)
        self.manipulation = manipulation  # Using a list of strategies to match invoices to transactions. ONLY ONE INSTANCE 6/22/2024.

        self.template_workbook_path = os.path.join(self.amex_path, self.TEMPLATE_WORKBOOK_NAME)
        self.template_workbook_manager = TemplateWorkbookManager(self.TEMPLATE_WORKBOOK_NAME, self.template_workbook_path)

        self.amex_workbook_path = os.path.join(self.amex_path, self.amex_statement)
        # self.amex_workbook_manager = AmexWorkbookManager(self.amex_statement, self.amex_workbook_path)

        self.macro_parameter_1 = macro_parameter_1
        self.macro_parameter_2 = macro_parameter_2

    @staticmethod
    # This adds in the duplicate CLOUDFLARE transaction row, do not add in via manual manipulation 6/24/2024.
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
        invoice_worksheet = self.template_workbook_manager.get_worksheet(self.TEMPLATE_INVOICES_WORKSHEET_NAME)

        # This step is to get the Xlookup table worksheet to be able to get vendors for pdfs
        xlookup_table_worksheet = self.template_workbook_manager.get_worksheet(self.XLOOKUP_TABLE_WORKSHEET_NAME)
        self.pdf_collection.populate_pdf_collections_vendors_list_from_xlookup_worksheet(xlookup_table_worksheet)

        # This step populates the pdf_collection_dataframe with all the pdf data 6/16/2024
        self.pdf_collection.populate_pdf_collection_dataframe_from_worksheet(invoice_worksheet)
        pdf_collection_df = self.pdf_collection.get_pdf_collection_dataframe()
        num_updates = len(pdf_collection_df.index)
        progress_bar = tqdm(total=num_updates, desc="Updating Invoices Worksheet from Extracted PDF Data", bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt}')

        # print(pdf_data.columns)

        # Resets the invoice counter
        self.pdf_collection.reset_counter()

        invoice_worksheet.update_data_from_dataframe_to_sheet(pdf_collection_df, progress_bar)  # Updates the Invoice worksheet with all the necessary fields for each pdf invoice 6/15/2024

        progress_bar.close()

    def process_invoices_worksheet(self):
        # Get initial invoice names and invoice file paths for "Invoices" worksheet of Template workbook
        self.template_workbook_manager.workbook.call_macro_workbook(self.LIST_INVOICE_NAME_AND_PATH_MACRO_NAME, self.macro_parameter_1, self.macro_parameter_2)

        # Update Template workbook "Invoices" worksheet from Invoice PDF Collection DataFrame
        self.update_invoices_worksheet_with_all_extracted_data()

        # Save the changes
        self.template_workbook_manager.workbook.save()

    def process_transaction_details_worksheet(self) -> None:

        invoices_worksheet = self.template_workbook_manager.get_worksheet(self.TEMPLATE_INVOICES_WORKSHEET_NAME)
        transaction_details_worksheet = self.template_workbook_manager.get_worksheet(self.TEMPLATE_TRANSACTION_DETAILS_2_WORKSHEET_NAME)

        # Convert the Invoice worksheet into DataFrame
        invoices_worksheet_df = invoices_worksheet.read_data_as_dataframe()
        print("Invoices DataFrame Before Matching Process")
        print(tabulate(invoices_worksheet_df, headers='keys', tablefmt='psql'))
        print("\n")

        # Convert Transaction Details 2 worksheet into DataFrame
        transaction_details_worksheet_df = transaction_details_worksheet.read_data_as_dataframe()
        # Print the transaction details DataFrame before matching
        print("Transaction Details 2 DataFrame Before Matching Process:")
        print(tabulate(transaction_details_worksheet_df, headers='keys', tablefmt='psql'))
        print("\n")

        # Update with 'File path' column to the end if it is not already present.
        if 'File path' not in transaction_details_worksheet_df.columns:
            transaction_details_worksheet_df['File path'] = ''

        # Sets the preprocessed dataframes to the DataManipulation class to do further processing 6/19/2024.
        self.manipulation.set_data(invoices_worksheet_df, transaction_details_worksheet_df)

        # Matches invoice files found in "Invoices" worksheet to "Transaction Details 2" worksheet transactions
        # Works through 3 strategies of 1: exact matching between vendor|date|amount 2: match between vendor|amount|non-matching date or 3: target total between subset of transactions that sum to amount of invoice
        self.manipulation.find_matching_transactions()

        # Call static method to duplicate Cloudflare rows and update file names starting from index 8
        transaction_details_worksheet_df = self.duplicate_and_label_rows(transaction_details_worksheet_df)

        # Drop the 'File path' column from the transaction_details_worksheet_df 6/23/2024
        if 'File path' in transaction_details_worksheet_df.columns:
            transaction_details_worksheet_df = transaction_details_worksheet_df.drop('File path', axis=1)

        print("\n")
        print("Transaction Details 2 DataFrame After Matching and Duplication Process")
        print(tabulate(transaction_details_worksheet_df, headers='keys', tablefmt='psql'))
        print("\n")

        num_updates = len(transaction_details_worksheet_df.index)
        progress_bar = tqdm(total=num_updates, desc="Updating Transaction Details 2 Worksheet From Matched Invoices", bar_format='{l_bar}{bar}| {n_fmt}/{total_fmt}')

        # Resets the invoice counter
        self.pdf_collection.reset_counter()

        transaction_details_worksheet.update_data_from_dataframe_to_sheet(transaction_details_worksheet_df, progress_bar)

        progress_bar.close()


# "H:/Amex Automation" Automation Truth--> amex_path
# "C:/Users/brand/IdeaProjects/Amex Automation DATA" -computer
# r"H:\Amex Automation\t3nas\APPS\\" -Truth--> macro_parameter_1
# r"C:\Users\brand\IdeaProjects\Amex Automation DATA\t3nas\APPS\\" -computer
path_truth = "H:/Amex Automation"
path_computer = "C:/Users/brand/IdeaProjects/Amex Automation DATA"
macro_truth = r"H:\Amex Automation\t3nas\APPS\\"
macro_computer = r"C:\Users\brand\IdeaProjects\Amex Automation DATA\t3nas\APPS\\"

controller = AutomationController(path_computer, "Amex Corp Feb'24 - Addisu Turi (IT).xlsx", "01/21/2024", "2/21/2024", macro_computer, "[02] Feb 2024")  # Make sure to have "r" and \ at the end to treat as raw string parameter 6/15/2024

controller.process_invoices_worksheet()
controller.process_transaction_details_worksheet()
