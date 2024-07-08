from abc import ABC, abstractmethod
from typing import Union

import pandas as pd

from models.worksheet import Worksheet


class UpdateStrategy:
    """The Strategy interface."""

    @abstractmethod
    def update_worksheet(self, worksheet: Worksheet, data: Union[pd.DataFrame, Worksheet]):
        pass


class TemplateInvoiceUpdateStrategy(UpdateStrategy):

    def update_worksheet(self, worksheet: Worksheet, data: pd.DataFrame):
        # Assuming 'data_df' is a DataFrame with columns ['File Name', 'File Path', 'Amount', 'Vendor', 'Date']

        # Read the existing data from the worksheet into a DataFrame
        existing_data_df = worksheet.read_data_as_dataframe()

        # Update the rows in existing_data_df with the data in pdf_data based on matching 'File Path'
        for _, data_row in data.iterrows():
            # Match by 'File Path'
            mask = existing_data_df['File Path'] == data_row['File Path']

            # If the file_path exists in the existing_data_df, update the corresponding row
            if mask.any():
                existing_data_df.loc[mask, 'Amount'] = data_row['Amount']
                existing_data_df.loc[mask, 'Vendor'] = data_row['Vendor']
                existing_data_df.loc[mask, 'Date'] = data_row['Date']

        # Write the updated DataFrame back to the Excel sheet
        # This step overwrites the existing data starting from the specified cell range
        worksheet.sheet.range('A7').options(index=False).value = existing_data_df.reset_index(drop=True)


class TemplateTransactionDetails2UpdateStrategy(UpdateStrategy):

    def update_worksheet(self, worksheet: Worksheet, data: pd.DataFrame):

        start_row = 8  # Headers are in row 7, data starts at row 8
        last_row = start_row + len(data) - 1

        # Update the DataFrame directly to the Excel worksheet
        worksheet.sheet.range(f'A{start_row}').options(index=False, header=False).value = data

        # Account formula set in column 'E', Sub-Account in 'F', Vendor in 'G', Explanation in 'H'
        # Apply the formulas row by row
        for index in range(start_row, last_row + 1):
            # Account formula set in column 'E'
            worksheet.sheet.range(f'E{index}').formula = f'=XLOOKUP(G{index}, Table2[[#All],[Vendors]], Table2[[#All],[Account]],,0,1)'
            # Sub-Account formula set in column 'F'
            worksheet.sheet.range(f'F{index}').formula = f'=XLOOKUP(G{index}, Table2[[#All],[Vendors]], Table2[[#All],[Code]], "PLEASE REVIEW", 0, 1)'
            # Vendor formula set in column 'G'
            worksheet.sheet.range(f'G{index}').formula = f'=TEXTBEFORE(C{index}," ")'
            # Explanation formula set in column 'H'
            worksheet.sheet.range(f'H{index}').formula = f'=TEXTJOIN("/", TRUE, "Amex", "IT", G{index}, TEXTAFTER(I{index},"- "))'


class AmexTransactionDetailsUpdateStrategy(UpdateStrategy):

    def update_worksheet(self, worksheet: Worksheet, data: Worksheet):
        last_row = data.sheet.range('A' + str(data.sheet.cells.last_cell.row)).end('up').row
        data_range = f'A7:K{last_row}'  # Update column range as necessary
        worksheet_start_row = 'A7'
        data.sheet.range(data_range).copy(worksheet.sheet.range(worksheet_start_row))


