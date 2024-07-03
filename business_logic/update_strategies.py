from abc import ABC, abstractmethod


class UpdateStrategy:
    """The Strategy interface."""

    @abstractmethod
    def update_worksheet(self, worksheet, data_df, progress_bar):
        pass


class InvoiceUpdateStrategy(UpdateStrategy):

    def update_worksheet(self, worksheet, data_df, progress_bar):
        # Assuming 'data_df' is a DataFrame with columns ['File Name', 'File Path', 'Amount', 'Vendor', 'Date']

        # Read the existing data from the worksheet into a DataFrame
        existing_data_df = worksheet.read_data_as_dataframe()

        # Update the rows in existing_data_df with the data in pdf_data based on matching 'File Path'
        for _, data_row in data_df.iterrows():
            # Match by 'File Path'
            mask = existing_data_df['File Path'] == data_row['File Path']

            # If the file_path exists in the existing_data_df, update the corresponding row
            if mask.any():
                existing_data_df.loc[mask, 'Amount'] = data_row['Amount']
                existing_data_df.loc[mask, 'Vendor'] = data_row['Vendor']
                existing_data_df.loc[mask, 'Date'] = data_row['Date']

            progress_bar.update(1)

        # Write the updated DataFrame back to the Excel sheet
        # This step overwrites the existing data starting from the specified cell range
        worksheet.sheet.range('A7').options(index=False).value = existing_data_df.reset_index(drop=True)


class TransactionDetails2UpdateStrategy(UpdateStrategy):

    def update_worksheet(self, worksheet, data_df, progress_bar):
        # Account formula set in column 'E', Sub-Account in 'F', Vendor in 'G', Explanation in 'H'

        start_row = 8  # Headers are in row 7, data starts at row 8
        last_row = start_row + len(data_df) - 1

        # Update the DataFrame directly to the Excel worksheet
        worksheet.sheet.range(f'A{start_row}').options(index=False, header=False).value = data_df

        # Apply the formulas row by row
        for idx in range(start_row, last_row + 1):
            # Account formula set in column 'E'
            worksheet.sheet.range(f'E{idx}').formula = f'=XLOOKUP(G{idx}, Table2[[#All],[Vendors]], Table2[[#All],[Account]],,0,1)'
            # Sub-Account formula set in column 'F'
            worksheet.sheet.range(f'F{idx}').formula = f'=XLOOKUP(G{idx}, Table2[[#All],[Vendors]], Table2[[#All],[Code]], "PLEASE REVIEW", 0, 1)'
            # Vendor formula set in column 'G'
            worksheet.sheet.range(f'G{idx}').formula = f'=TEXTBEFORE(C{idx}," ")'
            # Explanation formula set in column 'H'
            worksheet.sheet.range(f'H{idx}').formula = f'=TEXTJOIN("/", TRUE, "Amex", "IT", G{idx}, TEXTAFTER(I{idx},"- "))'

            # Update the progress bar after each row is processed
            progress_bar.update(1)
