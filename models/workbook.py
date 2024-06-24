import pandas as pd
import xlwings as xw


class Worksheet:
    def __init__(self, name, sheet):
        self.name = name
        self.sheet = sheet
        self.worksheet_dataframe = pd.DataFrame()
        self.strategy = None

    # We will assume whichever sheet we are interacting with Invoices, Transactions Details 2, etc. the sheet.range starts at 'A7' 6/19/2024
    def read_data_as_dataframe(self):
        # Use xlwings to read data into a DataFrame, header True to interpret first row as column headers for the dataframe, index=False to make sure the first column is not interpreted as an index column
        dataframe = self.worksheet_dataframe = self.sheet.range('A7').options(pd.DataFrame, expand='table', header=True, index=False).value

        if dataframe.empty:
            print("Could not read data from worksheet")
        else:
            return dataframe

    def set_strategy(self, strategy):
        self.strategy = strategy

    # # We will assume whichever sheet we are interacting with Invoices, Transactions Details 2, etc. the sheet.range starts at 'A7' 6/19/2024
    # def update_data_from_dataframe_to_sheet(self, data_df, progress_bar) -> None:
    #     # Assuming 'data_df' is a DataFrame with columns ['File Name', 'File Path', 'Amount', 'Vendor', 'Date']
    #
    #     # Read the existing data from worksheet into a dataframe
    #     existing_data_df = self.read_data_as_dataframe()
    #
    #     # Update the rows in existing_data_df with the data in pdf_data based on matching 'File Path'
    #     for _, data_row in data_df.iterrows():
    #         file_path = data_row['File Path']
    #         mask = existing_data_df['File Path'] == file_path
    #
    #         # If the file_path exists in the existing_data_df, update the corresponding row
    #         if mask.any():
    #             existing_data_df.loc[mask, 'Amount'] = data_row['Amount']
    #             existing_data_df.loc[mask, 'Vendor'] = data_row['Vendor']
    #             existing_data_df.loc[mask, 'Date'] = data_row['Date']
    #
    #         progress_bar.update(1)
    #
    #     # Write the updated DataFrame back to the Excel sheet
    #     # This will overwrite the existing data starting from the top-left cell where your data begins (e.g., 'A7')
    #     self.sheet.range('A7').options(index=False).value = existing_data_df.reset_index(drop=True)

    # We will assume whichever sheet we are interacting with Invoices, Transactions Details 2, etc. the sheet.range starts at 'A7' 6/19/2024

    def update_data_from_dataframe_to_sheet(self, data_df, progress_bar) -> None:
        self.strategy.update_worksheet(self, data_df, progress_bar)


class Workbook:

    def __init__(self, workbook_path=None):
        self.worksheets = {}

        if workbook_path is None:
            print("Workbook not found.")
        else:
            self.workbook = xw.Book(workbook_path)
            self.workbook_name = self.workbook.name

        # Automatically add all existing worksheets
        for sheet in self.workbook.sheets:
            self.worksheets[sheet.name] = Worksheet(sheet.name, sheet)

    def add_worksheet(self, worksheet_name: str) -> None:
        """
        Add a worksheet to the workbook that hasn't already been added into the worksheets dict.
        """
        if worksheet_name not in self.worksheets:
            sheet = self.workbook.sheets.add(worksheet_name)
            self.worksheets[worksheet_name] = Worksheet(worksheet_name, sheet)
        else:
            print(f"Worksheet '{worksheet_name}' already exists.")

    def remove_worksheet(self, worksheet_name) -> None:
        if worksheet_name in self.worksheets:
            self.workbook.sheets[worksheet_name].delete()
            del self.worksheets[worksheet_name]
        else:
            print(f"Worksheet '{worksheet_name}' not found.")

    def remove_all_worksheets_dict(self) -> None:
        self.worksheets = {}

    def get_worksheet(self, worksheet_name):
        worksheet = self.worksheets.get(worksheet_name)
        if worksheet is None:
            print(f"Worksheet '{worksheet_name}' not found in workbook.")
        return worksheet

    def get_all_worksheets(self) -> dict:
        worksheets_dict = self.worksheets

        if worksheets_dict is None:
            print("Could not return worksheets")
        else:
            return worksheets_dict

    def save(self, save_path: str = None) -> None:
        """
        Save to a specific path when specified, or just save at the current location

        :param save_path: String
        :return: None
        """
        if save_path:
            self.workbook.save(save_path)
        else:
            self.workbook.save()

    def call_macro_workbook(self, macro_name, macro_parameter_1: None, macro_parameter_2: None):
        macro_vba = self.workbook.app.macro(macro_name)
        macro_vba(macro_parameter_1, macro_parameter_2)
