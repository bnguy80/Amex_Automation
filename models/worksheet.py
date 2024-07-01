import pandas as pd


class Worksheet:
    def __init__(self, name, sheet):
        self.name = name
        self.sheet = sheet
        self.worksheet_dataframe = pd.DataFrame()
        self.strategy = None

    # We will assume whichever sheet we're interacting with Invoices, Transactions Details 2, and so on the sheet.range starts at 'A7' 6/19/2024
    def read_data_as_dataframe(self):
        # Use xlwings to read data into a DataFrame, header True to interpret the first row as column headers for the dataframe, index=False to make sure the first column is not interpreted as an index column
        dataframe = self.worksheet_dataframe = self.sheet.range('A7').options(pd.DataFrame, expand='table', header=True, index=False).value

        if dataframe.empty:
            print("Could not read data from worksheet")
        else:
            return dataframe

    def set_strategy(self, strategy):
        self.strategy = strategy

    def update_sheet(self, data_df, progress_bar) -> None:
        self.strategy.update_worksheet(self, data_df, progress_bar)
