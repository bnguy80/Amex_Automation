import xlwings as xw

from models.worksheet import Worksheet

from typing import Optional


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
        Add a worksheet to the workbook that hasn't already been added into the worksheets' dict.
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
        Save to a specific path when specified or save at the current location

        :param save_path: String
        :return: None
        """
        if save_path:
            self.workbook.save(save_path)
        else:
            self.workbook.save()

    def call_macro_workbook(self, macro_name, macro_parameter_1: Optional[str] = None, macro_parameter_2: Optional[str] = None):
        macro_vba = self.workbook.app.macro(macro_name)
        if macro_parameter_1 and macro_parameter_2:
            macro_vba(macro_parameter_1, macro_parameter_2)
        else:
            macro_vba()
