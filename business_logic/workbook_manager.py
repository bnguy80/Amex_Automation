from abc import ABC, abstractmethod

from models.workbook import Workbook
from business_logic.update_strategies import TemplateInvoiceUpdateStrategy, TemplateTransactionDetails2UpdateStrategy, AmexTransactionDetailsUpdateStrategy


class WorkbookManager(ABC):

    def __init__(self, workbook_name: str, workbook_path: str):
        self.workbook_name = workbook_name
        self.workbook_path = workbook_path
        self.workbook = Workbook(workbook_path)

    @abstractmethod
    def select_worksheet_strategy(self, worksheet_name: str):
        ...

    def get_worksheet(self, worksheet_name: str):
        worksheet = self.workbook.get_worksheet(worksheet_name)
        strategy = self.select_worksheet_strategy(worksheet_name)
        worksheet.set_strategy(strategy)
        return worksheet


class TemplateWorkbookManager(WorkbookManager):

    def select_worksheet_strategy(self, worksheet_name: str):
        if worksheet_name == "Invoices":
            strategy = TemplateInvoiceUpdateStrategy()
        elif worksheet_name == "Transaction Details 2":
            strategy = TemplateTransactionDetails2UpdateStrategy()
        else:
            strategy = None

        return strategy


class AmexWorkbookManager(WorkbookManager):

    def select_worksheet_strategy(self, worksheet_name: str):
        if worksheet_name == "Transaction Details":
            strategy = AmexTransactionDetailsUpdateStrategy()
        else:
            strategy = None

        return strategy
