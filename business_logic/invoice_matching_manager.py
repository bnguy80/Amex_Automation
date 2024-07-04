import pandas as pd
from typing import Optional, List, Set, Hashable

from business_logic.matching_strategies import ExactAmountDateStrategy, ExactAmountAndExcludeDateStrategy, CombinationTotalStrategy, VendorOnlyStrategy

from utils.util_functions import print_dataframe


class InvoiceMatchingManager:
    """
    The `InvoiceMatchingManager`
    class is responsible for matching invoices with transaction details using different strategies.
    It allows setting invoice and transaction data, executing the invoice matching process, and sequencing File Names.

    Attributes
        - `invoice_df`: A DataFrame representing the invoice data.
        - `transaction_details_df`: A DataFrame representing the transaction details data.
        - `matched_transactions`: A set that tracks the matched transaction indexes.
        - `matched_invoices`: A set that tracks the matched invoice indexes.
        - `Primary_strategy`: A list containing the primary strategies used to match invoices and transactions.
        - `Fallback_strategy`: A strategy used to match unmatched invoices and empty transaction details File Names.

    Methods
        - `__init__(primary_strategy, fallback_strategy)`:
        Initializes the `InvoiceMatchingManager` instance with the provided primary and fallback strategies.
        - `Set_data(invoice_df, transaction_details_df) -> None`: Sets the invoice and transaction details data.
        - `Execute_invoice_matching()`: Executes the invoice matching process using the primary and fallback strategies.
        - `Sequence_file_names()`: Sequences the File Names starting from index 8 across the transaction details data.

    Example usage
    ```
    manager = InvoiceMatchingManager(primary_strategy, fallback_strategy)
    manager.set_data(invoice_df, transaction_details_df)
    manager.execute_invoice_matching()
    manager.sequence_file_names()
    ```

    Note: This documentation only provides an overview of the class and its methods.
    For more details about the primary_strategy, fallback_strategy,
    and other classes used within this class, refer to their respective documentation.
    """
    def __init__(self, primary_strategy, fallback_strategy):
        self.invoice_df: Optional[pd.DataFrame] = None
        self.transaction_details_df: Optional[pd.DataFrame] = None
        self.matched_transactions: Set[int] = set()  # This set will track matched transactions indexes.
        self.matched_invoices: Set[Hashable] = set()  # This set will track matched invoice indexes.

        # The strategies used to match invoices and transactions; each invoice will go through each strategy one by one until a match is found 6/20/2024.
        self.primary_strategy: List = primary_strategy
        # After the pass through of primary strategies to match invoices and transactions; match with a broader approach
        self.fallback_strategy = fallback_strategy

    def set_data(self, invoice_df: pd.DataFrame, transaction_details_df: pd.DataFrame) -> None:
        """
        Set the invoice and transaction details data.

        :param invoice_df: The dataframe containing the invoice data.
        :param transaction_details_df: The dataframe containing the transaction details data.
        :return: None
        """
        self.invoice_df: pd.DataFrame = invoice_df
        self.transaction_details_df: pd.DataFrame = transaction_details_df

    def execute_invoice_matching(self) -> None:
        """
        Executes the invoice matching process using primary and fallback strategies.

        :return: None
        """
        # First pass: Iterate over each invoice row and attempt to match using primary strategies
        for _, invoice_row in self.invoice_df.iterrows():
            # Try to find a match using each strategy in sequence
            for strategy in self.primary_strategy:
                if strategy.execute(invoice_row, self.transaction_details_df, self.matched_transactions, self.matched_invoices):
                    break  # If a match is found, break out of the loop and proceed to the next invoice

        # Second pass: Apply the fallback strategy only to unmatched invoices and where transaction_details_df "File name" is empty
        for _, invoice_row in self.invoice_df.iterrows():
            if self.fallback_strategy.execute(invoice_row, self.transaction_details_df, self.matched_transactions, self.matched_invoices):
                continue  # If a match is found, proceed to the next unmatched invoice after finding a match

        # After invoices that could've been matched are matched, print unmatched invoices
        unmatched_invoices_df = self.invoice_df.loc[~self.invoice_df.index.isin(self.matched_invoices)]
        if not unmatched_invoices_df.empty:
            print_dataframe(unmatched_invoices_df, "Unmatched Invoices:")

    def sequence_file_names(self) -> None:
        """
        Sequence File Names starting from index 8 across the transaction_details_df.

        :return: None
        """
        for i in range(len(self.transaction_details_df)):
            self.transaction_details_df.at[i, 'File Name'] = f"{8 + i} - {self.transaction_details_df.loc[i, 'File Name']}"


invoice_matching_manager = InvoiceMatchingManager([ExactAmountDateStrategy(), ExactAmountAndExcludeDateStrategy(), CombinationTotalStrategy()], VendorOnlyStrategy())
