import pandas as pd
from typing import Optional, List, Set, Hashable

from business_logic.matching_strategies import MatchingStrategy, ExactAmountDateStrategy, ExactAmountAndExcludeDateStrategy, CombinationTotalStrategy, VendorOnlyStrategy

from utils.utilities import print_dataframe, ProgressTrackingMixin


class InvoiceMatchingManager(ProgressTrackingMixin):
    """
    The `InvoiceMatchingManager`
    class is responsible for matching invoices with transaction details using different strategies.
    It allows setting invoice and transaction data, executing the invoice matching process, and sequencing File Names.

    Attributes
        - `invoice_df`: A DataFrame representing the invoice data.
        - `transaction_details_df`: A DataFrame representing the transaction details data.
        - `matched_transactions`: A set that tracks the matched transaction indexes.
        - `matched_invoices`: A set that tracks the matched invoice indexes.
        - `primary_strategies`: A list containing the primary strategies used to match invoices and transactions.
        - `fallback_strategy`: A strategy used to match unmatched invoices and empty transaction details File Names.

    Methods
        - `__init__(primary_strategies, fallback_strategy, **kwargs)`: Initializes the `InvoiceMatchingManager` instance with the provided primary and fallback strategies.
        - `Set_data(invoice_df, transaction_details_df) -> None`: Sets the invoice and transaction details data.
        - `Execute_invoice_matching()`: Executes the invoice matching process using the primary and fallback strategies.
        - `Sequence_file_names()`: Sequences the File Names starting from index 8 across the transaction details data.

    Example usage
    ```
    manager = InvoiceMatchingManager(primary_strategies, fallback_strategy)
    manager.set_data(invoice_df, transaction_details_df)
    manager.execute_invoice_matching()
    manager.sequence_file_names()
    ```

    Note: This documentation only provides an overview of the class and its methods.
    For more details about the primary_strategy, fallback_strategy,
    and other classes used within this class, refer to their respective documentation.
    """
    def __init__(self, primary_strategies: List[MatchingStrategy], fallback_strategy: MatchingStrategy, **kwargs):
        super().__init__(**kwargs)  # Making sure that parameters aren't consumed by other classes through inheritance--> MRO 7/8/2024
        self.invoice_df: Optional[pd.DataFrame] = None
        self.transaction_details_df: Optional[pd.DataFrame] = None
        self.matched_transactions: Set[int] = set()  # This set will track matched transactions indexes.
        self.matched_invoices: Set[Hashable] = set()  # This set will track matched invoice indexes.

        # The strategies used to match invoices and transactions; each invoice will go through each strategy one by one until a match is found 6/20/2024.
        self.primary_strategy: List[MatchingStrategy] = primary_strategies
        # After the pass through of primary strategies to match invoices and transactions; match with a broader approach
        self.fallback_strategy: MatchingStrategy = fallback_strategy

    def set_data(self, invoice_df: pd.DataFrame, transaction_details_df: pd.DataFrame) -> None:
        """
        Set the invoice and transaction details data.

        :param invoice_df: The dataframe containing the invoice data.
        :param transaction_details_df: The dataframe containing the transaction details data.
        :return: None
        """
        self.invoice_df: pd.DataFrame = invoice_df
        self.start_progress_tracking(total_steps=len(invoice_df.index), description="Matching Invoices")
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
                    self.update_progress()
                    break  # If a match is found, break out of the loop and proceed to the next invoice

        # Second pass: Apply the fallback strategy only to unmatched invoices and where transaction_details_df "File name" is empty
        for _, invoice_row in self.invoice_df.iterrows():
            if self.fallback_strategy.execute(invoice_row, self.transaction_details_df, self.matched_transactions, self.matched_invoices):
                self.update_progress()
                continue  # If a match is found, proceed to the next unmatched invoice after finding a match

        # After invoices that could've been matched are matched, print unmatched invoices
        unmatched_invoices_df = self.invoice_df.loc[~self.invoice_df.index.isin(self.matched_invoices)]
        if not unmatched_invoices_df.empty:
            print_dataframe(unmatched_invoices_df, "Unmatched Invoices:")

        self.complete_progress()

    def sequence_file_names(self) -> None:
        """
        Sequence File Names starting from index 8 across the transaction_details_df.

        :return: None
        """
        self.start_progress_tracking(len(self.transaction_details_df), description="Sequencing File Names in Transaction Details 2 Dataframe:")
        for i in range(len(self.transaction_details_df)):
            self.transaction_details_df.at[i, 'File Name'] = f"{8 + i} - {self.transaction_details_df.loc[i, 'File Name']}"
            self.update_progress()

        self.complete_progress()


invoice_matching_manager = InvoiceMatchingManager([ExactAmountDateStrategy(), ExactAmountAndExcludeDateStrategy(), CombinationTotalStrategy()], VendorOnlyStrategy())
