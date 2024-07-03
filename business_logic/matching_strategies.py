from abc import abstractmethod, ABC
from itertools import combinations
from typing import Tuple, Hashable, Set

import numpy as np
import pandas as pd


class MatchingStrategy(ABC):

    @abstractmethod
    def execute(self, invoice_row: pd.Series, transaction_details_df: pd.DataFrame, matched_transactions: set, matched_invoices: set):
        pass

    # Static method can't be overridden by implementations 6/27/2024
    @staticmethod
    def _load_invoice_data(invoice_row: pd.Series) -> Tuple:
        """
        Load data from the invoice_df row

        :param invoice_row: Pd.Series
        :return: Tuple(vendor, total, date, file_name, file_path)
        """
        vendor: str = invoice_row['Vendor']
        total: float = invoice_row['Amount']
        date: pd.Timestamp = pd.to_datetime(invoice_row['Date'])
        file_name: str = invoice_row['File Name']
        file_path: str = invoice_row['File Path']

        return vendor, total, date, file_name, file_path

    @staticmethod
    def _add_match(transaction_details_df: pd.DataFrame, found_match_index: int, file_name: str, file_path: str, match_type: str, matched_transactions: Set[int], matched_invoices: Set[int], invoice_row_index: Hashable) -> None:
        """
        Update the transaction_details_df with the found match data.

        :param transaction_details_df: DataFrame of transaction details.
        :param found_match_index: Index of the matched transaction.
        :param file_name: File name of the matched invoice.
        :param file_path: File path of the matched invoice.
        :param match_type: Type of the match.
        :param matched_transactions: Set of matched transactions.
        :param matched_invoices: Set of matched invoices.
        :param invoice_row_index: Index of the invoice row.
        :return: None
        """
        transaction_details_df.at[found_match_index, 'File Name'] = file_name
        transaction_details_df.at[found_match_index, 'Column1'] = match_type
        transaction_details_df.at[found_match_index, 'File Path'] = file_path
        matched_transactions.add(found_match_index)
        matched_invoices.add(invoice_row_index)


class ExactAmountDateStrategy(MatchingStrategy):
    """
    This concrete class extends `MatchingStrategy` and implements the `execute` method.
    The `execute` method defined in this class uses the exact amount and date for matching.

    Methods
        - `execute()` -> bool: Executes the matching strategy based on the exact amount and date match.
    """
    def execute(self, invoice_row: pd.Series, transaction_details_df: pd.DataFrame, matched_transactions: Set[int], matched_invoices: Set[int]) -> bool:
        """
        Executes the matching strategy based on the exact amount and date match.

        :param invoice_row: A pd.Series representing the invoice row data.
        :param transaction_details_df: A pd.DataFrame representing the transaction details data.
        :param matched_transactions: A set containing the indexes of already matched transactions.
        :param matched_invoices: A set containing the indexes of already matched invoices.
        :return: A boolean indicating whether a match was found.
        """
        vendor, total, date, file_name, file_path = self._load_invoice_data(invoice_row)

        found_match: pd.DataFrame = transaction_details_df[
            (~transaction_details_df.index.isin(matched_transactions)) &  # Excludes transactions already matched
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False)) &  # Flexible, case-insensitive matching
            (transaction_details_df['Amount'] == total) &
            (pd.to_datetime(transaction_details_df['Date'], errors='coerce') == date)
        ]

        if not found_match.empty:
            found_match_index = found_match.iloc[0].name
            invoice_row_index = invoice_row.name
            self._add_match(transaction_details_df, found_match_index, file_name, file_path, 'Exact Amount and Date Match', matched_transactions, matched_invoices, invoice_row_index)

            print(f"Match Found For ExactAmountDateStrategy In Transaction With ID {found_match_index}!")
            return True

        return False


class ExactAmountAndExcludeDateStrategy(MatchingStrategy):
    """
    This concrete class extends `MatchingStrategy` and implements the `execute` method.
    The `execute` method defined in this class uses the exact amount and excludes the date for matching.

    Methods
        - `execute()` -> bool:  Executes the matching strategy based on the exact amount and excluding date.
    """
    def execute(self, invoice_row: pd.Series, transaction_details_df: pd.DataFrame, matched_transactions: Set[int], matched_invoices: Set[int]) -> bool:
        """
        Executes the matching strategy based on the exact amount and excluding date.

        :param invoice_row: A pd.Series representing the invoice row data.
        :param transaction_details_df: A pd.DataFrame representing the transaction details data.
        :param matched_transactions: A set containing the indexes of already matched transactions.
        :param matched_invoices: A set containing the indexes of already matched invoices.
        :return: A boolean indicating whether a match was found.
        """
        vendor, total, date, file_name, file_path = self._load_invoice_data(invoice_row)

        # Filter potential matches by vendor that match to invoice, ensuring they aren't previously matched in the matched_transactions set
        found_match: pd.DataFrame = transaction_details_df[
            (~transaction_details_df.index.isin(matched_transactions)) &  # Excludes transactions already matched
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False)) &  # Matches vendor name, case insensitive
            (transaction_details_df['Amount'] == total)
        ]

        if not found_match.empty:
            found_match_index = found_match.iloc[0].name
            invoice_row_index = invoice_row.name
            self._add_match(transaction_details_df, found_match_index, file_name, file_path, 'Amount and Exclude Date Match', matched_transactions, matched_invoices, invoice_row_index)

            print(f"Match Found For ExactAmountAndExcludeDateStrategy In Transaction With ID {found_match_index}!")
            return True

        return False


class CombinationTotalStrategy(MatchingStrategy):
    """
    This concrete class extends `MatchingStrategy` and implements the `execute` method.
    The `execute` method defined in this class uses a combination of transactions that,
    when summed up, matches the amount.

    Methods
        - `execute()` -> bool:
        Executes the matching strategy based on the date and combination of summed amount to match invoice.
    """
    def execute(self, invoice_row: pd.Series, transaction_details_df: pd.DataFrame, matched_transactions: Set[int], matched_invoices: Set[int]) -> bool:
        """
        Executes the matching strategy based on the date and combination of summed amount to match invoice.

        :param invoice_row: A pd.Series representing the invoice row data.
        :param transaction_details_df: A pd.DataFrame representing the transaction details data.
        :param matched_transactions: A set containing the indexes of already matched transactions.
        :param matched_invoices: A set containing the indexes of already matched invoices.
        :return: A boolean indicating whether a match was found.
        """
        vendor, total, date, file_name, file_path = self._load_invoice_data(invoice_row)

        # Filter potential invoice matches by vendor and exact date, excluding those already matched in the matched_transactions set
        potential_matches: pd.DataFrame = transaction_details_df[
            (~transaction_details_df.index.isin(matched_transactions)) &  # Excludes transactions that are already in matched_transactions
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False)) &  # Matches vendor name, case insensitive
            (pd.to_datetime(transaction_details_df['Date'], errors='coerce') == date)  # Matches exact date
        ]

        # Find combinations of transactions where the sum equals the invoice amount
        for r in range(1, min(4, len(potential_matches) + 1)):  # Limiting to combinations of up to 3 for complexity management
            for combo in combinations(potential_matches.itertuples(index=True), r):
                if np.isclose(sum(item.Amount for item in combo), total, atol=0.01):
                    # If a valid combination is found, mark all involved transactions
                    for item in combo:
                        found_match_index = item.Index
                        invoice_row_index = invoice_row.name
                        self._add_match(transaction_details_df, found_match_index, file_name, file_path, 'Combination Total Match', matched_transactions, matched_invoices, invoice_row_index)

                        print(f"Match Found For CombinationTotalStrategy In Transactions With IDs {[item.Index for item in combo]}!")
                        return True  # Stop after finding the first valid combination
        return False


class VendorOnlyStrategy(MatchingStrategy):
    """
    This concrete class extends `MatchingStrategy` and implements the `execute` method.
    The `execute` method is defined in this class using only the vendor for matching.

    Methods
        - `execute()` -> bool: Executes the matching strategy based only on the vendor.
    """
    def execute(self, invoice_row: pd.Series, transaction_details_df: pd.DataFrame, matched_transactions: Set[int], matched_invoices: Set[int]) -> bool:
        """
        Executes the matching strategy based only on the vendor.
        If the invoice has already been matched, the method returns False to skip processing.

        :param invoice_row: A pd.Series representing the invoice row data.
        :param transaction_details_df: A pd.DataFrame representing the transaction details data.
        :param matched_transactions: A set containing the indexes of already matched transactions.
        :param matched_invoices: A set containing the indexes of already matched invoices.
        :return: A boolean indicating whether a match was found.
        """
        # Not sure why I need to have this right now for this class, because right now I have this strategy 'continue'
        # in InvoiceTransactionManager when a match is found and move on to the next invoice.
        # But it is using the same invoice again in some cases if I don't include this section 7/1/2024.
        if invoice_row.name in matched_invoices:
            return False  # Skip processing if the invoice has already been matched

        vendor, total, date, file_name, file_path = self._load_invoice_data(invoice_row)

        found_match = transaction_details_df[
            (~transaction_details_df.index.isin(matched_transactions)) &
            (transaction_details_df['File Name'].isnull()) &  # Filter for potential matches where the 'File name' field is empty, indicating they haven't been matched yet
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False))  # Matches vendor name, case insensitive
        ]

        if not found_match.empty:
            found_match_index = found_match.iloc[0].name
            invoice_row_index = invoice_row.name
            self._add_match(transaction_details_df, found_match_index, file_name, file_path, 'Vendor Only Match', matched_transactions, matched_invoices, invoice_row_index)

            print(f"Match Found For VendorOnlyStrategy In Transaction With ID {found_match_index}!")
            return True

        return False
