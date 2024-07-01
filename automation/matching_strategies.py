from abc import abstractmethod, ABC
from itertools import combinations
from typing import Tuple

import numpy as np
import pandas as pd


class MatchingStrategy(ABC):

    # Static method can't be overridden by implementations 6/27/2024
    @staticmethod
    def load_data(invoice_row: pd.Series) -> Tuple:
        """
        Load data from the invoice_df row
        :param invoice_row: pd.Series
        :return: Tuple(vendor, total, date, file_name, file_path)
        """
        vendor = invoice_row['Vendor']
        total = invoice_row['Amount']
        date = pd.to_datetime(invoice_row['Date'])
        file_name = invoice_row['File Name']
        file_path = invoice_row['File Path']

        return vendor, total, date, file_name, file_path

    @abstractmethod
    def execute(self, invoice_row, transaction_details_df, matched_transactions, matched_invoices):
        pass


class ExactMatchStrategy(MatchingStrategy):
    def execute(self, invoice_row, transaction_details_df, matched_transactions, matched_invoices):
        vendor, total, date, file_name, file_path = self.load_data(invoice_row)

        potential_matches: pd.DataFrame = transaction_details_df[
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False)) &
            (~transaction_details_df.index.isin(matched_transactions))
            ]

        # Strategy 1: Exact match on Amount and Date
        exact_matches: pd.DataFrame = potential_matches[
            (potential_matches['Amount'] == total) &
            (pd.to_datetime(potential_matches['Date'], errors='coerce') == date)
            ]

        if not exact_matches.empty:
            first_match_index = exact_matches.iloc[0].name
            transaction_details_df.at[first_match_index, 'File name'] = file_name
            transaction_details_df.at[first_match_index, 'Column1'] = "Amount & Date Match"
            transaction_details_df.at[first_match_index, 'File path'] = file_path
            matched_transactions.add(first_match_index)
            matched_invoices.add(invoice_row.name)  # Use invoice_row.name if 'name' is the index or unique identifier

            print(f"Match found for Strategy 1 in transaction with id {first_match_index}!")
            return True

        return False


class ExactAmountAndExcludeDateStrategy(MatchingStrategy):

    def execute(self, invoice_row, transaction_details_df, matched_transactions, matched_invoices):
        vendor, total, date, file_name, file_path = self.load_data(invoice_row)

        # Filter potential matches by vendor that match to invoice, ensuring they aren't previously matched in the matched_transactions set
        potential_matches: pd.DataFrame = transaction_details_df[
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False)) &  # Matches vendor name, case insensitive
            (~transaction_details_df.index.isin(matched_transactions))  # Excludes transactions already matched
            ]

        # Strategy 2: Match on Amount and Excludes Date
        non_date_matches: pd.DataFrame = potential_matches[
            (potential_matches['Amount'] == total)  # Matches exact amount
        ]

        if not non_date_matches.empty:
            first_match_index = non_date_matches.iloc[0].name
            transaction_details_df.at[first_match_index, 'File name'] = file_name
            transaction_details_df.at[first_match_index, 'Column1'] = 'Amount & Non-Date Match'
            transaction_details_df.at[first_match_index, 'File path'] = file_path
            matched_transactions.add(first_match_index)
            matched_invoices.add(invoice_row.name)  # Use invoice_row.name if 'name' is the index or unique identifier

            print(f"Match found for Strategy 2 in transaction with id {first_match_index}!")
            return True

        return False


class CombinationStrategy(MatchingStrategy):

    def execute(self, invoice_row, transaction_details_df, matched_transactions, matched_invoices):

        vendor, total, date, file_name, file_path = self.load_data(invoice_row)

        # Filter potential invoice matches by vendor and exact date, excluding those already matched in the matched_transactions set
        potential_matches: pd.DataFrame = transaction_details_df[
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False)) &  # Matches vendor name, case insensitive
            (~transaction_details_df.index.isin(matched_transactions)) &  # Excludes transactions that are already in matched_transactions
            (pd.to_datetime(transaction_details_df['Date'], errors='coerce') == date)  # Matches exact date
            ]

        # Find combinations of transactions where the sum equals the invoice amount
        for r in range(1, min(4,
                              len(potential_matches) + 1)):  # Limiting to combinations of up to 3 for complexity management
            for combo in combinations(potential_matches.itertuples(index=True), r):
                if np.isclose(sum(item.Amount for item in combo), total, atol=0.01):
                    # If a valid combination is found, mark all involved transactions
                    for item in combo:
                        idx = item.Index
                        transaction_details_df.at[idx, 'File name'] = file_name
                        transaction_details_df.at[idx, 'Column1'] = 'Combination Amount Match'
                        transaction_details_df.at[idx, 'File path'] = file_path
                        matched_transactions.add(idx)
                        matched_invoices.add(
                            invoice_row.name)  # Assuming 'name' is the DataFrame index or unique identifier

                        print(f"Match found for Strategy 3 in transactions with ids {[item.Index for item in combo]}!")
                        return True  # Stop after finding the first valid combination
        return False


class VendorOnlyStrategy(MatchingStrategy):

    def execute(self, invoice_row, transaction_details_df, matched_transactions, matched_invoices):

        if invoice_row.name in matched_invoices:
            return False  # Skip processing if the invoice has already been matched

        vendor, total, date, file_name, file_path = self.load_data(invoice_row)

        # Filter for potential matches where 'File name' field is empty, indicating they haven't been matched yet
        potential_matches = transaction_details_df[
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False)) &
            (transaction_details_df['File name'].isnull()) &
            (~transaction_details_df.index.isin(matched_transactions))
            ]

        vendor_only_matches = potential_matches[
            (potential_matches['Vendor'] == vendor)
        ]

        if not vendor_only_matches.empty:
            first_match_index = vendor_only_matches.iloc[0].name
            transaction_details_df.at[first_match_index, 'File name'] = file_name
            transaction_details_df.at[first_match_index, 'Column1'] = 'Vendor Only'
            transaction_details_df.at[first_match_index, 'File path'] = file_path
            matched_transactions.add(first_match_index)
            matched_invoices.add(invoice_row.name)

            print(f"Match found for Vendor Only strategy in transaction with id {first_match_index}!")
            return True

        return False
