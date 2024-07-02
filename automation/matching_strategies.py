from abc import abstractmethod, ABC
from itertools import combinations
from typing import Tuple

import numpy as np
import pandas as pd


class MatchingStrategy(ABC):

    @abstractmethod
    def execute(self, invoice_row, transaction_details_df, matched_transactions, matched_invoices):
        pass

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


class ExactMatchStrategy(MatchingStrategy):
    def execute(self, invoice_row, transaction_details_df, matched_transactions, matched_invoices):

        vendor, total, date, file_name, file_path = self.load_data(invoice_row)

        found_match: pd.DataFrame = transaction_details_df[
            (~transaction_details_df.index.isin(matched_transactions)) &  # Excludes transactions already matched
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False)) &  # Flexible, case-insensitive matching
            (transaction_details_df['Amount'] == total) &
            (pd.to_datetime(transaction_details_df['Date'], errors='coerce') == date)
        ]

        if not found_match.empty:
            first_match_index = found_match.iloc[0].name
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
        found_match: pd.DataFrame = transaction_details_df[
            (~transaction_details_df.index.isin(matched_transactions)) &  # Excludes transactions already matched
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False)) &  # Matches vendor name, case insensitive
            (transaction_details_df['Amount'] == total)
        ]

        if not found_match.empty:
            first_match_index = found_match.iloc[0].name
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
                        idx = item.Index
                        transaction_details_df.at[idx, 'File name'] = file_name
                        transaction_details_df.at[idx, 'Column1'] = 'Combination Amount Match'
                        transaction_details_df.at[idx, 'File path'] = file_path
                        matched_transactions.add(idx)
                        matched_invoices.add(invoice_row.name)  # Assuming 'name' is the DataFrame index or unique identifier

                        print(f"Match found for Strategy 3 in transactions with ids {[item.Index for item in combo]}!")
                        return True  # Stop after finding the first valid combination
        return False


class VendorOnlyStrategy(MatchingStrategy):

    def execute(self, invoice_row, transaction_details_df, matched_transactions, matched_invoices):

        # Not sure why I need to have this right now for this class, because right now I have this strategy 'continue'
        # in InvoiceTransactionManager when a match is found and move on to the next invoice.
        # But it is using the same invoice again in some cases if I don't include this section 7/1/2024.
        if invoice_row.name in matched_invoices:
            return False  # Skip processing if the invoice has already been matched

        vendor, total, date, file_name, file_path = self.load_data(invoice_row)

        found_match = transaction_details_df[
            (~transaction_details_df.index.isin(matched_transactions)) &
            (transaction_details_df['File name'].isnull()) &  # Filter for potential matches where the 'File name' field is empty, indicating they haven't been matched yet
            (transaction_details_df['Vendor'].str.contains(vendor, case=False, na=False))  # Matches vendor name, case insensitive
        ]

        if not found_match.empty:
            first_match_index = found_match.iloc[0].name
            transaction_details_df.at[first_match_index, 'File name'] = file_name
            transaction_details_df.at[first_match_index, 'Column1'] = 'Vendor Only'
            transaction_details_df.at[first_match_index, 'File path'] = file_path
            matched_transactions.add(first_match_index)
            matched_invoices.add(invoice_row.name)

            print(f"Match found for Vendor Only strategy in transaction with id {first_match_index}!")
            return True

        return False
