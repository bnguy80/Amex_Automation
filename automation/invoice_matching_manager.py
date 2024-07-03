from matching_strategies import ExactMatchStrategy, ExactAmountAndExcludeDateStrategy, CombinationStrategy, \
    VendorOnlyStrategy

from utils.util_functions import print_dataframe


class InvoiceMatchingManager:

    def __init__(self, primary_strategy, fallback_strategy):
        self.invoice_df = None
        self.transaction_details_df = None
        self.matched_transactions = set()  # This set will track matched transactions
        self.matched_invoices = set()  # This set will track matched invoice indexes.

        # The strategies used to match invoices and transactions; each invoice will go through each strategy one by one until a match is found 6/20/2024.
        self.primary_strategy = primary_strategy
        # After the pass through of primary strategies to match invoices and transactions; match with a broader approach
        self.fallback_strategy = fallback_strategy

    def set_data(self, invoice_df, transaction_details_df) -> None:
        self.invoice_df = invoice_df
        self.transaction_details_df = transaction_details_df

    def execute_invoice_matching(self):
        # First pass: Iterate over each invoice row and attempt to match using primary strategies
        for index, invoice_row in self.invoice_df.iterrows():
            # Try to find a match using each strategy in sequence
            for strategy in self.primary_strategy:
                if strategy.execute(invoice_row, self.transaction_details_df, self.matched_transactions, self.matched_invoices):
                    break  # If a match is found, break out of the loop and proceed to the next invoice

        # Second pass: Apply the fallback strategy only to unmatched invoices and where transaction_details_df "File name" is empty
        for index, invoice_row in self.invoice_df.iterrows():
            if self.fallback_strategy.execute(invoice_row, self.transaction_details_df, self.matched_transactions, self.matched_invoices):
                continue  # If a match is found, proceed to the next unmatched invoice after finding a match

        # After invoices that could've been matched are matched, print unmatched invoices
        unmatched_invoices_df = self.invoice_df.loc[~self.invoice_df.index.isin(self.matched_invoices)]
        if not unmatched_invoices_df.empty:
            print_dataframe(unmatched_invoices_df, "Unmatched Invoices:")

    def sequence_file_names(self):
        # Sequence file names starting from index 8 across the transaction_details_df
        for i in range(len(self.transaction_details_df)):
            self.transaction_details_df.at[i, 'File Name'] = f"{8 + i} - {self.transaction_details_df.loc[i, 'File Name']}"


invoice_matching_manager = InvoiceMatchingManager([ExactMatchStrategy(), ExactAmountAndExcludeDateStrategy(), CombinationStrategy()], VendorOnlyStrategy())
