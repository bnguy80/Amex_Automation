from tabulate import tabulate


from matching_strategies import ExactMatchStrategy, AmountAndNonDatesStrategy, CombinationStrategy


class InvoiceTransactionMatcher:

    def __init__(self, strategies):
        self.invoice_df = None
        self.transaction_details_df = None
        self.matched_transactions = set()  # This set will track matched transactions
        self.matched_invoices = set()  # This set will track matched invoice indices.

        # The strategies used to match invoices and transactions; each invoice will go through each strategy one by one until a match is found 6/20/2024.
        self.strategies = strategies

    def set_data(self, invoice_df, transaction_details_df) -> None:
        self.invoice_df = invoice_df
        self.transaction_details_df = transaction_details_df

    def find_matching_transactions(self):
        # Iterate over each invoice row
        for index, invoice_row in self.invoice_df.iterrows():
            # Try to find a match using each strategy in sequence
            for strategy in self.strategies:
                if strategy.execute(invoice_row, self.transaction_details_df, self.matched_transactions, self.matched_invoices):
                    break  # If a match is found, break out of the loop and proceed to the next invoice

        # After all invoices have been processed, print unmatched invoices
        unmatched_invoices = self.invoice_df.loc[~self.invoice_df.index.isin(self.matched_invoices)]
        if not unmatched_invoices.empty:
            print("\nUnmatched Invoices:")
            print(tabulate(unmatched_invoices, headers='keys', tablefmt='psql'))


manipulation = InvoiceTransactionMatcher([ExactMatchStrategy(), AmountAndNonDatesStrategy(), CombinationStrategy()])
