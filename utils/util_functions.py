from tabulate import tabulate


def print_dataframe(df, message: str):
    print(message)
    print(tabulate(df, headers='keys', tablefmt='psql'))
    print("\n")
