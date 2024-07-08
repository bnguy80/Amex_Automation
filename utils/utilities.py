from tabulate import tabulate
from tqdm import tqdm


def print_dataframe(df, message: str):
    print(message)
    print(tabulate(df, headers='keys', tablefmt='psql'))
    print("\n")


def show_progress(steps: int, message: str):
    progress_bar = tqdm(total=steps, desc=message)
