import time

import pandas as pd

from tabulate import tabulate
from tqdm import tqdm


def print_dataframe(df: pd.DataFrame, message: str):
    print(message)
    print(tabulate(df, headers='keys', tablefmt='psql'))
    print("\n")


class ProgressTrackingMixin:

    def __init__(self):
        self.progress_bar = None

    def start_progress_tracking(self, total_steps: int, description: str = ""):
        if self.progress_bar is None:
            self.progress_bar = tqdm(total=total_steps, desc=description)

    def update_progress(self, steps: int = 1):
        if self.progress_bar is not None:
            self.progress_bar.update(steps)
            time.sleep(0.005)

    def complete_progress(self):
        if self.progress_bar is not None:
            self.progress_bar.close()
            self.progress_bar = None
            time.sleep(1)
