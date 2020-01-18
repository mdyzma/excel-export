from typing import Callable
import pandas as pd


def filter_func():
    pass


def filter_rows(data: pd.DataFrame, filter_func: Callable = None) -> pd.DataFrame:
    filtered_data = data.apply(filter_func, data)
    return filtered_data


def get_samples(data: pd.DataFrame, n_samples: str) -> pd.DataFrame:
    n_samples = int(n_samples)
    samples = data.sample(n_samples)
    return samples
