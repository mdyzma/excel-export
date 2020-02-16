# -*- coding: utf-8 -*-
from typing import Callable
import pandas as pd


def filter_rows(data: pd.DataFrame, filter_func: Callable = None) -> pd.DataFrame:
    filtered_data = data.apply(filter_func, data)
    return filtered_data


def get_samples(data: pd.DataFrame, n_samples: int) -> pd.DataFrame:
    samples = data.head(n_samples)
    return samples
