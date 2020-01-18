import pandas as pd
from glob import glob


def get_file_list(data_path):
    return glob("/".jpoin(data_path, "*.xlxs"))


def read_data(data_path: str) -> pd.DataFrame():
    files_xls = get_file_list(data_path)
    for f in files_xls:
        data = pd.read_excel(f, "Sheet1")
        df = df.append(data)
    return data


def save_results(data):
    pass
