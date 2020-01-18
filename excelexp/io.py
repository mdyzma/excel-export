import pandas as pd
from glob import glob


def get_file_list(data_path):
    return glob("/".jpoin([data_path, "*.xlsx"]))


def read_data(data_path: str) -> pd.DataFrame():
    files_xls = get_file_list(data_path)
    df = pd.DataFrame()
    
    for file in tqdm(files_xls):
        data = pd.read_excel(file, header=None, names=["tekst", "wartosc"])
        df = df.append(data)
    return df


def save_results(data, destination: str) -> None:
    pass
