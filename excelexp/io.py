import os
import logging
from datetime import datetime
from typing import List
from glob import glob

import docx
import pandas as pd
from tqdm import tqdm


def get_file_list(data_path: str) -> List[str]:
    return glob("/".join([data_path, "*.xlsx"]))


def read_data(data_path: str) -> pd.DataFrame:
    files_xls = get_file_list(data_path)
    df = pd.DataFrame()
    logging.info("Extracting data from {} files".format(len(files_xls)))
    for file in tqdm(files_xls):
        data = pd.read_excel(file, header=None, names=["tekst", "wartosc"])
        df = df.append(data)
    return df


def save_results(data: pd.DataFrame, destination: str, n_samples: int) -> None:
    doc = docx.Document()
    t = doc.add_table(data.shape[0] + 1, data.shape[1])

    time_now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    fname = "-".join(["results", time_now, ".docx"])
    path = os.path.join(destination, fname)

    logging.info("Writing {} data samples to {}".format(n_samples, destination))

    # add the header rows.
    for j in range(data.shape[-1]):
        t.cell(0, j).text = data.columns[j]

    # add the rest of the data frame
    for i in tqdm(range(data.shape[0])):
        for j in range(data.shape[-1]):
            t.cell(i + 1, j).text = str(data.values[i, j])

    doc.save(path)
