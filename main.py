# -*- coding: utf-8 -*-
import os
import logging

import numpy as np

from excelexp.cli import create_parser
from excelexp.io import read_data, \
    export_to_docx, \
    export_to_pdf, \
    get_file_list
from excelexp.transform import get_samples

# set logging
FORMAT = "%(asctime)-15s:  %(message)s"


if __name__ == "__main__":

    logging.basicConfig(level=logging.INFO, format=FORMAT)

    parser = create_parser()
    args = parser.parse_args()

    print(args)

    if args.files:
        inputs = [os.path.abspath(f.name) for f in args.files]
        data = read_data(inputs)
    else:
        inputs = os.path.abspath(args.input)
        excl_files = get_file_list(inputs)
        data = read_data(excl_files)

    n_samples = int(args.nsamples)

    if args.filter:
        filter_value = int(args.filter)

        samples = get_samples(data, n_samples)
        filtered_data = samples.loc[samples["wartosc"] == filter_value]
    else:
        filtered_data = get_samples(data, n_samples)

    destination = os.path.abspath(args.destination)
    os.makedirs(destination, exist_ok=True)

    try:
        number_of_rows = len(filtered_data)
        export_to_docx(filtered_data, destination, number_of_rows)
        export_to_pdf(filtered_data, destination)
    except Exception as e:
        logging.error("Could not export data \n\n{}".format(e))
