import os
import logging

import numpy as np

from excelexp.cli import create_parser
from excelexp.io import read_data, export_to_docx, export_to_pdf
from excelexp.transform import get_samples

# set logging
FORMAT = "%(asctime)-15s:  %(message)s"


if __name__ == "__main__":

    logging.basicConfig(level=logging.INFO, format=FORMAT)

    parser = create_parser()
    args = parser.parse_args()

    inputs = os.path.abspath(args.input)
    data = read_data(inputs)

    if args.filter:
        filter_value = int(args.filter)
        
        n_samples = int(args.nsamples)
        samples = get_samples(data, n_samples)
        filtered_data = samples.loc[samples["wartosc"] == filter_value]
        filtered_data["wartosc"] = np.nan
    else:
        filtered_data = samples
    
    destination = os.path.abspath(args.destination)
    os.makedirs(destination, exist_ok=True)

    try:
        export_to_docx(filtered_data, destination, n_samples)
        export_to_pdf(filtered_data, destination)
    except Exception as e:
        logging.error("Could not export data \n\n{}".format(e))
