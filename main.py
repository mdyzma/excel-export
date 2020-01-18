import os
import logging

from excelexp.cli import create_parser
from excelexp.io import read_data, save_results
from excelexp.transform import get_samples  # , filter_rows, filter_func

# set logging
FORMAT = '%(asctime)-15s:  %(message)s'


if __name__ == "__main__":

    logging.basicConfig(level=logging.INFO, format=FORMAT)
    
    parser = create_parser()
    args = parser.parse_args()

    inputs = os.path.abspath(args.input)
    data = read_data(inputs)

    n_samples = int(args.nsamples)
    samples = get_samples(data, n_samples)

    # filtered_data = filter_rows(samples, filter_func=filter_func)

    destination = os.path.abspath(args.destination)
    os.makedirs(destination, exist_ok=True)

    save_results(samples, destination)
