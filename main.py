import os

from excelexp.cli import create_parser
from excelexp.io import read_data, save_results
from excelexp.transform import get_samples, filter_rows, filter_func


if __name__ == "__main__":

    parser = create_parser()
    args = parser.parse_args()

    inputs = os.path.abspath(args.input)
    data = read_data(inputs)

    n_samples = int(args.nsamples)
    samples = get_samples(data, n_samples)

    filtered_data = filter_rows(samples, filter_func=filter_func)

    destination = os.path.abspath(args.destination)
    save_results(filtered_data, destination)
