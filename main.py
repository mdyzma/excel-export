from excelexp.io import read_data, save_results
from excelexp.transform import get_samples
from excelexp.cli import create_parser


if __name__ == "__main__":

    parser = create_parser()
    args = parser.parse_args()

    data = read_data(args.input, args.nsamples)

    sample = get_samples(data)
    save
