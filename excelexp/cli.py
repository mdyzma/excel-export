# -*- coding: utf-8 -*-
import os
import argparse
from datetime import datetime


def default_output_path() -> str:

    home = os.path.expanduser("~")
    time_now = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return os.path.join(home, "exls_exports", time_now)


def create_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input", "-i", help="Ścieżka do folderu z plikami xls")
    parser.add_argument(
        "--destination",
        "-d",
        default=default_output_path(),
        help="Folder do zapiasania wynikow",
    )
    parser.add_argument(
        "--nsamples",
        "-s",
        default=10,
        help="Liczba wierszy wylosowana losowo z calego datasetu.",
    )
    parser.add_argument(
        "--filter",
        "-f",
        help="Liczba do filtrowania datasetu.",
    )
    parser.add_argument(
        "--files",
        type=argparse.FileType('r'),
        nargs='+',
        help="Lista plikow.",
    )

    return parser
