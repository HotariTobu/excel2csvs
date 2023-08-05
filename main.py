import argparse
import shutil

import pandas as pd

from pathlib import Path
from typing import List

def process(input_file_path: Path, output_dir_path: Path):
    """process a file

    Args:
        input_file_path (Path): path to input pdf
        output_dir_path (Path): path to output dir

    Returns:
        bool: True if succeed in processing file, otherwise False
    """

    excel_file = pd.ExcelFile(input_file_path)

    for sheet_name in excel_file.sheet_names:
        sheet_data = excel_file.parse(sheet_name)
        output_file_path = output_dir_path / f"{sheet_name}.csv"
        sheet_data.to_csv(output_file_path, index=False)

    return True

# Parse command-line parameters
parser = argparse.ArgumentParser(description='Convert Excel file into CSV files')
parser.add_argument('input_path', type=Path, help='path to the input Excel file or directory')
args = parser.parse_args()

input_path: Path = args.input_path
input_file_paths: List[Path] = []

if input_path.is_dir():
    for input_file_path in input_path.glob('*.xls?'):
        input_file_paths.append(input_file_path)
else:
    input_file_paths.append(input_path)

for input_file_path in input_file_paths:
    output_dir_path = input_file_path.with_suffix('')
    if output_dir_path.exists():
        print('Directory already exists:', output_dir_path)
        conf = input('Overwrite? (Y/n): ')
        if conf != 'Y':
            continue

        shutil.rmtree(output_dir_path)

    output_dir_path.mkdir()

    process(input_file_path, output_dir_path)
