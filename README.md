# CellColorUnifier
# Excel Color Unifier

This Python script is designed to unify cell colors in Excel spreadsheets (.xlsx) by mapping specific color codes to standard colors—red, blue, green, and orange. It utilizes the `openpyxl` library for handling Excel files and provides a function to modify cell colors based on given themes.

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [How it Works](#how-it-works)
- [License](#license)
- [Acknowledgments](#acknowledgments)

## Features

- Extracts Excel files from a zip folder.
- Iterates through each cell and identifies its color.
- Changes identified colors to specific shades of red, blue, green, or orange based on predefined color codes.
- Saves modified Excel files in the same format.

## Requirements

- Python 3.x
- `openpyxl` library
- `zipfile` module (part of the standard library)

Install the `openpyxl` library using pip:

```bash
pip install openpyxl
```

## Installation

1. Clone this repository to your local machine:
   ```bash
   git clone https://github.com/your-username/excel-color-unifier.git
   ```

2. Navigate to the project directory:
   ```bash
   cd excel-color-unifier
   ```

3. Ensure you have the required packages installed.

## Usage

1. Prepare an Excel file or a zip folder containing multiple Excel files named `sample files.zip` in the root directory of the project.
2. Run the script:
   ```bash
   python your_script.py
   ```
3. The script will save the modified Excel files in their original format within the same directory as the input files.

## How it Works

- The script uses `openpyxl` to load workbooks.
- Each cell's color is checked against predefined lists of color codes for red, blue, green, and orange.
- If a cell matches a specific shade, it is changed to the designated standard color.
- The modified workbooks are then saved back to their respective files.

### Important Note

The predefined color codes are specific to the Excel theme "Office 2013-2022". If using a different theme, you may need to adjust the color codes accordingly.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Acknowledgments

- The codebase includes contributions from the following source: [Mike Honey GitHub Gist](https://gist.github.com/Mike-Honey/b36e651e9a7f1d2e1d60ce1c63b9b633).
- Thanks to the [openpyxl](https://openpyxl.readthedocs.io/en/stable/) library for Python, which allows easy manipulation of Excel files. 

Feel free to contribute or raise issues if you encounter any problems!
