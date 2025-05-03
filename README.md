# Excel Merger

A Python utility to combine multiple Excel files into a single workbook, with each file's contents placed in a separate sheet named according to the original file name.

## Overview

This tool is designed to:
- Read multiple Excel files from a specified directory
- Create a new Excel workbook
- Copy the contents of each file into a separate sheet
- Name each sheet based on the original file name (removing numeric prefixes)
- Sort sheets based on the numeric prefix in the original filenames

## Features

- Preserves all cell values and formulas
- Maintains original formatting and styles
- Automatic sheet naming based on file names
- Progress tracking during processing
- Handles large numbers of Excel files efficiently

## Prerequisites

- Python 3.x
- openpyxl library

## Installation

1. Clone this repository:
```bash
git clone https://github.com/tejasprasadd/Multi-Excel-Combiner-Agent.git
cd Multi-Excel-Combiner-Agent
```

2. Create and activate a virtual environment (recommended):
```bash
python -m venv venv
# On Windows
venv\Scripts\activate
# On macOS/Linux
source venv/bin/activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Place all Excel files you want to combine in the `All-Excels` directory
2. Run the script:
```bash
python copy_excels_to_sheets.py
```

3. The script will:
   - Process all Excel files in the `All-Excels` directory
   - Create a new file called `Combined_Excels.xlsx`
   - Show progress for each file being processed

## File Naming Convention

The script expects Excel files to follow this naming pattern:
```
[number]-[sheet-name].xlsx
```
Example: `001-IR-E-U-0456.xlsx`

The resulting sheet in the combined workbook will be named: `IR-E-U-0456`

## Output

The script generates a single Excel file named `Combined_Excels.xlsx` containing:
- One sheet per input file
- Sheets named according to the original file names (without numeric prefixes)
- Sheets ordered based on the numeric prefix in the original filenames

## Example

Input files:
```
All-Excels/
├── 001-IR-E-U-0456.xlsx
├── 002-IR-E-I-1074.xlsx
└── 003-IR-E-U-0306.xlsx
```

Output file (`Combined_Excels.xlsx`):
```
Sheets:
├── IR-E-U-0456
├── IR-E-I-1074
└── IR-E-U-0306
```

## Contributing

Feel free to submit issues and enhancement requests!


## Author

B R Tejas Prasad