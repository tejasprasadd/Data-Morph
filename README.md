# DataMorph - NIRF Data Processing System

A comprehensive toolkit for extracting, transforming, and analyzing data from National Institutional Ranking Framework (NIRF) reports for educational institutions in India.

## Project Overview

The NIRF Data Processing System is designed to automate the extraction of structured data from NIRF PDF reports, convert it to standardized formats (XML, Excel), and enable advanced analytics through MongoDB integration. This system helps educational administrators, researchers, and policymakers to efficiently process and analyze institutional performance data across multiple metrics.

## Features

- **PDF Data Extraction**: Extract structured data directly from NIRF PDF reports
- **Multiple Output Formats**: Convert data to XML or Excel formats
- **Data Normalization**: Clean and standardize extracted data
- **MongoDB Integration**: Import processed data into MongoDB for advanced analytics
- **Batch Processing**: Process multiple institution reports in a single operation
- **Comprehensive Data Coverage**: Extract data on intake, student demographics, placements, PhD programs, financial resources, and more

## Directory Structure

```
├── All-Pdfs/              # Directory containing source PDF files
├── All-XML/               # Directory for XML output files
├── All-Excels/            # Directory containing Excel files
├── Normalized-Excels/     # Directory for normalized Excel files
├── pdf-to-xml-agent.py    # Converts PDFs directly to XML format
├── pdf-to-excel-agent.py  # Converts PDFs to Excel format
├── excel-to-xml-agent.py  # Converts Excel files to XML format
├── xml_to_mongodb.py      # Imports XML data into MongoDB
├── normalize_excel_sheets.py # Normalizes data in Excel sheets
├── copy_excels_to_sheets.py # Combines Excel files into a single workbook
├── requirements.txt       # Python dependencies
└── README_MONGODB.md      # MongoDB-specific documentation
```

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/tejasprasadd/nirf-data-processing-system.git
   cd nirf-data-processing-system
   ```

2. Install required dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Prepare your environment:
   - Create the required directories if they don't exist:
     ```
     mkdir -p All-Pdfs All-XML All-Excels Normalized-Excels
     ```
   - Place your NIRF PDF files in the `All-Pdfs` directory

## Usage

### PDF to XML Conversion

```bash
python pdf-to-xml-agent.py
```

This script:
- Reads PDF files from the `All-Pdfs` directory
- Extracts structured data from NIRF reports
- Creates a hierarchical XML structure that preserves data relationships
- Outputs XML files to the `All-XML` directory

### PDF to Excel Conversion

```bash
python pdf-to-excel-agent.py
```

This script:
- Processes PDF files from the `All-Pdfs` directory
- Extracts data tables and structured information
- Converts the data into Excel format
- Outputs Excel files with standardised structure

### Excel to XML Conversion

```bash
python excel-to-xml-agent.py
```

This script:
- Reads Excel files from the `All-Excels` directory
- Converts the data into a hierarchical XML structure
- Preserves data relationships and metadata
- Creates XML files with standardised format

### Importing Data to MongoDB

```bash
python xml_to_mongodb.py
```

This script:
- Processes all XML files in the `All-XML` directory
- Connects to MongoDB (default: localhost:27017)
- Creates a database named `nirf_database`
- Creates collections for individual institutions and a master database

For detailed MongoDB instructions, see [README_MONGODB.md](README_MONGODB.md).

### Data Normalisation and Combination

```bash
python normalize_excel_sheets.py
python copy_excels_to_sheets.py
```

These scripts:
- Normalize data across Excel sheets for consistency
- Combine multiple Excel files into a single workbook for easier analysis

## Data Structure

The system extracts and processes the following data categories from NIRF reports:

1. **Sanctioned Intake**: Program-wise approved student intake by year
2. **Student Strength**: Demographics including gender, geographic origin, and economic background
3. **Placement Data**: Employment statistics, median salary, and higher education metrics
4. **PhD Data**: Full-time and part-time PhD student counts and graduation rates
5. **Financial Resources**: Capital and operational expenditure details
6. **Sponsored Projects**: Research funding information
7. **Consultancy Projects**: Industry collaboration metrics
8. **Facilities**: Infrastructure and accessibility information
9. **Faculty Count**: Teaching staff statistics

## Comparison of Approaches

### PDF-to-XML Benefits:
- Direct conversion eliminates intermediate steps, reducing potential for data loss
- Maintains original data structure from the source document
- Avoids Excel-specific formatting or calculation issues
- More efficient for batch processing of multiple documents
- Better for preserving text-based content like descriptions and notes

### Excel-to-XML Benefits:
- Excel data is already structured in rows and columns, simplifying conversion
- Excel files may have already undergone data cleaning and normalization
- Better for data that has been manually verified or enhanced
- Easier to handle complex calculations or derived values
- May preserve formatting and styling information better

## Requirements

- Python 3.6+
- pdfplumber>=0.7.0
- pandas>=1.3.0
- openpyxl>=3.0.10
- pymongo>=4.0.0 (for MongoDB integration)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- National Institutional Ranking Framework (NIRF) for providing the source data
- All contributors who have helped to improve this system
