# **Nessus** Scan Output CSV to Excel Spreadsheet Converter

Outputs from **Nessus** scans in **CSV** format are a bit hard to read. For this reason, I made this here little script that outputs results into an Excel spreadsheet, using a python script. 

## Features

- **Escapes newline characters**: Removes any unwanted line breaks within cells.
- **Column sorting**: Sorts rows by the `cvss` column in descending order (if available).
- **Header styling**: Applies a bold, slightly darker blue style to the header row.
- **Alternating row colors**: Alternates between white and light blue rows for better readability.

## Prerequisites

Ensure you have Python 3+ installed on your system. Required Python libraries:
- `pandas`
- `xlsxwriter`

Install dependencies with:
```
pip install pandas xlsxwriter 
```


## Usage

```
python script.py <csv_file> <excel_file>
```

### Example: 
```
python script.py raw_data.csv formatted_output.xlsx
```

### Arguments: 
- `<csv_file>`: Path to the input CSV file
- `<excel_file>`: Path to the output excel file

## Recommendations 
- Use .xlsx as the output file extension to ensure compatibility and avoid file extension warnings.
- Ensure the output file is not open in another program (like Excel) to prevent permission errors.

## Help 
``` 
python script.py --help
```
