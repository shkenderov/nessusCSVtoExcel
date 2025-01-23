import pandas as pd
from xlsxwriter.utility import xl_rowcol_to_cell
import argparse
import os

def csv_to_excel_with_pandas(csv_file_path, excel_file_path):
    try:
        # Validate file paths
        if not os.path.isfile(csv_file_path):
            raise FileNotFoundError(f"The file '{csv_file_path}' does not exist.")

        # Read the CSV file into a DataFrame
        df = pd.read_csv(csv_file_path)
        
        # Escape newline characters in all cells

        #deprecated
        #df = df.applymap(lambda x: str(x).replace('\n', ' ') if isinstance(x, str) else x)
        
        df = df.apply(lambda col: col.str.replace('\n', ' ') if col.dtype == 'O' else col)



        # Sort the DataFrame by the 'cvss' column in descending order
        if 'CVSS v2.0 Base Score' in df.columns:
            df = df.sort_values(by='CVSS v2.0 Base Score', ascending=False)
        
        # Save the DataFrame to an Excel file with alternating row colors
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Sheet1', startrow=1, header=False)
            
            # Get the workbook and worksheet objects
            workbook = writer.book
            worksheet = writer.sheets['Sheet1']

            # Define formats for header and rows
            header_format = workbook.add_format({
                'bg_color': '#B0C4DE',  # Slightly darker blue for header
                'bold': True,
                'border': 1
            })
            white_format = workbook.add_format({'bg_color': '#FFFFFF'})
            light_blue_format = workbook.add_format({'bg_color': '#DDEBF7'})

            # Write header with format
            worksheet.set_row(0, None, header_format)
            for col_idx, value in enumerate(df.columns):
                worksheet.write(0, col_idx, value, header_format)

            # Apply alternating colors to rows (starting after the header)
            for row_idx in range(1, len(df) + 1):
                fmt = light_blue_format if row_idx % 2 == 0 else white_format
                worksheet.set_row(row_idx, None, fmt)

        print(f"Successfully saved to '{excel_file_path}' with alternating row colors.")
    except FileNotFoundError as e:
        print(e)
    except PermissionError:
        print(f"Permission for '{excel_file_path}' denied. Check if the file already exists and is open.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == "__main__":
    # Set up argument parser
    parser = argparse.ArgumentParser(
        description="Convert a CSV file to an Excel file with alternating row colors. Recommended file extension for the output is '.xlsx'.",
        usage="python script.py <csv_file> <excel_file>"
    )
    parser.add_argument("csv_file", help="Path to the input CSV file.")
    parser.add_argument("excel_file", help="Path to the output Excel file.")

    # Parse arguments
    try:
        args = parser.parse_args()

        # Validate input arguments
        if not os.path.isfile(args.csv_file):
            print(f"Error: The input file '{args.csv_file}' does not exist or the path is incorrect.")
            print("Call with --help to see usage information.")
            exit(1)

        if not args.excel_file.endswith('.xlsx'):
            print("Warning: It is recommended to use the '.xlsx' extension for the output file.")

        # Run the conversion function with the provided arguments
        csv_to_excel_with_pandas(args.csv_file, args.excel_file)
    except SystemExit:
        print("Error: Invalid arguments provided. Call with --help to see usage information.")
        raise
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
