#   
#
# Script for merging xlsx file by Romi Syuhada aka Ni'am H Sahid
# Thanks to copilot, i don't write this actually, it's only giving idea and prompting. I love AI..!!!!
# 
# You will need to install pandas and openpyxl
# Windows :
# py -m pip install pandas
# py -m pip install openpyxl
#
# 

import sys

# Check if pandas and openpyxl are installed
try:
    import pandas as pd
except ImportError:
    print("Error: pandas is not installed. Please install it using 'pip install pandas'.")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("Error: openpyxl is not installed. Please install it using 'pip install openpyxl'.")
    sys.exit(1)

import os
import argparse

def merge_excel_files(files, output_file, verbose=False):
    # Ensure output file path is a string
    output_file = str(output_file)

    # Convert relative paths to absolute paths
    output_file = os.path.abspath(output_file)

    print(f"Output file: {output_file}")

    # Read and concatenate all specified Excel files
    dataframes = []
    for file in files:
        file_path = os.path.abspath(file)
        try:
            df = pd.read_excel(file_path)
            dataframes.append(df)
            print(f"Successfully read {file_path}")
            if verbose:
                print(f"Contents of {file_path}:")
                print(df)
        except Exception as e:
            print(f"Error reading {file_path}: {e}")

    if dataframes:
        try:
            # Merge all DataFrames
            merged_df = pd.concat(dataframes, ignore_index=True)
            # Save the merged DataFrame to a new Excel file
            merged_df.to_excel(output_file, index=False)
            print(f"Files merged successfully into {output_file}")
            if verbose:
                print("Merged DataFrame:")
                print(merged_df)
        except Exception as e:
            print(f"Error merging files: {e}")
    else:
        print("No Excel files to merge.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Merge Excel files.')
    parser.add_argument('-d', '--directory', help='Directory containing Excel files to merge')
    parser.add_argument('-f', '--files', action='append', help='Individual Excel files to merge', default=[])
    parser.add_argument('-o', '--output', required=True, help='Output file name for the merged Excel file')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose mode to display file contents')

    args = parser.parse_args()

    if args.directory:
        # List all .xlsx files in the directory
        directory = os.path.abspath(args.directory)
        excel_files = [os.path.join(directory, f) for f in os.listdir(directory) if f.endswith('.xlsx')]
    else:
        excel_files = args.files

    if not excel_files:
        print("No Excel files to merge. Please specify files with -f or a directory with -d.")
    else:
        merge_excel_files(excel_files, args.output, args.verbose)

