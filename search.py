"""
Script: search.py

Description:
This script allows users to search for a specified string or regular expression pattern within the names or contents of text files (.txt), CSV files (.csv), and Excel files (.xlsx) within a directory and its subdirectories. It can ignore case sensitivity, use regular expressions, or focus on file names only for its search criteria.

Usage:
python search_tool.py [options] <search_string>

Options:
  -i, --ignore-case    Ignore case sensitivity when searching.
  -r, --regex          Use regular expressions for searching.
  -f, --file-name      Search for the string or pattern within file names only, not their contents.
  -d, --directory     Directory to start the search from (default: current directory)

Arguments:
  <search_string>      The text or regular expression pattern to search for within the files or their names.

Supported File Types:
- Text Files (*.txt)
- CSV Files (*.csv)
- Excel Files (*.xlsx)

Permissions:
- If the script encounters a file for which it doesn't have the necessary permissions to read, it will print a message indicating the permission issue and skip processing that file.

Error Handling:
- BadZipFile: If the script encounters an Excel file that is not a valid workbook, it will print a message indicating the issue and skip processing that file.

Note:
- The script recursively searches all files within the current directory and its subdirectories.
- Ensure that the Python interpreter has the necessary permissions to read the files you intend to search or their names.

Using Regex with File Names:
- When using the `--regex` option in conjunction with the `--file-name` option, the script will apply the regex pattern to match against the names of files. This allows for complex search patterns, such as finding files with specific formats or patterns in their names.

Example:
- To find all Excel files that have a date format (YYYY-MM-DD) in their file name, you could use the following command:
  `python search.py -r -f "\\d{4}-\\d{2}-\\d{2}" --file-name`
  This uses a regex pattern (`\\d{4}-\\d{2}-\\d{2}`) to search for file names that match a specific date format.

"""

import os
import sys
import csv
import re
import openpyxl
import argparse
from zipfile import BadZipFile

def search_txt(file_path, search_string, ignore_case, use_regex, search_in_file_name=False):
    if search_in_file_name:
        if matches_search_criteria(os.path.basename(file_path), search_string, ignore_case, use_regex):
            print(f"File name match found: {file_path}")
        return  # Skip content search when searching by file name
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            for line_num, line in enumerate(file, 1):
                if matches_search_criteria(line, search_string, ignore_case, use_regex):
                    print(f"Found in {file_path} on line {line_num}")
    except PermissionError:
        print(f"Permission denied for {file_path}. Skipping...")

def search_csv(file_path, search_string, ignore_case, use_regex, search_in_file_name=False):
    if search_in_file_name:
        if matches_search_criteria(os.path.basename(file_path), search_string, ignore_case, use_regex):
            print(f"File name match found: {file_path}")
        return  # Skip content search when searching by file name
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as file:
            reader = csv.reader(file)
            for row_num, row in enumerate(reader, 1):
                for cell in row:
                    if matches_search_criteria(cell, search_string, ignore_case, use_regex):
                        print(f"Found in {file_path} in row {row_num}")
    except PermissionError:
        print(f"Permission denied for {file_path}. Skipping...")

def search_xlsx(file_path, search_string, ignore_case, use_regex, search_in_file_name=False):
    if search_in_file_name:
        if matches_search_criteria(os.path.basename(file_path), search_string, ignore_case, use_regex):
            print(f"File name match found: {file_path}")
        return  # Skip content search when searching by file name
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        for sheet in wb:
            for row_num, row in enumerate(sheet.iter_rows(), 1):
                for cell in row:
                    cell_value = str(cell.value)
                    if matches_search_criteria(cell_value, search_string, ignore_case, use_regex):
                        print(f"Found in {file_path} in sheet {sheet.title} at cell {cell.coordinate}")
    except PermissionError:
        print(f"Permission denied for {file_path}. Skipping...")
    except BadZipFile:
        print(f"{file_path} is not a valid Excel workbook. Skipping...")

def matches_search_criteria(text, search_string, ignore_case, use_regex):
    if text is None:
        return False
    text = str(text)
    return (use_regex and re.search(search_string, text, flags=re.I if ignore_case else 0)) \
           or (ignore_case and search_string.lower() in text.lower()) \
           or (search_string in text)

def main():
    parser = argparse.ArgumentParser(description='Search for text in files with directory selection.')
    parser.add_argument('search_string', help='The text or pattern to search for')
    parser.add_argument('-i', '--ignore-case', action='store_true', help='Ignore case sensitivity when searching')
    parser.add_argument('-r', '--regex', action='store_true', help='Use regular expressions for searching')
    parser.add_argument('-f', '--file-name', action='store_true', help='Search in file names only')
    parser.add_argument('-d', '--directory', default='.', help='Directory to start the search from (default: current directory)')
    
    args = parser.parse_args()

    # Convert relative path to absolute path
    search_dir = os.path.abspath(args.directory)
    
    if not os.path.exists(search_dir):
        print(f"Error: Directory '{search_dir}' does not exist.")
        sys.exit(1)
        
    print(f"Searching in directory: {search_dir}")
    print(f"Search string: {args.search_string}")
    print(f"Options: ignore_case={args.ignore_case}, regex={args.regex}, file_name_only={args.file_name}")
    
    for dirpath, dirnames, filenames in os.walk(search_dir):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            if args.file_name:
                search_txt(file_path, args.search_string, args.ignore_case, args.regex, args.file_name)
            else:
                if filename.endswith('.txt'):
                    search_txt(file_path, args.search_string, args.ignore_case, args.regex)
                elif filename.endswith('.csv'):
                    search_csv(file_path, args.search_string, args.ignore_case, args.regex)
                elif filename.endswith(('.xlsx', '.xlsm')):
                    search_xlsx(file_path, args.search_string, args.ignore_case, args.regex)

if __name__ == "__main__":
    main()