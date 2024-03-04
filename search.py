import os
import sys
import csv
import re
import openpyxl
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
    return (use_regex and re.search(search_string, text, flags=re.I if ignore_case else 0)) \
           or (ignore_case and search_string.lower() in text.lower()) \
           or (search_string in text)

def main():
    ignore_case = False
    use_regex = False
    search_in_file_name = False
    search_string = ""

    # Parsing the command-line arguments
    for arg in sys.argv[1:]:
        if arg == "--ignore-case" or arg == "-i":
            ignore_case = True
        elif arg == "--regex" or arg == "-r":
            use_regex = True
        elif arg == "--file-name" or arg == "-f":
            search_in_file_name = True
        else:
            search_string = arg

    if not search_string:
        print("Please provide a search string or regex pattern as an argument.")
        sys.exit(1)

    for dirpath, dirnames, filenames in os.walk("."):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            if search_in_file_name:
                search_txt(file_path, search_string, ignore_case, use_regex, search_in_file_name)
            else:
                if filename.endswith('.txt'):
                    search_txt(file_path, search_string, ignore_case, use_regex)
                elif filename.endswith('.csv'):
                    search_csv(file_path, search_string, ignore_case, use_regex)
                elif filename.endswith('.xlsx'):
                    search_xlsx(file_path, search_string, ignore_case, use_regex)

if __name__ == "__main__":
    main()
