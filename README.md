# File Search Tool

This Python script provides a powerful way to search for a specified string or regular expression pattern within the names or contents of text, CSV, and Excel files. It's designed to work within a directory and its subdirectories, offering options to ignore case sensitivity, use regular expressions, or focus solely on file names for searching.

## Features

- Search through text files, CSV files, and Excel files for specific text or regex patterns.
- Option to ignore case sensitivity.
- Ability to use regular expressions for more complex search patterns.
- Search can be conducted within file names only, excluding their contents.

## Installation

No installation is required beyond having Python installed on your system. This script was developed and tested with Python 3.8. Ensure you have Python 3.x installed before running the script.

## Usage

To use the script, navigate to the directory containing `search.py` and run it with Python, followed by the options and search string you wish to use.

`python search.py [options] <search_string>`

### Options

- `-i`, `--ignore-case`: Ignore case sensitivity when searching.
- `-r`, `--regex`: Use regular expressions for searching.
- `-f`, `--file-name`: Search for the string or pattern within file names only, not their contents.

### Example

To search for all Excel files with a date format (YYYY-MM-DD) in their file name:

`python search.py -r -f "\d{4}-\d{2}-\d{2}" --file-name`

## Supported File Types

- Text Files (*.txt)
- CSV Files (*.csv)
- Excel Files (*.xlsx)

## Permissions

The script will print a message and skip files that it does not have permission to read.

## Error Handling

- Files that are not valid Excel workbooks will be skipped with a message indicating the issue.

## Contributing

Contributions to the script are welcome. Please fork the repository, make your changes, and submit a pull request.

## License

[MIT License](LICENSE.txt)

## Acknowledgments

- Thanks to all contributors who have helped to improve this script.

