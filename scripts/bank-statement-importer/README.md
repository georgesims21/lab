# Bank Statement Importer

A utility for combining multiple CSV bank statement files and uploading them to Google Sheets.

## Features

- Combine multiple CSV files into one, preserving headers from a specified file
- Upload CSV data to a Google Sheet, appending to existing content
- Flexible operation modes: concatenate only, upload only, or both

## Prerequisites

- Python 3.6+
- Required packages:
  - pandas
  - google-api-python-client
  - google-auth-oauthlib
  - google-auth

## Setup

### Using Virtual Environment (recommended)

1. Create a virtual environment:
```bash
python -m venv venv
```

2. Activate the virtual environment:
```bash
# On Windows
venv\Scripts\activate

# On macOS/Linux
source venv/bin/activate
```

3. Install the required packages:
```bash
pip install pandas google-api-python-client google-auth-oauthlib google-auth
```

### Direct Installation

If not using a virtual environment, install the required packages using pip:
```bash
pip install pandas google-api-python-client google-auth-oauthlib google-auth
```

## Google Sheets API Setup

1. Create a Google Cloud project
2. Enable the Google Sheets API
3. Create OAuth 2.0 credentials and download as `credentials.json`
4. Place `credentials.json` in the same directory as the script

## Usage

```bash
python import-statements.py [DIRECTORY] [OPTIONS]
```

### Basic Arguments

- `DIRECTORY`: Path to the directory containing CSV files to process

### Operation Modes

- `-u, --upload`: Concatenate files and upload the result to Google Sheets (default combined behavior)
- `--concat-only`: Only concatenate CSV files without uploading
- `--upload-only`: Only upload a specific file without concatenation

### Output Options

- `-o, --output`: Output filename for the combined CSV (default: 'combined.csv')
- `-f, --header-file`: Specific CSV file to use for the header row (default: first CSV file found)

### Upload Options

- `-s, --spreadsheet-id`: Google Spreadsheet ID to upload data to
- `-t, --sheet-name`: Tab/Sheet name in the spreadsheet
- `-i, --input-file`: Specific CSV file to upload (required when using --upload-only)

## Examples

### Combine CSV files only

```bash
python import-statements.py ~/bank_statements --concat-only -o combined_statements.csv
```

### Upload an existing CSV file to Google Sheets

```bash
python import-statements.py . --upload-only -i path/to/file.csv -s your_spreadsheet_id -t SheetName
```

### Combine CSV files and upload to Google Sheets (traditional usage)

```bash
python import-statements.py ~/bank_statements -u -s your_spreadsheet_id -t SheetName
```

## Notes

- When uploading to Google Sheets, data is appended to any existing content in the sheet
- The first time you run the script with upload functionality, it will open a browser window for authentication
- Authentication tokens are saved in `token.json` for future use
