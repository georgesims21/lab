# Bank Statement Importer

A Python utility for combining multiple CSV bank statement files and optionally uploading them to Google Sheets.

## Overview

This tool helps you process multiple CSV bank statement files by:
1. Combining them into a single consolidated CSV file (preserving headers)
2. Optionally uploading the combined data to a Google Sheets spreadsheet

## Requirements

- Python 3.6+
- Required Python packages:
  - pandas
  - google-api-python-client
  - google-auth-oauthlib
  - google-auth

Install the required packages using the provided requirements.txt file:

```bash
pip install -r requirements.txt
```

## Configuration

### Google Sheets API Setup

To use the Google Sheets upload functionality, you need to:

1. Create a project in the [Google Cloud Console](https://console.cloud.google.com/)
2. Enable the Google Sheets API
3. Create OAuth 2.0 credentials
4. Download the credentials JSON file
5. Save the credentials file as `credentials.json` in the same directory as the script

## Usage

### Basic Usage

Combine all CSV files in a directory:

```bash
python import-statements.py /path/to/csv/files
```

### Specify Output File

```bash
python import-statements.py /path/to/csv/files -o output.csv
```

### Use a Specific File for Headers

```bash
python import-statements.py /path/to/csv/files -f specific_header_file.csv
```

### Upload to Google Sheets

```bash
python import-statements.py /path/to/csv/files -u -s SPREADSHEET_ID -t "Sheet Name"
```

## Command-line Arguments

- `directory`: Directory containing CSV files (required)
- `-o, --output`: Output file name (default: 'combined.csv')
- `-f, --header-file`: Specific file to use for the header (default: first CSV file found)
- `-u, --upload`: Flag to upload the combined CSV to Google Sheets
- `-s, --spreadsheet-id`: Google Spreadsheet ID for uploading (required with -u)
- `-t, --sheet-name`: Tab/Sheet name in the spreadsheet (required with -u)

## How It Works

1. **CSV Combination Process**:
   - The script scans the specified directory for CSV files
   - By default, it takes the header from the first CSV file found (alphabetically)
   - Alternatively, you can specify a specific file to use for the header
   - It combines all data rows (excluding headers) from all CSV files into a single output file
   - The resulting file is saved in the same directory as the input files
   
2. **Google Sheets Upload** (if enabled):
   - Authenticates with Google using OAuth 2.0
   - Finds the specified sheet within the spreadsheet
   - Clears any existing data in the target sheet
   - Uploads the combined CSV data to the specified Google Sheet
   - Updates the sheet with proper formatting

## Authentication Notes

- The first time you run the upload function, it will open a browser window prompting you to authorize the application
- Authentication tokens are stored in `token.json` in the script directory
- Subsequent runs will use the stored token without requiring re-authorization unless the token expires

## Error Handling

The script includes error handling for common issues:
- Missing CSV files in the specified directory
- Invalid header file specification
- Google Sheets API authentication failures
- Missing spreadsheet or sheet name
