import os
import glob
import argparse
import sys
import csv
import re
from dotenv import load_dotenv
from pathlib import Path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables from .env file
load_dotenv()

# Function to load account IBANs from environment variables
def load_account_ibans():
    """Load account names and IBANs from environment variables"""
    account_ibans = {}
    
    # Map from environment variable name to human-readable account name
    env_to_name_map = {
        "PERSONAL": "Personal",
        "BILLS_JOINT": "Bills (Joint)",
        "TRANSPORT": "Transport",
        "SAVINGS_NO_INT": "Savings (No Int.)",
        "SAVINGS": "Savings",
        "DONATIONS": "Donations",
        "BILLS": "Bills",
        "INVESTING_JOINT": "Investing (Joint)"
    }
    
    # Load IBANs from environment variables
    for env_name, display_name in env_to_name_map.items():
        iban = os.getenv(env_name)
        if iban:
            account_ibans[display_name] = iban
        else:
            print(f"Warning: IBAN for {display_name} not found in environment variables")
    
    return account_ibans

# Load account IBANs from environment variables
ACCOUNT_IBANS = load_account_ibans()

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "12EwoC2M-gtoPDShhrpeYK_9qdiPzGGtdcedXaJEINrw"
SAMPLE_RANGE_NAME = "test_sheet!A1:E10"

def create_client(credentials_file="credentials.json"):
    # If modifying these scopes, delete the file token.json.
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", scopes)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_file, scopes
            )
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    return creds

def print_sheet(spreadsheet_id=SAMPLE_SPREADSHEET_ID, range_name=SAMPLE_RANGE_NAME, creds=None):
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """

    creds = create_client()
    if not creds or not creds.valid:
        print("Invalid credentials. Please re-authenticate.")
        return

    try:
        service = build("sheets", "v4", credentials=creds)

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=spreadsheet_id, range=range_name)
            .execute()
        )
        values = result.get("values", [])

        if not values:
            print("No data found.")
            return

        print("Name, Major:")
        for row in values:
            # Print columns A and E, which correspond to indices 0 and 4.
            print(f"{row[0]}, {row[4]}")
    except HttpError as err:
        print(err)

def update_sheet(values, spreadsheet_id=SAMPLE_SPREADSHEET_ID, range_name=SAMPLE_RANGE_NAME, creds=None):
    """Updates values in a specified range of a spreadsheet.
    
    Args:
        values: The 2D array of values to write to the sheet
        spreadsheet_id: The ID of the spreadsheet to update
        range_name: The A1 notation of the range to update
        creds: Optional credentials object
        
    Returns:
        The update response from the API
    """
    if not creds:
        creds = create_client()
    if not creds or not creds.valid:
        print("Invalid credentials. Please re-authenticate.")
        return None

    try:
        service = build("sheets", "v4", credentials=creds)
        
        # Create the value range object with the provided values
        value_range_body = {
            'values': values
        }
        
        # Call the Sheets API to update the values
        request = service.spreadsheets().values().append(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption='USER_ENTERED',  # Interprets user input like Google Sheets would
            body=value_range_body
        )
        response = request.execute()
        
        print(f"Updated {response.get('updatedCells')} cells.")
        return response
        
    except HttpError as err:
        print(f"An error occurred: {err}")
        return None

def combine_csv_files(directory_path, output_file='combined.csv', header_file=None):
    try:
        # Change to the specified directory
        original_dir = os.getcwd()
        os.chdir(directory_path)
        
        # Get all CSV files in the specified directory
        csv_files = glob.glob('*.csv')
        
        # Exclude files named 'combined.csv' from the list
        csv_files = [f for f in csv_files if f.lower() != 'combined.csv']
        
        if not csv_files:
            print(f"No CSV files found in {directory_path} (excluding combined.csv)")
            return False
        
        # Determine which file to use for the header
        if header_file is None:
            header_file = csv_files[0]
        elif header_file not in csv_files:
            print(f"Header file {header_file} not found in {directory_path}")
            return False
        
        # Create a new combined output file
        with open(output_file, 'w') as outfile:
            # Write the header (first line) from the specified header file
            try:
                with open(header_file, 'r') as firstfile:
                    header = firstfile.readline().strip()
                    
                    # Add the new "Category" column to the header
                    if ',' in header:
                        # Comma-separated CSV
                        outfile.write(header + ",Category\n")
                    elif ';' in header:
                        # Semicolon-separated CSV
                        outfile.write(header + ";Category\n")
                    else:
                        # Default case - append with comma
                        outfile.write(header + ",Category\n")
                    
            except Exception as e:
                print(f"Error reading header from {header_file}: {str(e)}")
                return False
            
            # Append all lines except the header from each CSV file
            for filename in csv_files:
                try:
                    with open(filename, 'r') as infile:
                        next(infile)  # Skip the header
                        for line in infile:
                            line = line.rstrip('\n')
                            # Add an empty field for the Category column
                            if ',' in line:
                                outfile.write(line + ",\n")
                            elif ';' in line:
                                outfile.write(line + ";\n")
                            else:
                                outfile.write(line + ",\n")
                except Exception as e:
                    print(f"Error processing file {filename}: {str(e)}")

        filePath = os.path.join(directory_path, output_file)
        print(f"Combined CSV files created in '{filePath}' with Category column added")
        print(f"Combined {len(csv_files)} files (excluded 'combined.csv')")
        return filePath
    
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return False
    finally:
        # Return to the original directory
        os.chdir(original_dir)

def parse_csv_file(csv_file_path, delimiter=None, add_category_column=False):
    """
    Parses a CSV file into a 2D array suitable for Google Sheets.
    
    Args:
        csv_file_path: Path to the CSV file to parse
        delimiter: Optional specific delimiter to use (e.g., ',', ';')
        add_category_column: Whether to add an empty Category column if not present
        
    Returns:
        A 2D array representing the CSV data
    """
    data = []
    try:
        # Try to detect the delimiter if not specified
        if delimiter is None:
            # Read the first line to attempt to detect the delimiter
            with open(csv_file_path, 'r', newline='', encoding='utf-8') as f:
                first_line = f.readline().strip()
                # Check common delimiters
                if ';' in first_line:
                    delimiter = ';'
                    print(f"Detected semicolon (;) as delimiter")
                elif ',' in first_line:
                    delimiter = ','
                    print(f"Detected comma (,) as delimiter")
                elif '\t' in first_line:
                    delimiter = '\t'
                    print(f"Detected tab as delimiter")
                else:
                    # Default to comma if we can't detect
                    delimiter = ','
                    print(f"Could not detect delimiter, using comma (,) as default")
        
        # Parse the CSV file with the determined delimiter
        with open(csv_file_path, 'r', newline='', encoding='utf-8') as csvfile:
            csv_reader = csv.reader(csvfile, delimiter=delimiter)
            has_rows = False
            
            # Process the header first to check for Category column and find Account column index
            header_row = next(csv_reader, None)
            has_rows = header_row is not None
            
            if has_rows:
                # Find the Account column index if it exists
                account_column_index = -1
                for i, header in enumerate(header_row):
                    if header.lower() == 'account':
                        account_column_index = i
                        print(f"Found Account column at index {i}")
                        break
                
                # Check if Category column exists in the header
                has_category = 'Category' in header_row
                
                # Add the header to our data
                if add_category_column and not has_category:
                    header_row.append('Category')
                data.append(header_row)
                
                # Process remaining rows
                for row in csv_reader:
                    if row:  # Skip empty rows
                        # Replace IBAN with account name if the Account column is found
                        if account_column_index >= 0 and account_column_index < len(row):
                            iban = row[account_column_index].strip()
                            # Find matching account name for this IBAN
                            account_name = None
                            for name, account_iban in ACCOUNT_IBANS.items():
                                if account_iban == iban:
                                    account_name = name
                                    break
                            
                            # Replace IBAN with account name if found
                            if account_name:
                                row[account_column_index] = account_name
                                print(f"Replaced IBAN {iban} with account name: {account_name}")
                            
                        # Add empty category column if needed
                        if add_category_column and not has_category:
                            row.append('')
                        data.append(row)
            else:
                print(f"Warning: File {csv_file_path} is empty")
        
        # If we only have one column per row, we might have the wrong delimiter
        if data and all(len(row) == 1 for row in data):
            print(f"Warning: All rows have only one column. The delimiter might be incorrect.")
            
        return data
    except Exception as e:
        print(f"Error parsing CSV file {csv_file_path}: {str(e)}")
        return None

def upload_csv_to_sheet(csv_file_path, spreadsheet_id=SAMPLE_SPREADSHEET_ID, range_name=SAMPLE_RANGE_NAME, creds=None, delimiter=None, add_category_column=False):
    """
    Uploads data from a CSV file to a Google Sheet.
    
    Args:
        csv_file_path: Path to the CSV file to upload
        spreadsheet_id: The ID of the target spreadsheet
        range_name: The range in A1 notation to update
        creds: Optional credentials object
        delimiter: Optional delimiter for CSV parsing
        add_category_column: Whether to add a Category column if not present
        
    Returns:
        The update response from the API or None on error
    """
    data = parse_csv_file(csv_file_path, delimiter, add_category_column)
    if not data:
        return None
    
    # Print preview of the parsed data
    print(f"Preview of parsed data (first 2 rows):")
    for i, row in enumerate(data[:2]):
        print(f"Row {i}: {row}")
    print(f"Total rows: {len(data)}")
    
    # Upload the parsed data to the sheet
    return update_sheet(data, spreadsheet_id, range_name, creds)

def main():
    creds = create_client()

    # take as input the directory path containing CSV files
    parser = argparse.ArgumentParser(description="Combine CSV files and update Google Sheets.")
    parser.add_argument(
        "-d", "--directory-path", 
        required=False,
        type=str, 
        help="The directory containing CSV files to combine."
    )
    parser.add_argument(
        "-f", "--file-path",
        required=False,
        type=str,
        help="Path to a single CSV file to upload directly to Google Sheets."
    )
    parser.add_argument(
        "--delimiter",
        type=str,
        help="Delimiter used in the CSV file(s). Common values: ',' (comma), ';' (semicolon), '\\t' (tab)"
    )
    parser.add_argument(
        "--add-category",
        action="store_true",
        help="Add a Category column to the data if not present"
    )
    args = parser.parse_args()
    
    # Process delimiter
    delimiter = None
    if args.delimiter:
        if args.delimiter == '\\t':
            delimiter = '\t'
        else:
            delimiter = args.delimiter
    
    if args.file_path:
        # If a single file is specified, upload it directly
        if not os.path.exists(args.file_path):
            print(f"Error: File {args.file_path} not found.")
            return 1
            
        print(f"Uploading CSV file: {args.file_path}")
        response = upload_csv_to_sheet(args.file_path, creds=creds, delimiter=delimiter, 
                                       add_category_column=args.add_category)
        if response:
            print("CSV file uploaded successfully.")
            return 0
        else:
            print("Failed to upload CSV file.")
            return 1
            
    elif args.directory_path:
        # Directory processing code
        print(f"Combining CSV files from directory: {args.directory_path}")
        combined_file = combine_csv_files(args.directory_path)
        if not combined_file:
            print("Failed to combine CSV files.")
            return 1
        
        print(f"Combined file created: {combined_file}")
        
        # Upload the combined file to Google Sheets
        print(f"Uploading combined CSV file to Google Sheets...")
        # Category column is already added in combine_csv_files function
        response = upload_csv_to_sheet(combined_file, creds=creds, delimiter=delimiter)
        if response:
            print("Combined CSV file uploaded successfully.")
            return 0
        else:
            print("Failed to upload combined CSV file.")
            return 1
    else:
        parser.print_help()
        return 1
    

if __name__ == "__main__":
    sys.exit(main())
