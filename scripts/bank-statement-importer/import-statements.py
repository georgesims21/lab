import os
import glob
import argparse
import sys
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

def combine_csv_files(directory_path, output_file='combined.csv', header_file=None):
    try:
        # Change to the specified directory
        original_dir = os.getcwd()
        os.chdir(directory_path)
        
        # Get all CSV files in the specified directory
        csv_files = glob.glob('*.csv')
        
        if not csv_files:
            print(f"No CSV files found in {directory_path}")
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
                    header = firstfile.readline()
                    outfile.write(header)
            except Exception as e:
                print(f"Error reading header from {header_file}: {str(e)}")
                return False
            
            # Append all lines except the header from each CSV file
            for filename in csv_files:
                try:
                    with open(filename, 'r') as infile:
                        next(infile)  # Skip the header
                        for line in infile:
                            outfile.write(line)
                except Exception as e:
                    print(f"Error processing file {filename}: {str(e)}")
        
        print(f"Combined CSV files created in '{os.path.join(directory_path, output_file)}'")
        return True
    
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        return False
    finally:
        # Return to the original directory
        os.chdir(original_dir)

def upload_to_google_sheet(csv_file_path, spreadsheet_id, sheet_name):
    """
    Uploads CSV data to an existing tab in a Google Sheet
    """
    try:
        # Define the OAuth scopes
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        
        # Authentication process
        creds = None
        token_file = 'token.json'
        
        if os.path.exists(token_file):
            creds = Credentials.from_authorized_user_file(token_file, SCOPES)
            
        # If credentials don't exist or are invalid, prompt the user to log in
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            
            # Save the credentials for future use
            with open(token_file, 'w') as token:
                token.write(creds.to_json())
        
        # Create the Sheets API service
        service = build('sheets', 'v4', credentials=creds)
        
        # Read the CSV file
        df = pd.read_csv(csv_file_path)
        
        # Prepare the data for upload
        values = [df.columns.tolist()] + df.values.tolist()
        
        # Find the sheet ID from the sheet name
        sheet_metadata = service.spreadsheets().get(spreadsheetId=spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        sheet_id = None
        
        for sheet in sheets:
            if sheet['properties']['title'] == sheet_name:
                sheet_id = sheet['properties']['sheetId']
                break
        
        if sheet_id is None:
            print(f"Error: Sheet '{sheet_name}' not found in the spreadsheet")
            return False
        
        # Clear existing data in the sheet
        clear_request = service.spreadsheets().values().clear(
            spreadsheetId=spreadsheet_id,
            range=sheet_name
        )
        clear_request.execute()
        
        # Update the sheet with the CSV data
        body = {'values': values}
        result = service.spreadsheets().values().update(
            spreadsheetId=spreadsheet_id,
            range=f"{sheet_name}!A1",
            valueInputOption='USER_ENTERED',
            body=body
        ).execute()
        
        print(f"Data successfully uploaded to Google Sheet tab '{sheet_name}'")
        print(f"{result.get('updatedCells')} cells updated.")
        return True
        
    except Exception as e:
        print(f"Error uploading to Google Sheet: {str(e)}")
        return False

def main():
    parser = argparse.ArgumentParser(description='Combine multiple CSV files into one and optionally upload to Google Sheets.')
    parser.add_argument('directory', help='Directory containing CSV files')
    parser.add_argument('-o', '--output', default='combined.csv', help='Output file name')
    parser.add_argument('-f', '--header-file', help='File to use for the header')
    parser.add_argument('-u', '--upload', action='store_true', help='Upload the combined CSV to Google Sheets')
    parser.add_argument('-s', '--spreadsheet-id', help='Google Spreadsheet ID to upload data to')
    parser.add_argument('-t', '--sheet-name', help='Tab/Sheet name in the spreadsheet')
    
    args = parser.parse_args()
    
    if not os.path.isdir(args.directory):
        print(f"Error: {args.directory} is not a valid directory")
        return 1
    
    success = combine_csv_files(args.directory, args.output, args.header_file)
    if not success:
        return 1
    
    # Upload to Google Sheets if requested
    if args.upload:
        if not args.spreadsheet_id or not args.sheet_name:
            print("Error: Spreadsheet ID and Sheet name are required for uploading to Google Sheets")
            return 1
        
        output_path = os.path.join(args.directory, args.output)
        upload_success = upload_to_google_sheet(output_path, args.spreadsheet_id, args.sheet_name)
        return 0 if upload_success else 1
    
    return 0

if __name__ == "__main__":
    sys.exit(main())
