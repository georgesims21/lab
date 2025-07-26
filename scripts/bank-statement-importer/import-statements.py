import os
import glob
import argparse
import sys
import csv
import re
from dotenv import load_dotenv
from pathlib import Path
from enum import Enum, auto
from typing import Dict, List, Optional, Set, Tuple, Union

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Load environment variables from .env file
load_dotenv()

# Define expense categories
class ExpenseCategory(Enum):
    HEALTH = "Health"
    TRAVEL = "Travel"
    SUBSCRIPTIONS = "Subscriptions"
    FOOD_DRINK = "Food & Drink"
    DONATIONS = "Donations"
    GROCERIES = "Groceries"
    SHOPPING = "Shopping"
    BOOKS = "Books"
    ENTERTAINMENT = "Entertainment"
    INVESTING = "Investing"
    PETER = "Peter"
    DEBT = "Debt"
    HOUSEHOLD = "Household"
    GIFTS = "Gifts"
    MISC = "Misc"
    INTEREST = "Interest"
    SAVING = "Saving"
    OV = "OV"
    PERSONAL_DEVELOPMENT = "Personal Development"
    SPORT = "Sport"
    INCOME = "Income"  # Added category for income
    
    @classmethod
    def from_string(cls, category_str: str) -> Optional['ExpenseCategory']:
        """Get enum value from string, case-insensitive"""
        for category in cls:
            if category.value.lower() == category_str.lower():
                return category
        return None
    
    @classmethod
    def get_all_categories(cls) -> List[str]:
        """Return list of all category names"""
        return [category.value for category in cls]

# Dictionary for category classification rules based on transaction descriptions
CATEGORY_RULES = {
    # Groceries
    "albert heijn": ExpenseCategory.GROCERIES,
    "ah": ExpenseCategory.GROCERIES,
    "jumbo": ExpenseCategory.GROCERIES,
    "lidl": ExpenseCategory.GROCERIES,
    "aldi": ExpenseCategory.GROCERIES,
    "supermarket": ExpenseCategory.GROCERIES,
    "tesco": ExpenseCategory.GROCERIES,
    "mercadona": ExpenseCategory.GROCERIES,
    "expsjoan98": ExpenseCategory.GROCERIES,
    "market roger de flor": ExpenseCategory.GROCERIES,
    "624 - ea aer": ExpenseCategory.GROCERIES,
    "paseo sant joan": ExpenseCategory.GROCERIES,
    "a taste of home": ExpenseCategory.GROCERIES,
    
    # Food & Drink
    "restaurant": ExpenseCategory.FOOD_DRINK,
    "cafe": ExpenseCategory.FOOD_DRINK,
    "bar": ExpenseCategory.FOOD_DRINK,
    "mcdonalds": ExpenseCategory.FOOD_DRINK,
    "starbucks": ExpenseCategory.FOOD_DRINK,
    "coffee": ExpenseCategory.FOOD_DRINK,
    "pret a manger": ExpenseCategory.FOOD_DRINK,
    "costa coffee": ExpenseCategory.FOOD_DRINK,
    "tio bigotes": ExpenseCategory.FOOD_DRINK,
    "honest greens": ExpenseCategory.FOOD_DRINK,
    "glovo": ExpenseCategory.FOOD_DRINK,
    "uber * eats": ExpenseCategory.FOOD_DRINK,
    "lounge": ExpenseCategory.FOOD_DRINK,
    "tandoori": ExpenseCategory.FOOD_DRINK,
    "gau lounge": ExpenseCategory.FOOD_DRINK,
    "bokaal": ExpenseCategory.FOOD_DRINK,
    "harvest coffee": ExpenseCategory.FOOD_DRINK,
    "het kroket": ExpenseCategory.FOOD_DRINK,
    "yellowriverlanz": ExpenseCategory.FOOD_DRINK,
    "fenix food": ExpenseCategory.FOOD_DRINK,
    "bar kauffmann": ExpenseCategory.FOOD_DRINK,
    "five ways coffee": ExpenseCategory.FOOD_DRINK,
    "artisanal cuis": ExpenseCategory.FOOD_DRINK,
    "zuylen": ExpenseCategory.FOOD_DRINK,
    "nypdamsterdamarena": ExpenseCategory.FOOD_DRINK,
    "boulevard bv": ExpenseCategory.FOOD_DRINK,
    
    # Travel
    "ns ": ExpenseCategory.TRAVEL,
    "ns reizigers": ExpenseCategory.TRAVEL,
    "train": ExpenseCategory.TRAVEL,
    "flight": ExpenseCategory.TRAVEL,
    "airline": ExpenseCategory.TRAVEL,
    "hotel": ExpenseCategory.TRAVEL,
    "airbnb": ExpenseCategory.TRAVEL,
    "national express": ExpenseCategory.TRAVEL,
    "renfe": ExpenseCategory.TRAVEL,
    "metro barcelona": ExpenseCategory.TRAVEL,
    "uber * pending": ExpenseCategory.TRAVEL,
    "monbus": ExpenseCategory.TRAVEL,
    "e521": ExpenseCategory.TRAVEL, # NS e-Tickets
    "ov-chipkaart": ExpenseCategory.TRAVEL,
    
    # Subscriptions
    "netflix": ExpenseCategory.SUBSCRIPTIONS,
    "spotify": ExpenseCategory.SUBSCRIPTIONS,
    "duo": ExpenseCategory.SUBSCRIPTIONS, # Student loan
    "google one": ExpenseCategory.SUBSCRIPTIONS,
    "amazon prime": ExpenseCategory.SUBSCRIPTIONS,
    "obsidian.md": ExpenseCategory.SUBSCRIPTIONS,
    "odido": ExpenseCategory.SUBSCRIPTIONS,
    "monthly acc. costs": ExpenseCategory.SUBSCRIPTIONS,
    "zilveren kruis": ExpenseCategory.SUBSCRIPTIONS, # Health insurance
    "google*google play": ExpenseCategory.SUBSCRIPTIONS,
    "altafit": ExpenseCategory.SUBSCRIPTIONS, # Gym subscription
    
    # Household
    "rent": ExpenseCategory.HOUSEHOLD,
    "bills": ExpenseCategory.HOUSEHOLD,
    "insurance": ExpenseCategory.HOUSEHOLD,
    "utilities": ExpenseCategory.HOUSEHOLD,
    "electricity": ExpenseCategory.HOUSEHOLD,
    "water": ExpenseCategory.HOUSEHOLD,
    "gas": ExpenseCategory.HOUSEHOLD,
    "wight": ExpenseCategory.HOUSEHOLD, # Landlord
    
    # Entertainment
    "ziggo dome": ExpenseCategory.ENTERTAINMENT,
    "funk fest": ExpenseCategory.ENTERTAINMENT,
    "wildlife par": ExpenseCategory.ENTERTAINMENT,
    "limp biz": ExpenseCategory.ENTERTAINMENT,
    "meaker": ExpenseCategory.ENTERTAINMENT,
    "ticketmaster": ExpenseCategory.ENTERTAINMENT,
    "pierre via tikkie": ExpenseCategory.ENTERTAINMENT, # Concert ticket
    
    # Shopping
    "uniqlo": ExpenseCategory.SHOPPING,
    "amazon": ExpenseCategory.SHOPPING,
    "towel": ExpenseCategory.SHOPPING,
    
    # Health
    "doctor": ExpenseCategory.HEALTH,
    "hospital": ExpenseCategory.HEALTH,
    "pharmacy": ExpenseCategory.HEALTH,
    "medical": ExpenseCategory.HEALTH,
    "zilveren kruis": ExpenseCategory.HEALTH,
    
    # Donations
    "donation": ExpenseCategory.DONATIONS,
    "charity": ExpenseCategory.DONATIONS,
    "gofundme": ExpenseCategory.DONATIONS,
    "gfm*gofundme": ExpenseCategory.DONATIONS,
    
    # Books
    "kindle": ExpenseCategory.BOOKS,
    "book": ExpenseCategory.BOOKS,
    
    # Investing
    "flatexdegiro": ExpenseCategory.INVESTING,
    "invest": ExpenseCategory.INVESTING,
    
    # Sport
    "playtomic": ExpenseCategory.SPORT,
    "gym": ExpenseCategory.SPORT,
    "fitness": ExpenseCategory.SPORT,
    
    # Debt
    "transferwise": ExpenseCategory.DEBT,
    
    # Saving
    "trade republic": ExpenseCategory.SAVING,
    
    # Interest
    "bunq payday": ExpenseCategory.INTEREST,
    "interest": ExpenseCategory.INTEREST,
    
    # Gifts
    "sagrada familia": ExpenseCategory.GIFTS,
    "gift": ExpenseCategory.GIFTS,
    
    # OV (Public Transport)
    "ov": ExpenseCategory.OV,
    "gvb": ExpenseCategory.OV,
    "ns automaat": ExpenseCategory.OV, # Train station purchases
    
    # Misc (for transactions that don't fit other categories)
    "afecte": ExpenseCategory.MISC,
    "rihab ibne": ExpenseCategory.MISC,
    "bhoekhan": ExpenseCategory.MISC,
    "valencia": ExpenseCategory.MISC,
    
    # Income
    "bynder": ExpenseCategory.INCOME,  # We'll check amount separately
    "salary": ExpenseCategory.INCOME,
    "payroll": ExpenseCategory.INCOME,
    "loon": ExpenseCategory.INCOME,   # Dutch for salary
    "inkomen": ExpenseCategory.INCOME, # Dutch for income
    "wage": ExpenseCategory.INCOME,
}

# Function to improve category matching with more sophisticated logic
def get_category_for_description(description: str) -> Optional[ExpenseCategory]:
    """
    Determine expense category based on transaction description
    with more sophisticated matching logic
    
    Args:
        description: Transaction description text
        
    Returns:
        Matching ExpenseCategory or None if no match found
    """
    if not description:
        return None
        
    # Convert to lowercase for case-insensitive matching
    desc_lower = description.lower()
    
    # Special case handling for common patterns
    if "albert heijn" in desc_lower or "jumbo" in desc_lower or "aldi" in desc_lower:
        return ExpenseCategory.GROCERIES
    
    if "rent" in desc_lower and ("bills" in desc_lower or "wight" in desc_lower):
        return ExpenseCategory.HOUSEHOLD
    
    if ("zilveren kruis" in desc_lower) and "premie" in desc_lower:
        return ExpenseCategory.SUBSCRIPTIONS
        
    if "duo" in desc_lower and "studieschuld" in desc_lower:
        return ExpenseCategory.SUBSCRIPTIONS
    
    # Check for NS train tickets
    if "e52" in desc_lower and "ns" in desc_lower and "tickets" in desc_lower:
        return ExpenseCategory.TRAVEL
    
    # Check for food delivery services
    if "uber" in desc_lower and "eats" in desc_lower:
        return ExpenseCategory.FOOD_DRINK
    
    # Check for keywords in the description
    for keyword, category in CATEGORY_RULES.items():
        if keyword in desc_lower:
            return category
            
    return None

# Helper function to extract amount value from string
def extract_amount_value(amount_str: str) -> float:
    """
    Extract numerical value from amount string, handling various formats
    
    Args:
        amount_str: String representation of amount (e.g., '€3,000.00', '-€24,06')
        
    Returns:
        Float value of the amount
    """
    if not amount_str or not isinstance(amount_str, str):
        return 0.0
    
    # Remove currency symbols and spaces
    cleaned = amount_str.replace('€', '').replace('$', '').replace(' ', '')
    
    try:
        # Handle European format (comma as decimal separator)
        if ',' in cleaned and '.' in cleaned:
            # Format with both comma and dot - assume comma is thousands separator
            cleaned = cleaned.replace(',', '')
            return float(cleaned)
        elif ',' in cleaned:
            # Comma as decimal separator
            return float(cleaned.replace(',', '.'))
        else:
            # Standard format or just a dot as decimal separator
            return float(cleaned)
    except ValueError:
        print(f"Could not parse amount: '{amount_str}'")
        return 0.0

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

def format_european_number(amount_str):
    """
    Format a number string from English format to European format
    - Changes decimal point from '.' to ','
    - Changes thousand separator from ',' to '.'
    
    Examples:
        "123.45" -> "123,45"
        "1,234.56" -> "1.234,56"
    """
    if not amount_str or not isinstance(amount_str, str):
        return amount_str
        
    amount_str = amount_str.strip()
    
    # Handle simple case first (no thousand separators)
    if '.' in amount_str and ',' not in amount_str:
        return amount_str.replace('.', ',')
        
    # Handle numbers with thousand separators
    if ',' in amount_str:
        # First, replace commas with temporary placeholder
        temp_str = amount_str.replace(',', '|')
        # Then replace decimal point with comma
        if '.' in temp_str:
            temp_str = temp_str.replace('.', ',')
        # Finally replace the placeholder with dots
        final_str = temp_str.replace('|', '.')
        return final_str
    
    # Return original if no formatting needed
    return amount_str

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
        
        # Expand path if needed (for ~ home directory)
        expanded_path = os.path.expanduser(csv_file_path)
        
        # Parse the CSV file with the determined delimiter
        with open(expanded_path, 'r', newline='', encoding='utf-8') as csvfile:
            csv_reader = csv.reader(csvfile, delimiter=delimiter)
            has_rows = False
            
            # Process the header first to check for Category column and find column indices
            header_row = next(csv_reader, None)
            has_rows = header_row is not None
            
            if has_rows:
                print(f"CSV Headers: {header_row}")
                
                # Find various column indices
                account_column_index = -1
                amount_column_index = -1
                description_column_index = -1
                category_column_index = -1
                counterparty_column_index = -1  # For counterparty name
                
                for i, header in enumerate(header_row):
                    header_lower = header.lower() if header else ""
                    
                    if header_lower == 'account':
                        account_column_index = i
                        print(f"Found Account column at index {i}")
                    elif header_lower in ('amount', 'bedrag'):  # 'bedrag' is Dutch for 'amount'
                        amount_column_index = i
                        print(f"Found Amount column at index {i}")
                    elif header_lower in ('description', 'desc', 'details', 'transaction', 'omschrijving'):
                        description_column_index = i
                        print(f"Found Description column at index {i}")
                    elif header_lower == 'category':
                        category_column_index = i
                        print(f"Found existing Category column at index {i}")
                    elif header_lower in ('counterparty', 'tegenrekening', 'naam', 'name', 'counterparty name'):
                        counterparty_column_index = i
                        print(f"Found Counterparty column at index {i}")
                
                # Check if Category column exists in the header
                has_category = category_column_index >= 0
                
                # Add the Category column to the header if needed and requested
                if add_category_column and not has_category:
                    header_row.append('Category')
                    category_column_index = len(header_row) - 1
                    print(f"Added Category column at index {category_column_index}")
                    
                data.append(header_row)
                
                # Set a counter for how many rows were categorized
                categorized_rows = 0
                total_rows = 0
                
                # Process remaining rows
                for row in csv_reader:
                    if not row:  # Skip empty rows
                        continue
                        
                    total_rows += 1
                    
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
                    
                    # Format the Amount column to use European number format
                    if amount_column_index >= 0 and amount_column_index < len(row):
                        amount_str = row[amount_column_index].strip()
                        if amount_str:
                            try:
                                # Use our formatting function
                                original = amount_str
                                formatted = format_european_number(amount_str)
                                row[amount_column_index] = formatted
                                
                                if original != formatted:
                                    print(f"Reformatted amount: '{original}' -> '{formatted}'")
                            except Exception as e:
                                print(f"Error formatting amount '{amount_str}': {str(e)}")
                    
                    # Handle categorization
                    if category_column_index >= 0:
                        # Ensure we have enough columns for the category
                        while len(row) <= category_column_index:
                            row.append('')
                            
                        # Only suggest category if the field is currently empty or we're explicitly asked to categorize
                        if not row[category_column_index].strip() or add_category_column:
                            # Collect all text that might help determine the category
                            category_text = ""
                            
                            # Add description if available
                            if description_column_index >= 0 and description_column_index < len(row):
                                category_text += " " + row[description_column_index]
                                
                            # Add counterparty if available
                            if counterparty_column_index >= 0 and counterparty_column_index < len(row):
                                category_text += " " + row[counterparty_column_index]
                            
                            # Get amount if available for special categorization rules
                            amount_value = 0.0
                            if amount_column_index >= 0 and amount_column_index < len(row):
                                amount_str = row[amount_column_index].strip()
                                amount_value = extract_amount_value(amount_str)
                            
                            # Special case for income from Bynder
                            if "bynder" in category_text.lower() and amount_value >= 3000:
                                row[category_column_index] = ExpenseCategory.INCOME.value
                                categorized_rows += 1
                                print(f"Auto-categorized as 'Income' based on Bynder payment of {amount_value} euros")
                                continue  # Skip further categorization attempts
                                
                            # If we have some text to work with, try to determine category
                            if category_text.strip():
                                category = get_category_for_description(category_text.strip())
                                if category:
                                    row[category_column_index] = category.value
                                    categorized_rows += 1
                                    print(f"Auto-categorized as '{category.value}' based on text: '{category_text.strip()}'")
                    
                    # Add empty category column if needed
                    elif add_category_column:
                        row.append('')
                        
                    data.append(row)
                
                # Show categorization summary
                if add_category_column and total_rows > 0:
                    success_rate = (categorized_rows / total_rows) * 100
                    print(f"Categorization complete: {categorized_rows} out of {total_rows} rows categorized ({success_rate:.1f}%)")
            else:
                print(f"Warning: File {csv_file_path} is empty")
        
        # If we only have one column per row, we might have the wrong delimiter
        if data and len(data) > 1 and all(len(row) == 1 for row in data):
            print(f"Warning: All rows have only one column. The delimiter might be incorrect.")
            
        return data
    except Exception as e:
        print(f"Error parsing CSV file {csv_file_path}: {str(e)}")
        import traceback
        traceback.print_exc()
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
    parser.add_argument(
        "--auto-categorize",
        action="store_true",
        help="Automatically assign categories based on transaction descriptions"
    )
    parser.add_argument(
        "--force-categorize",
        action="store_true",
        help="Force recategorization even if Category column has values"
    )
    args = parser.parse_args()
    
    # Process delimiter
    delimiter = None
    if args.delimiter:
        if args.delimiter == '\\t':
            delimiter = '\t'
        else:
            delimiter = args.delimiter
    
    # Determine if we should add and auto-fill categories
    add_category = args.add_category or args.auto_categorize or args.force_categorize
    
    if args.file_path:
        # If a single file is specified, upload it directly
        expanded_path = os.path.expanduser(args.file_path)
        if not os.path.exists(expanded_path):
            print(f"Error: File {expanded_path} not found.")
            return 1
            
        print(f"Uploading CSV file: {expanded_path}")
        response = upload_csv_to_sheet(expanded_path, creds=creds, delimiter=delimiter, 
                                       add_category_column=add_category)
        if response:
            print("CSV file uploaded successfully.")
            return 0
        else:
            print("Failed to upload CSV file.")
            return 1
            
    elif args.directory_path:
        # Directory processing code
        expanded_dir = os.path.expanduser(args.directory_path)
        print(f"Combining CSV files from directory: {expanded_dir}")
        combined_file = combine_csv_files(expanded_dir)
        if not combined_file:
            print("Failed to combine CSV files.")
            return 1
        
        print(f"Combined file created: {combined_file}")
        
        # Upload the combined file to Google Sheets
        print(f"Uploading combined CSV file to Google Sheets...")
        # Category column is already added in combine_csv_files function
        response = upload_csv_to_sheet(combined_file, creds=creds, delimiter=delimiter, 
                                      add_category_column=add_category)
        if response:
            print("Combined CSV file uploaded successfully.")
            return 0
        else:
            print("Failed to upload combined CSV file.")
            return 1
    else:
        # Print list of available categories
        print("Available expense categories:")
        for category in ExpenseCategory.get_all_categories():
            print(f"  - {category}")
        
        parser.print_help()
        return 1
    

if __name__ == "__main__":
    sys.exit(main())
