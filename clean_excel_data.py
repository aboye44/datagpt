#!/usr/bin/env python3
"""
Excel Data Cleaning Script for BCC Bulk Mailer

This script reads Excel files with messy mailing data, parses combined fields,
removes duplicates, validates addresses, and outputs BCC Bulk Mailer-ready CSV files.

BCC Bulk Mailer Business/Professional Import Format:
- FullName, Company, Address1, Address2, City, State, ZIP, ZIP4, Email, Notes
- UTF-8 or ANSI encoded CSV
- Clean headers for automatic field mapping
- USPS-compliant address formatting
"""

import pandas as pd
import argparse
import sys
import re
from pathlib import Path
from typing import Dict, Tuple, Optional


# USPS State Abbreviations
STATE_ABBREVIATIONS = {
    'alabama': 'AL', 'alaska': 'AK', 'arizona': 'AZ', 'arkansas': 'AR', 'california': 'CA',
    'colorado': 'CO', 'connecticut': 'CT', 'delaware': 'DE', 'florida': 'FL', 'georgia': 'GA',
    'hawaii': 'HI', 'idaho': 'ID', 'illinois': 'IL', 'indiana': 'IN', 'iowa': 'IA',
    'kansas': 'KS', 'kentucky': 'KY', 'louisiana': 'LA', 'maine': 'ME', 'maryland': 'MD',
    'massachusetts': 'MA', 'michigan': 'MI', 'minnesota': 'MN', 'mississippi': 'MS',
    'missouri': 'MO', 'montana': 'MT', 'nebraska': 'NE', 'nevada': 'NV', 'new hampshire': 'NH',
    'new jersey': 'NJ', 'new mexico': 'NM', 'new york': 'NY', 'north carolina': 'NC',
    'north dakota': 'ND', 'ohio': 'OH', 'oklahoma': 'OK', 'oregon': 'OR', 'pennsylvania': 'PA',
    'rhode island': 'RI', 'south carolina': 'SC', 'south dakota': 'SD', 'tennessee': 'TN',
    'texas': 'TX', 'utah': 'UT', 'vermont': 'VT', 'virginia': 'VA', 'washington': 'WA',
    'west virginia': 'WV', 'wisconsin': 'WI', 'wyoming': 'WY', 'district of columbia': 'DC',
    'puerto rico': 'PR', 'guam': 'GU', 'virgin islands': 'VI', 'american samoa': 'AS',
    'northern mariana islands': 'MP'
}

# Column name mapping for standardization
COLUMN_MAPPING = {
    # Name variations
    'first_name': 'FirstName', 'firstname': 'FirstName', 'first': 'FirstName',
    'fname': 'FirstName', 'given_name': 'FirstName', 'givenname': 'FirstName',
    'last_name': 'LastName', 'lastname': 'LastName', 'last': 'LastName',
    'lname': 'LastName', 'surname': 'LastName', 'family_name': 'LastName',
    'familyname': 'LastName', 'name': 'FullName', 'full_name': 'FullName',
    'fullname': 'FullName', 'contact': 'FullName', 'contact_name': 'FullName',

    # Company variations
    'company': 'Company', 'business': 'Company', 'organization': 'Company',
    'org': 'Company', 'company_name': 'Company', 'business_name': 'Company',

    # Address variations
    'address': 'Address1', 'address1': 'Address1', 'address_1': 'Address1',
    'street': 'Address1', 'street_address': 'Address1', 'streetaddress': 'Address1',
    'addr': 'Address1', 'addr1': 'Address1', 'address_line_1': 'Address1',
    'address2': 'Address2', 'address_2': 'Address2', 'apt': 'Address2',
    'suite': 'Address2', 'unit': 'Address2', 'address_line_2': 'Address2',

    # City variations
    'city': 'City', 'town': 'City', 'municipality': 'City',

    # State variations
    'state': 'State', 'province': 'State', 'region': 'State', 'st': 'State',

    # ZIP variations
    'zip': 'ZIP', 'zipcode': 'ZIP', 'zip_code': 'ZIP', 'postal_code': 'ZIP',
    'postalcode': 'ZIP', 'postcode': 'ZIP', 'zip5': 'ZIP',
    'zip4': 'ZIP4', 'zip_4': 'ZIP4', 'plus4': 'ZIP4',

    # Email variations
    'email': 'Email', 'email_address': 'Email', 'e-mail': 'Email', 'mail': 'Email',

    # Notes variations
    'notes': 'Notes', 'note': 'Notes', 'comments': 'Notes', 'comment': 'Notes',
    'remarks': 'Notes', 'memo': 'Notes'
}


def parse_combined_address(address_string: str) -> Dict[str, str]:
    """
    Parse a combined address string like "John Doe, 123 Main St, Boston, MA 02101"
    into separate components.

    Args:
        address_string: Combined address string

    Returns:
        Dictionary with parsed components
    """
    result = {
        'FullName': '', 'Address1': '', 'Address2': '',
        'City': '', 'State': '', 'ZIP': '', 'ZIP4': ''
    }

    if not address_string or pd.isna(address_string):
        return result

    address_string = str(address_string).strip()

    # Try to extract ZIP code first (most reliable anchor)
    zip_pattern = r'\b(\d{5})(?:-(\d{4}))?\b'
    zip_match = re.search(zip_pattern, address_string)
    if zip_match:
        result['ZIP'] = zip_match.group(1)
        if zip_match.group(2):
            result['ZIP4'] = zip_match.group(2)
        # Remove ZIP from string for further processing
        address_string = address_string[:zip_match.start()] + address_string[zip_match.end():]

    # Try to extract state (2-letter code before ZIP)
    state_pattern = r'\b([A-Z]{2})\b'
    state_matches = re.findall(state_pattern, address_string)
    if state_matches:
        # Take the last 2-letter code (most likely the state)
        potential_state = state_matches[-1]
        if potential_state in STATE_ABBREVIATIONS.values():
            result['State'] = potential_state
            # Remove state from string
            address_string = re.sub(r'\b' + potential_state + r'\b', '', address_string, count=1)

    # Split by commas or newlines
    parts = [p.strip() for p in re.split(r'[,\n]', address_string) if p.strip()]

    if not parts:
        return result

    # Heuristic parsing
    if len(parts) >= 1:
        # First part is likely name (if it doesn't look like an address)
        first_part = parts[0]
        if not re.search(r'\d', first_part):  # No numbers = probably a name
            result['FullName'] = first_part
            parts = parts[1:]

    if len(parts) >= 1:
        # Next part with numbers is likely street address
        for i, part in enumerate(parts):
            if re.search(r'\d', part):
                result['Address1'] = part
                parts.pop(i)
                break

    if len(parts) >= 1:
        # Remaining parts are likely city and possibly address2
        # City is usually the last or second-to-last part
        result['City'] = parts[-1] if parts else ''
        if len(parts) > 1:
            # Check if second-to-last looks like address2 (has unit/apt/suite)
            if re.search(r'\b(apt|suite|unit|ste|#)\b', parts[-2], re.IGNORECASE):
                result['Address2'] = parts[-2]
            else:
                # Otherwise it might be part of the city or an additional address line
                result['City'] = parts[-2] + ' ' + result['City']

    return result


def parse_name(name_string: str) -> Tuple[str, str]:
    """
    Parse a name string into FirstName and LastName.
    Handles "Last, First" and "First Last" formats.

    Args:
        name_string: Name to parse

    Returns:
        Tuple of (FirstName, LastName)
    """
    if not name_string or pd.isna(name_string):
        return ('', '')

    name_string = str(name_string).strip()

    # Check for "Last, First" format
    if ',' in name_string:
        parts = [p.strip() for p in name_string.split(',', 1)]
        if len(parts) == 2:
            return (parts[1], parts[0])  # (First, Last)

    # Split on whitespace
    parts = name_string.split()

    if len(parts) == 0:
        return ('', '')
    elif len(parts) == 1:
        return ('', parts[0])  # Assume single name is last name
    elif len(parts) == 2:
        return (parts[0], parts[1])
    else:
        # 3+ parts: First name is first part, last name is last part
        return (parts[0], parts[-1])


def normalize_state(state_value: str) -> str:
    """
    Normalize state to USPS 2-letter abbreviation.

    Args:
        state_value: State name or abbreviation

    Returns:
        2-letter state abbreviation or original value
    """
    if not state_value or pd.isna(state_value):
        return ''

    state_value = str(state_value).strip()

    # Already a 2-letter code
    if len(state_value) == 2:
        state_upper = state_value.upper()
        if state_upper in STATE_ABBREVIATIONS.values():
            return state_upper

    # Try to match full state name
    state_lower = state_value.lower()
    if state_lower in STATE_ABBREVIATIONS:
        return STATE_ABBREVIATIONS[state_lower]

    # Try to match with common variations
    state_lower = re.sub(r'[^\w\s]', '', state_lower)  # Remove punctuation
    if state_lower in STATE_ABBREVIATIONS:
        return STATE_ABBREVIATIONS[state_lower]

    return state_value  # Return original if no match


def validate_zip(zip_value: str) -> Tuple[str, str]:
    """
    Validate and format ZIP code.

    Args:
        zip_value: ZIP code value

    Returns:
        Tuple of (ZIP5, ZIP4)
    """
    if not zip_value or pd.isna(zip_value):
        return ('', '')

    zip_string = str(zip_value).strip()

    # Extract 5-digit ZIP and optional +4
    zip_pattern = r'(\d{5})(?:-?(\d{4}))?'
    match = re.search(zip_pattern, zip_string)

    if match:
        zip5 = match.group(1)
        zip4 = match.group(2) if match.group(2) else ''
        return (zip5, zip4)

    return ('', '')


def title_case_address(address: str) -> str:
    """
    Properly capitalize address using title case with exceptions.

    Args:
        address: Address string

    Returns:
        Properly capitalized address
    """
    if not address or pd.isna(address):
        return ''

    address = str(address).strip()

    # Common abbreviations that should stay uppercase
    upper_words = {'PO', 'NW', 'NE', 'SW', 'SE', 'N', 'S', 'E', 'W', 'US', 'USA'}

    words = address.split()
    result = []

    for word in words:
        if word.upper() in upper_words:
            result.append(word.upper())
        else:
            result.append(word.title())

    return ' '.join(result)


def standardize_column_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    Standardize column names for BCC Bulk Mailer.

    Args:
        df: pandas DataFrame

    Returns:
        DataFrame with standardized column names
    """
    df = df.copy()

    # Normalize column names (lowercase, strip whitespace, replace spaces with underscores)
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('-', '_')

    # Apply column mapping
    df.rename(columns=COLUMN_MAPPING, inplace=True)

    return df


def parse_messy_data(df: pd.DataFrame) -> pd.DataFrame:
    """
    Parse messy data where multiple fields might be combined in single cells.

    Args:
        df: pandas DataFrame

    Returns:
        DataFrame with parsed fields
    """
    df = df.copy()

    # Check if we have very few columns - might indicate combined data
    if len(df.columns) <= 2:
        print("Detected possible combined data format - attempting to parse...")

        # Try to find the column with the most data
        main_col = None
        for col in df.columns:
            if df[col].notna().sum() > 0:
                main_col = col
                break

        if main_col:
            parsed_rows = []
            for idx, row in df.iterrows():
                parsed = parse_combined_address(row[main_col])
                parsed_rows.append(parsed)

            # Create new dataframe from parsed data
            df = pd.DataFrame(parsed_rows)
            print(f"Parsed {len(df)} rows from combined format")

    # Split FullName into FirstName/LastName if needed
    if 'FullName' in df.columns and 'FirstName' not in df.columns:
        print("Splitting FullName into FirstName and LastName...")
        names = df['FullName'].apply(parse_name)
        df['FirstName'] = names.apply(lambda x: x[0])
        df['LastName'] = names.apply(lambda x: x[1])

    # Combine FirstName/LastName into FullName if FullName is missing
    if 'FirstName' in df.columns and 'LastName' in df.columns and 'FullName' not in df.columns:
        print("Creating FullName from FirstName and LastName...")
        df['FullName'] = (df['FirstName'].fillna('').astype(str).str.strip() + ' ' +
                         df['LastName'].fillna('').astype(str).str.strip()).str.strip()

    return df


def clean_and_validate(df: pd.DataFrame) -> pd.DataFrame:
    """
    Clean and validate address data for BCC Bulk Mailer.

    Args:
        df: pandas DataFrame

    Returns:
        Cleaned DataFrame with validation flags
    """
    df = df.copy()

    # Clean whitespace from all string columns
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].astype(str).str.strip()
        # Replace 'nan' string with empty string
        df[col] = df[col].replace('nan', '')

    # Normalize State
    if 'State' in df.columns:
        print("Normalizing state abbreviations...")
        df['State'] = df['State'].apply(normalize_state)

    # Validate and format ZIP codes
    if 'ZIP' in df.columns:
        print("Validating ZIP codes...")
        zip_data = df['ZIP'].apply(validate_zip)
        df['ZIP'] = zip_data.apply(lambda x: x[0])
        if 'ZIP4' not in df.columns or df['ZIP4'].fillna('').eq('').all():
            df['ZIP4'] = zip_data.apply(lambda x: x[1])

    # Title case addresses
    if 'Address1' in df.columns:
        print("Formatting addresses...")
        df['Address1'] = df['Address1'].apply(title_case_address)

    if 'Address2' in df.columns:
        df['Address2'] = df['Address2'].apply(title_case_address)

    if 'City' in df.columns:
        df['City'] = df['City'].apply(title_case_address)

    # Flag incomplete records
    required_fields = ['Address1', 'City', 'State', 'ZIP']
    incomplete_mask = False
    for field in required_fields:
        if field in df.columns:
            incomplete_mask |= df[field].fillna('').eq('')

    if incomplete_mask.any():
        incomplete_count = incomplete_mask.sum()
        print(f"Warning: {incomplete_count} records have incomplete address data")

        # Add to Notes if column exists, otherwise create it
        if 'Notes' not in df.columns:
            df['Notes'] = ''

        df.loc[incomplete_mask, 'Notes'] = df.loc[incomplete_mask, 'Notes'].apply(
            lambda x: (x + '; ' if x else '') + 'INCOMPLETE_ADDRESS'
        )

    return df


def remove_duplicates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Remove duplicate rows based on matching Address1 + Name.

    Args:
        df: pandas DataFrame

    Returns:
        DataFrame with duplicates removed
    """
    df = df.copy()

    # Determine deduplication columns based on what's available
    dedupe_columns = []

    if 'Address1' in df.columns:
        dedupe_columns.append('Address1')

    if 'FullName' in df.columns:
        dedupe_columns.append('FullName')
    elif 'FirstName' in df.columns and 'LastName' in df.columns:
        dedupe_columns.extend(['FirstName', 'LastName'])

    if not dedupe_columns:
        print("Warning: Cannot remove duplicates - no address/name columns found")
        return df

    # Create normalized comparison columns
    df_compare = df.copy()
    for col in dedupe_columns:
        if col in df_compare.columns:
            df_compare[col] = df_compare[col].astype(str).str.lower().str.strip()

    # Count duplicates before removal
    initial_count = len(df)

    # Find duplicates
    duplicates_mask = df_compare.duplicated(subset=dedupe_columns, keep='first')
    df = df[~duplicates_mask]

    duplicates_removed = initial_count - len(df)
    print(f"Removed {duplicates_removed} duplicate rows based on: {', '.join(dedupe_columns)}")

    return df


def prepare_bcc_format(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepare final DataFrame in BCC Bulk Mailer format with correct column order.

    Args:
        df: pandas DataFrame

    Returns:
        DataFrame with BCC-compliant columns
    """
    # BCC Bulk Mailer expected columns in priority order
    bcc_columns = [
        'FirstName', 'LastName', 'FullName', 'Company',
        'Address1', 'Address2', 'City', 'State', 'ZIP', 'ZIP4',
        'Email', 'Notes'
    ]

    # Ensure all expected columns exist
    for col in bcc_columns:
        if col not in df.columns:
            df[col] = ''

    # Select and reorder columns
    final_columns = [col for col in bcc_columns if col in df.columns]
    df = df[final_columns]

    # Replace NaN with empty strings
    df = df.fillna('')

    return df


def clean_excel_file(input_file: str, output_file: Optional[str] = None) -> str:
    """
    Main function to clean Excel file for BCC Bulk Mailer import.

    Args:
        input_file: Path to input Excel file
        output_file: Path to output CSV file (optional)

    Returns:
        Path to output file
    """
    # Validate input file exists
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    # Determine output file path
    if output_file is None:
        output_file = input_path.stem + "_BCC_cleaned.csv"
    output_path = Path(output_file)

    print(f"Reading Excel file: {input_file}")
    print("=" * 60)

    # Read Excel file
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")

    print(f"Initial rows: {len(df)}")
    print(f"Initial columns: {list(df.columns)}")
    print()

    # Standardize column names
    print("Step 1: Standardizing column names...")
    df = standardize_column_names(df)

    # Parse messy/combined data
    print("\nStep 2: Parsing combined or messy data fields...")
    df = parse_messy_data(df)

    # Clean and validate data
    print("\nStep 3: Cleaning and validating address data...")
    df = clean_and_validate(df)

    # Remove duplicates
    print("\nStep 4: Removing duplicate records...")
    df = remove_duplicates(df)

    # Prepare final BCC format
    print("\nStep 5: Preparing BCC Bulk Mailer format...")
    df = prepare_bcc_format(df)

    # Output final statistics
    print()
    print("=" * 60)
    print(f"Final rows: {len(df)}")
    print(f"Final columns: {list(df.columns)}")
    print()

    # Save to CSV (UTF-8 encoding for BCC Bulk Mailer)
    print(f"Saving cleaned data to: {output_path}")
    df.to_csv(output_path, index=False, encoding='utf-8')

    print()
    print("✓ Done! File ready for BCC Bulk Mailer import.")
    print()

    return str(output_path)


def main():
    """Main entry point for command-line usage"""
    parser = argparse.ArgumentParser(
        description='Clean Excel data for BCC Bulk Mailer: parse messy data, remove duplicates, validate addresses',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python clean_excel_data.py contacts.xlsx
  python clean_excel_data.py messy_data.xlsx -o clean_mailing_list.csv

Output Format:
  CSV file ready for BCC Bulk Mailer Business/Professional import
  Columns: FullName, Company, Address1, Address2, City, State, ZIP, ZIP4, Email, Notes
        """
    )
    parser.add_argument(
        'input_file',
        help='Path to input Excel file (.xlsx or .xls)'
    )
    parser.add_argument(
        '-o', '--output',
        dest='output_file',
        help='Path to output CSV file (default: <input>_BCC_cleaned.csv)',
        default=None
    )

    args = parser.parse_args()

    try:
        clean_excel_file(args.input_file, args.output_file)
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
