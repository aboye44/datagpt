#!/usr/bin/env python3
"""
Excel Data Cleaning Script for BCC Bulk Mailer - Enhanced Edition

This script reads Excel files with messy mailing data, parses combined fields,
removes duplicates (including fuzzy matching), validates addresses/emails/phones,
and outputs BCC Bulk Mailer-ready CSV files with detailed reporting.

BCC Bulk Mailer Business/Professional Import Format:
- FullName, Company, Address1, Address2, City, State, ZIP, ZIP4, Phone, Email, Notes
- UTF-8 or ANSI encoded CSV
- Clean headers for automatic field mapping
- USPS-compliant address formatting

Enhanced Features:
- usaddress library for intelligent address parsing
- Fuzzy duplicate matching (catches typos and variations)
- Phone number parsing and validation
- Email validation
- Business name detection and handling
- Detailed processing report
"""

import pandas as pd
import argparse
import sys
import re
from pathlib import Path
from typing import Dict, Tuple, Optional, List
from datetime import datetime
from collections import defaultdict

try:
    import usaddress
    USADDRESS_AVAILABLE = True
except ImportError:
    USADDRESS_AVAILABLE = False
    print("Warning: usaddress library not available. Using fallback parser.")

try:
    from rapidfuzz import fuzz, process
    FUZZY_AVAILABLE = True
except ImportError:
    FUZZY_AVAILABLE = False
    print("Warning: rapidfuzz library not available. Fuzzy matching disabled.")


# Processing statistics tracker
class ProcessingStats:
    def __init__(self):
        self.initial_rows = 0
        self.final_rows = 0
        self.duplicates_removed = 0
        self.fuzzy_duplicates_removed = 0
        self.states_normalized = 0
        self.zips_formatted = 0
        self.combined_fields_parsed = 0
        self.incomplete_addresses = 0
        self.invalid_emails = 0
        self.phone_numbers_formatted = 0
        self.business_names_found = 0
        self.warnings = []
        self.start_time = None
        self.end_time = None

    def add_warning(self, row_num: int, message: str):
        self.warnings.append(f"Row {row_num}: {message}")

    def generate_report(self, input_file: str, output_file: str) -> str:
        """Generate a detailed processing report"""
        duration = (self.end_time - self.start_time).total_seconds() if self.start_time and self.end_time else 0

        report = f"""
{'='*70}
BCC BULK MAILER DATA CLEANING REPORT
{'='*70}
Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
Processing Time: {duration:.2f} seconds

INPUT/OUTPUT
{'─'*70}
Input File:  {input_file}
Output File: {output_file}
Initial Rows: {self.initial_rows}
Final Rows:   {self.final_rows}

DATA CLEANING SUMMARY
{'─'*70}
✓ Removed {self.duplicates_removed} exact duplicate(s)
✓ Removed {self.fuzzy_duplicates_removed} fuzzy duplicate(s) (similar addresses/names)
✓ Parsed {self.combined_fields_parsed} combined address field(s)
✓ Normalized {self.states_normalized} state abbreviation(s)
✓ Formatted {self.zips_formatted} ZIP code(s)
✓ Formatted {self.phone_numbers_formatted} phone number(s)
✓ Identified {self.business_names_found} business name(s)

VALIDATION RESULTS
{'─'*70}
⚠ {self.incomplete_addresses} record(s) with incomplete addresses (flagged in Notes)
⚠ {self.invalid_emails} invalid email address(es) (flagged in Notes)
"""

        if self.warnings:
            report += f"\nWARNINGS ({len(self.warnings)} total)\n{'─'*70}\n"
            # Show first 20 warnings
            for warning in self.warnings[:20]:
                report += f"  • {warning}\n"
            if len(self.warnings) > 20:
                report += f"  ... and {len(self.warnings) - 20} more warnings\n"

        report += f"\n{'='*70}\n"
        report += "✓ File ready for BCC Bulk Mailer import!\n"
        report += f"{'='*70}\n"

        return report


# Global stats object
stats = ProcessingStats()


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

# Common business entity suffixes
BUSINESS_SUFFIXES = [
    'LLC', 'Inc', 'Corp', 'Corporation', 'Company', 'Co', 'Ltd', 'Limited',
    'LLP', 'LP', 'PC', 'PA', 'PLLC', 'LTD', 'L.L.C.', 'INC.', 'CORP.'
]

# Column name mapping for standardization
COLUMN_MAPPING = {
    # Name variations
    'first_name': 'FirstName', 'firstname': 'FirstName', 'first': 'FirstName',
    'fname': 'FirstName', 'given_name': 'FirstName', 'givenname': 'FirstName',
    'last_name': 'LastName', 'lastname': 'LastName', 'last': 'LastName',
    'lname': 'LastName', 'surname': 'LastName', 'family_name': 'LastName',
    'familyname': 'LastName', 'name': 'FullName', 'full_name': 'FullName',
    'fullname': 'FullName', 'contact': 'FullName', 'contact_name': 'FullName',
    'address_name': 'FullName',  # Common in property management systems

    # Company variations
    'company': 'Company', 'business': 'Company', 'organization': 'Company',
    'org': 'Company', 'company_name': 'Company', 'business_name': 'Company',
    'firm': 'Company', 'business': 'Company',

    # Address variations
    'address': 'Unit', 'address1': 'Address1', 'address_1': 'Address1',
    'street': 'Address1', 'street_address': 'Address1', 'streetaddress': 'Address1',
    'addr': 'Address1', 'addr1': 'Address1', 'address_line_1': 'Address1',
    'full_address': '_FullAddressToParse', # Special: needs parsing
    'address2': 'Address2', 'address_2': 'Address2', 'apt': 'Address2',
    'suite': 'Address2', 'unit': 'Unit', 'address_line_2': 'Address2',

    # City variations
    'city': 'City', 'town': 'City', 'municipality': 'City',

    # State variations
    'state': 'State', 'province': 'State', 'region': 'State', 'st': 'State',

    # ZIP variations
    'zip': 'ZIP', 'zipcode': 'ZIP', 'zip_code': 'ZIP', 'postal_code': 'ZIP',
    'postalcode': 'ZIP', 'postcode': 'ZIP', 'zip5': 'ZIP',
    'zip4': 'ZIP4', 'zip_4': 'ZIP4', 'plus4': 'ZIP4',

    # Phone variations
    'phone': 'Phone', 'phone_number': 'Phone', 'telephone': 'Phone', 'tel': 'Phone',
    'mobile': 'Phone', 'cell': 'Phone', 'cell_phone': 'Phone', 'contact_number': 'Phone',

    # Email variations
    'email': 'Email', 'email_address': 'Email', 'e-mail': 'Email', 'e_mail': 'Email',

    # Notes variations
    'notes': 'Notes', 'note': 'Notes', 'comments': 'Notes', 'comment': 'Notes',
    'remarks': 'Notes', 'memo': 'Notes'
}


def parse_phone_number(phone: str) -> str:
    """
    Parse and format phone number to (XXX) XXX-XXXX format.

    Args:
        phone: Phone number in any format

    Returns:
        Formatted phone number or empty string
    """
    if not phone or pd.isna(phone):
        return ''

    # Extract digits only
    digits = re.sub(r'\D', '', str(phone))

    # Handle different lengths
    if len(digits) == 10:
        # Format as (XXX) XXX-XXXX
        return f"({digits[0:3]}) {digits[3:6]}-{digits[6:10]}"
    elif len(digits) == 11 and digits[0] == '1':
        # Strip leading 1 and format
        return f"({digits[1:4]}) {digits[4:7]}-{digits[7:11]}"
    elif len(digits) > 0:
        # Return as-is if it has some digits but wrong length
        return digits

    return ''


def validate_email(email: str) -> Tuple[bool, str]:
    """
    Validate email address format.

    Args:
        email: Email address to validate

    Returns:
        Tuple of (is_valid, cleaned_email)
    """
    if not email or pd.isna(email):
        return (True, '')  # Empty is valid

    email = str(email).strip().lower()

    # Basic email regex
    email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

    if re.match(email_pattern, email):
        return (True, email)
    else:
        return (False, email)


def is_business_name(name: str) -> bool:
    """
    Detect if a name is likely a business rather than a person.

    Args:
        name: Name to check

    Returns:
        True if likely a business name
    """
    if not name or pd.isna(name):
        return False

    name_upper = str(name).upper()

    # Check for business suffixes
    for suffix in BUSINESS_SUFFIXES:
        if re.search(r'\b' + re.escape(suffix) + r'\b', name_upper):
            return True

    # Check for business keywords
    business_keywords = ['SHOP', 'STORE', 'MARKET', 'CENTER', 'SERVICES', 'GROUP', 'ASSOCIATES', 'SOLUTIONS']
    for keyword in business_keywords:
        if keyword in name_upper:
            return True

    return False


def parse_combined_address_usaddress(address_string: str) -> Dict[str, str]:
    """
    Parse combined address using usaddress library (more accurate).

    Args:
        address_string: Combined address string

    Returns:
        Dictionary with parsed components
    """
    result = {
        'FullName': '', 'Company': '', 'Address1': '', 'Address2': '',
        'City': '', 'State': '', 'ZIP': '', 'ZIP4': '', 'Phone': '', 'Email': ''
    }

    if not address_string or pd.isna(address_string):
        return result

    address_string = str(address_string).strip()

    # First, extract phone and email if present
    phone_match = re.search(r'(\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})', address_string)
    if phone_match:
        result['Phone'] = parse_phone_number(phone_match.group(1))
        address_string = address_string.replace(phone_match.group(1), '')

    email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', address_string)
    if email_match:
        result['Email'] = email_match.group(1).lower()
        address_string = address_string.replace(email_match.group(1), '')

    try:
        # Parse with usaddress
        parsed, addr_type = usaddress.tag(address_string)

        # Map usaddress tags to our fields
        address_parts = []
        if 'AddressNumber' in parsed:
            address_parts.append(parsed['AddressNumber'])
        if 'StreetNamePreDirectional' in parsed:
            address_parts.append(parsed['StreetNamePreDirectional'])
        if 'StreetName' in parsed:
            address_parts.append(parsed['StreetName'])
        if 'StreetNamePostType' in parsed:
            address_parts.append(parsed['StreetNamePostType'])
        if 'StreetNamePostDirectional' in parsed:
            address_parts.append(parsed['StreetNamePostDirectional'])

        result['Address1'] = ' '.join(address_parts) if address_parts else ''

        # Address2 (apartment, suite, etc.)
        address2_parts = []
        if 'OccupancyType' in parsed:
            address2_parts.append(parsed['OccupancyType'])
        if 'OccupancyIdentifier' in parsed:
            address2_parts.append(parsed['OccupancyIdentifier'])

        result['Address2'] = ' '.join(address2_parts) if address2_parts else ''

        # City, State, ZIP
        result['City'] = parsed.get('PlaceName', '')
        result['State'] = parsed.get('StateName', '')

        zipcode = parsed.get('ZipCode', '')
        if zipcode:
            if '-' in zipcode:
                zip_parts = zipcode.split('-')
                result['ZIP'] = zip_parts[0]
                result['ZIP4'] = zip_parts[1] if len(zip_parts) > 1 else ''
            else:
                result['ZIP'] = zipcode

        # Recipient (person or business name)
        recipient_parts = []
        if 'Recipient' in parsed:
            recipient_parts.append(parsed['Recipient'])

        recipient = ' '.join(recipient_parts) if recipient_parts else ''
        if recipient:
            if is_business_name(recipient):
                result['Company'] = recipient
            else:
                result['FullName'] = recipient

    except Exception as e:
        # Fallback to regex-based parsing if usaddress fails
        return parse_combined_address_regex(address_string)

    return result


def parse_combined_address_regex(address_string: str) -> Dict[str, str]:
    """
    Parse combined address using regex (fallback method).

    Args:
        address_string: Combined address string

    Returns:
        Dictionary with parsed components
    """
    result = {
        'FullName': '', 'Company': '', 'Address1': '', 'Address2': '',
        'City': '', 'State': '', 'ZIP': '', 'ZIP4': '', 'Phone': '', 'Email': ''
    }

    if not address_string or pd.isna(address_string):
        return result

    address_string = str(address_string).strip()

    # Extract phone
    phone_match = re.search(r'(\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})', address_string)
    if phone_match:
        result['Phone'] = parse_phone_number(phone_match.group(1))
        address_string = address_string.replace(phone_match.group(1), '')

    # Extract email
    email_match = re.search(r'([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})', address_string)
    if email_match:
        result['Email'] = email_match.group(1).lower()
        address_string = address_string.replace(email_match.group(1), '')

    # Try to extract ZIP code first (most reliable anchor)
    zip_pattern = r'\b(\d{5})(?:-(\d{4}))?\b'
    zip_match = re.search(zip_pattern, address_string)
    if zip_match:
        result['ZIP'] = zip_match.group(1)
        if zip_match.group(2):
            result['ZIP4'] = zip_match.group(2)
        address_string = address_string[:zip_match.start()] + address_string[zip_match.end():]

    # Try to extract state
    state_pattern = r'\b([A-Z]{2})\b'
    state_matches = re.findall(state_pattern, address_string)
    if state_matches:
        potential_state = state_matches[-1]
        if potential_state in STATE_ABBREVIATIONS.values():
            result['State'] = potential_state
            address_string = re.sub(r'\b' + potential_state + r'\b', '', address_string, count=1)

    # Split by commas or newlines
    parts = [p.strip() for p in re.split(r'[,\n]', address_string) if p.strip()]

    if not parts:
        return result

    # Heuristic parsing
    if len(parts) >= 1:
        first_part = parts[0]
        # Check if it's a business name
        if is_business_name(first_part):
            result['Company'] = first_part
            parts = parts[1:]
        elif not re.search(r'\d', first_part):  # No numbers = probably a name
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
        result['City'] = parts[-1] if parts else ''
        if len(parts) > 1:
            if re.search(r'\b(apt|suite|unit|ste|#)\b', parts[-2], re.IGNORECASE):
                result['Address2'] = parts[-2]
            else:
                result['City'] = parts[-2] + ' ' + result['City']

    return result


def parse_name(name_string: str) -> Tuple[str, str]:
    """
    Parse a name string into FirstName and LastName.
    Handles "Last, First", "First Last", titles, and suffixes.

    Args:
        name_string: Name to parse

    Returns:
        Tuple of (FirstName, LastName)
    """
    if not name_string or pd.isna(name_string):
        return ('', '')

    name_string = str(name_string).strip()

    # Remove titles
    titles = ['Dr.', 'Mr.', 'Mrs.', 'Ms.', 'Miss', 'Prof.', 'Rev.', 'Dr', 'Mr', 'Mrs', 'Ms']
    for title in titles:
        name_string = re.sub(r'\b' + re.escape(title) + r'\b', '', name_string, flags=re.IGNORECASE)

    # Remove suffixes
    suffixes = ['Jr.', 'Sr.', 'II', 'III', 'IV', 'Esq.', 'Jr', 'Sr', 'Esq']
    for suffix in suffixes:
        name_string = re.sub(r'\b' + re.escape(suffix) + r'\b', '', name_string, flags=re.IGNORECASE)

    name_string = name_string.strip()

    # Check for "Last, First" format
    if ',' in name_string:
        parts = [p.strip() for p in name_string.split(',', 1)]
        if len(parts) == 2:
            return (parts[1], parts[0])

    # Split on whitespace
    parts = name_string.split()

    if len(parts) == 0:
        return ('', '')
    elif len(parts) == 1:
        return ('', parts[0])
    elif len(parts) == 2:
        return (parts[0], parts[1])
    else:
        # 3+ parts: First name is first part, last name is last part
        return (parts[0], parts[-1])


def normalize_state(state_value: str) -> Tuple[str, bool]:
    """
    Normalize state to USPS 2-letter abbreviation.

    Args:
        state_value: State name or abbreviation

    Returns:
        Tuple of (normalized_state, was_changed)
    """
    if not state_value or pd.isna(state_value):
        return ('', False)

    original = state_value
    state_value = str(state_value).strip()

    # Already a 2-letter code
    if len(state_value) == 2:
        state_upper = state_value.upper()
        if state_upper in STATE_ABBREVIATIONS.values():
            return (state_upper, state_upper != original)

    # Try to match full state name
    state_lower = state_value.lower()
    if state_lower in STATE_ABBREVIATIONS:
        return (STATE_ABBREVIATIONS[state_lower], True)

    # Try with punctuation removed
    state_lower = re.sub(r'[^\w\s]', '', state_lower)
    if state_lower in STATE_ABBREVIATIONS:
        return (STATE_ABBREVIATIONS[state_lower], True)

    return (state_value, False)


def validate_zip(zip_value: str) -> Tuple[str, str, bool]:
    """
    Validate and format ZIP code.

    Args:
        zip_value: ZIP code value

    Returns:
        Tuple of (ZIP5, ZIP4, was_formatted)
    """
    if not zip_value or pd.isna(zip_value):
        return ('', '', False)

    original = zip_value
    zip_string = str(zip_value).strip()

    # Extract 5-digit ZIP and optional +4
    zip_pattern = r'(\d{5})(?:-?(\d{4}))?'
    match = re.search(zip_pattern, zip_string)

    if match:
        zip5 = match.group(1)
        zip4 = match.group(2) if match.group(2) else ''
        was_changed = (zip5 + ('-' + zip4 if zip4 else '')) != original
        return (zip5, zip4, was_changed)

    return ('', '', False)


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
    upper_words = {'PO', 'NW', 'NE', 'SW', 'SE', 'N', 'S', 'E', 'W', 'US', 'USA', 'PMB', 'APT', 'STE'}

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

    # Normalize column names
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

    # Handle special "_FullAddressToParse" column (from "Full Address" in input)
    if '_FullAddressToParse' in df.columns:
        print("Detected 'Full Address' column - parsing addresses...")

        parsed_addresses = []
        for idx, addr_string in df['_FullAddressToParse'].items():
            if pd.notna(addr_string) and str(addr_string).strip():
                parsed = parse_combined_address_regex(str(addr_string))
                parsed_addresses.append(parsed)
            else:
                parsed_addresses.append({
                    'FullName': '', 'Company': '', 'Address1': '', 'Address2': '',
                    'City': '', 'State': '', 'ZIP': '', 'ZIP4': '', 'Phone': '', 'Email': ''
                })

        # Add parsed columns to dataframe
        for key in ['Address1', 'Address2', 'City', 'State', 'ZIP', 'ZIP4', 'Phone', 'Email']:
            if key not in df.columns:
                df[key] = [p[key] for p in parsed_addresses]

        # Drop the temporary column
        df = df.drop(columns=['_FullAddressToParse'])
        stats.combined_fields_parsed = len(df)
        print(f"Parsed {len(df)} addresses from 'Full Address' column")

    # Check if we have very few columns - might indicate combined data
    if len(df.columns) <= 2:
        print("Detected possible combined data format - attempting to parse...")

        # Find the column with the most data
        main_col = None
        for col in df.columns:
            if df[col].notna().sum() > 0:
                main_col = col
                break

        if main_col:
            parsed_rows = []
            for idx, row in df.iterrows():
                if USADDRESS_AVAILABLE:
                    parsed = parse_combined_address_usaddress(row[main_col])
                else:
                    parsed = parse_combined_address_regex(row[main_col])
                parsed_rows.append(parsed)

            df = pd.DataFrame(parsed_rows)
            stats.combined_fields_parsed = len(df)
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

    # Detect business names in FullName and move to Company
    if 'FullName' in df.columns:
        for idx, row in df.iterrows():
            if pd.notna(row['FullName']) and is_business_name(row['FullName']):
                if 'Company' not in df.columns or pd.isna(row.get('Company', '')):
                    if 'Company' not in df.columns:
                        df['Company'] = ''
                    df.at[idx, 'Company'] = row['FullName']
                    df.at[idx, 'FullName'] = ''
                    stats.business_names_found += 1

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
        df[col] = df[col].replace('nan', '')

    # Handle Unit column (from property management systems)
    if 'Unit' in df.columns:
        print("Incorporating unit numbers into Address2...")
        for idx, unit in df['Unit'].items():
            if unit and unit != '':
                # Format unit with # prefix if not already there
                unit_formatted = f"#{unit}" if not str(unit).startswith('#') else str(unit)

                # Add to Address2, or create Address2 if it doesn't exist
                if 'Address2' not in df.columns:
                    df['Address2'] = ''

                # Prepend unit to existing Address2 if present, otherwise just use unit
                current_addr2 = df.at[idx, 'Address2'] if pd.notna(df.at[idx, 'Address2']) and df.at[idx, 'Address2'] != '' else ''
                if current_addr2:
                    df.at[idx, 'Address2'] = f"{unit_formatted} {current_addr2}"
                else:
                    df.at[idx, 'Address2'] = unit_formatted

        # Drop Unit column after incorporating it
        df = df.drop(columns=['Unit'])

    # Normalize State
    if 'State' in df.columns:
        print("Normalizing state abbreviations...")
        state_data = df['State'].apply(normalize_state)
        df['State'] = state_data.apply(lambda x: x[0])
        stats.states_normalized = state_data.apply(lambda x: x[1]).sum()

    # Validate and format ZIP codes
    if 'ZIP' in df.columns:
        print("Validating ZIP codes...")
        zip_data = df['ZIP'].apply(validate_zip)
        df['ZIP'] = zip_data.apply(lambda x: x[0])
        if 'ZIP4' not in df.columns or df['ZIP4'].fillna('').eq('').all():
            df['ZIP4'] = zip_data.apply(lambda x: x[1])
        stats.zips_formatted = zip_data.apply(lambda x: x[2]).sum()

    # Format phone numbers
    if 'Phone' in df.columns:
        print("Formatting phone numbers...")
        original_phones = df['Phone'].copy()
        df['Phone'] = df['Phone'].apply(parse_phone_number)
        stats.phone_numbers_formatted = (original_phones != df['Phone']).sum()

    # Validate emails
    if 'Email' in df.columns:
        print("Validating email addresses...")
        for idx, email in df['Email'].items():
            if email and email != '':
                is_valid, cleaned = validate_email(email)
                df.at[idx, 'Email'] = cleaned
                if not is_valid:
                    stats.invalid_emails += 1
                    stats.add_warning(idx + 2, f"Invalid email: {email}")
                    if 'Notes' not in df.columns:
                        df['Notes'] = ''
                    df.at[idx, 'Notes'] = (df.at[idx, 'Notes'] + '; ' if df.at[idx, 'Notes'] else '') + 'INVALID_EMAIL'

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
    for idx, row in df.iterrows():
        incomplete = False
        for field in required_fields:
            if field in df.columns and (pd.isna(row[field]) or row[field] == ''):
                incomplete = True
                break

        if incomplete:
            stats.incomplete_addresses += 1
            stats.add_warning(idx + 2, "Incomplete address (missing required fields)")
            if 'Notes' not in df.columns:
                df['Notes'] = ''
            df.at[idx, 'Notes'] = (df.at[idx, 'Notes'] + '; ' if df.at[idx, 'Notes'] else '') + 'INCOMPLETE_ADDRESS'

    return df


def remove_duplicates(df: pd.DataFrame, fuzzy_threshold: int = 85) -> pd.DataFrame:
    """
    Remove duplicate rows using exact and fuzzy matching.

    Args:
        df: pandas DataFrame
        fuzzy_threshold: Similarity threshold for fuzzy matching (0-100)

    Returns:
        DataFrame with duplicates removed
    """
    df = df.copy()

    # Determine deduplication columns
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

    initial_count = len(df)

    # STEP 1: Exact duplicate removal
    df_compare = df.copy()
    for col in dedupe_columns:
        if col in df_compare.columns:
            df_compare[col] = df_compare[col].astype(str).str.lower().str.strip()

    exact_duplicates_mask = df_compare.duplicated(subset=dedupe_columns, keep='first')
    df = df[~exact_duplicates_mask]
    stats.duplicates_removed = exact_duplicates_mask.sum()

    print(f"Removed {stats.duplicates_removed} exact duplicate(s)")

    # STEP 2: Fuzzy duplicate removal (if library available)
    if FUZZY_AVAILABLE and len(df) > 1:
        print(f"Checking for fuzzy duplicates (similarity ≥ {fuzzy_threshold}%)...")

        # Create comparison strings
        df['_compare_str'] = ''
        for col in dedupe_columns:
            if col in df.columns:
                df['_compare_str'] += df[col].astype(str).str.lower().str.strip() + ' '
        df['_compare_str'] = df['_compare_str'].str.strip()

        # Find fuzzy duplicates
        to_remove = set()
        compare_strings = df['_compare_str'].tolist()

        for i in range(len(df)):
            if i in to_remove:
                continue

            current = compare_strings[i]
            if not current or current == '':
                continue

            for j in range(i + 1, len(df)):
                if j in to_remove:
                    continue

                other = compare_strings[j]
                if not other or other == '':
                    continue

                # Calculate similarity
                similarity = fuzz.ratio(current, other)

                if similarity >= fuzzy_threshold:
                    to_remove.add(j)
                    stats.fuzzy_duplicates_removed += 1

        # Remove fuzzy duplicates
        if to_remove:
            df = df.drop(df.index[list(to_remove)])
            df = df.reset_index(drop=True)
            print(f"Removed {len(to_remove)} fuzzy duplicate(s)")

        # Clean up comparison column
        df = df.drop(columns=['_compare_str'])

    final_count = len(df)
    total_removed = initial_count - final_count
    print(f"Total duplicates removed: {total_removed}")

    return df


def prepare_bcc_format(df: pd.DataFrame) -> pd.DataFrame:
    """
    Prepare final DataFrame in BCC Bulk Mailer format with correct column order.

    Args:
        df: pandas DataFrame

    Returns:
        DataFrame with BCC-compliant columns
    """
    # BCC Bulk Mailer expected columns
    bcc_columns = [
        'FirstName', 'LastName', 'FullName', 'Company',
        'Address1', 'Address2', 'City', 'State', 'ZIP', 'ZIP4',
        'Phone', 'Email', 'Notes'
    ]

    # Ensure all expected columns exist
    for col in bcc_columns:
        if col not in df.columns:
            df[col] = ''

    # Select and reorder columns
    df = df[bcc_columns]

    # Replace NaN with empty strings
    df = df.fillna('')

    return df


def clean_excel_file(input_file: str, output_file: Optional[str] = None,
                     fuzzy_threshold: int = 85, save_report: bool = True) -> str:
    """
    Main function to clean Excel file for BCC Bulk Mailer import.

    Args:
        input_file: Path to input Excel file
        output_file: Path to output CSV file (optional)
        fuzzy_threshold: Similarity threshold for fuzzy duplicate matching (0-100)
        save_report: Whether to save a text report file

    Returns:
        Path to output file
    """
    global stats
    stats = ProcessingStats()  # Reset stats
    stats.start_time = datetime.now()

    # Validate input file exists
    input_path = Path(input_file)
    if not input_path.exists():
        raise FileNotFoundError(f"Input file not found: {input_file}")

    # Determine output file path
    if output_file is None:
        output_file = input_path.stem + "_BCC_cleaned.csv"
    output_path = Path(output_file)

    print(f"\n{'='*70}")
    print("BCC BULK MAILER DATA CLEANING TOOL")
    print(f"{'='*70}")
    print(f"Input:  {input_file}")
    print(f"Output: {output_file}")
    print(f"{'='*70}\n")

    # Read Excel file
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")

    stats.initial_rows = len(df)
    print(f"Initial rows: {stats.initial_rows}")
    print(f"Initial columns: {list(df.columns)}\n")

    # Processing steps
    print("Step 1: Standardizing column names...")
    df = standardize_column_names(df)

    print("\nStep 2: Parsing combined or messy data fields...")
    df = parse_messy_data(df)

    print("\nStep 3: Cleaning and validating address data...")
    df = clean_and_validate(df)

    print(f"\nStep 4: Removing duplicate records (fuzzy threshold: {fuzzy_threshold}%)...")
    df = remove_duplicates(df, fuzzy_threshold)

    print("\nStep 5: Preparing BCC Bulk Mailer format...")
    df = prepare_bcc_format(df)

    stats.final_rows = len(df)
    stats.end_time = datetime.now()

    # Save to CSV
    print(f"\nSaving cleaned data to: {output_path}")
    df.to_csv(output_path, index=False, encoding='utf-8')

    # Generate and display report
    report = stats.generate_report(str(input_path), str(output_path))
    print(report)

    # Save report to file
    if save_report:
        report_path = output_path.with_suffix('.txt')
        with open(report_path, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"Detailed report saved to: {report_path}\n")

    return str(output_path)


def main():
    """Main entry point for command-line usage"""
    parser = argparse.ArgumentParser(
        description='Clean Excel data for BCC Bulk Mailer: parse messy data, remove duplicates, validate addresses',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python clean_excel_data.py contacts.xlsx
  python clean_excel_data.py messy_data.xlsx -o clean_list.csv
  python clean_excel_data.py data.xlsx --fuzzy-threshold 90

Features:
  ✓ Intelligent address parsing (usaddress library)
  ✓ Fuzzy duplicate matching (catches typos and variations)
  ✓ Phone number formatting
  ✓ Email validation
  ✓ Business name detection
  ✓ Detailed processing report

Output Format:
  CSV file ready for BCC Bulk Mailer Business/Professional import
  Columns: FirstName, LastName, FullName, Company, Address1, Address2,
           City, State, ZIP, ZIP4, Phone, Email, Notes
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
    parser.add_argument(
        '--fuzzy-threshold',
        type=int,
        default=85,
        help='Fuzzy matching threshold for duplicates, 0-100 (default: 85)'
    )
    parser.add_argument(
        '--no-report',
        action='store_true',
        help='Do not save a separate report file'
    )

    args = parser.parse_args()

    # Validate fuzzy threshold
    if not 0 <= args.fuzzy_threshold <= 100:
        print("Error: fuzzy-threshold must be between 0 and 100", file=sys.stderr)
        return 1

    try:
        clean_excel_file(
            args.input_file,
            args.output_file,
            fuzzy_threshold=args.fuzzy_threshold,
            save_report=not args.no_report
        )
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        return 1


if __name__ == '__main__':
    sys.exit(main())
