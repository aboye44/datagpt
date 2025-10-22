#!/usr/bin/env python3
"""
Excel Data Cleaning Script

This script reads an Excel file, removes duplicate rows based on matching
address + name, standardizes column names, and outputs a clean CSV file.
"""

import pandas as pd
import argparse
import sys
from pathlib import Path


# Column name mapping for standardization
COLUMN_MAPPING = {
    # FirstName variations
    'first_name': 'FirstName',
    'firstname': 'FirstName',
    'first': 'FirstName',
    'fname': 'FirstName',
    'given_name': 'FirstName',
    'givenname': 'FirstName',

    # LastName variations
    'last_name': 'LastName',
    'lastname': 'LastName',
    'last': 'LastName',
    'lname': 'LastName',
    'surname': 'LastName',
    'family_name': 'LastName',
    'familyname': 'LastName',

    # Address1 variations
    'address': 'Address1',
    'address1': 'Address1',
    'address_1': 'Address1',
    'street': 'Address1',
    'street_address': 'Address1',
    'streetaddress': 'Address1',
    'addr': 'Address1',
    'addr1': 'Address1',

    # City variations
    'city': 'City',
    'town': 'City',
    'municipality': 'City',

    # State variations
    'state': 'State',
    'province': 'State',
    'region': 'State',
    'st': 'State',

    # ZIP variations
    'zip': 'ZIP',
    'zipcode': 'ZIP',
    'zip_code': 'ZIP',
    'postal_code': 'ZIP',
    'postalcode': 'ZIP',
    'postcode': 'ZIP',
}


def standardize_column_names(df):
    """
    Standardize column names to FirstName/LastName/Address1/City/State/ZIP

    Args:
        df: pandas DataFrame

    Returns:
        DataFrame with standardized column names
    """
    # Create a copy to avoid modifying the original
    df = df.copy()

    # Normalize column names (lowercase, strip whitespace)
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_')

    # Apply column mapping
    df.rename(columns=COLUMN_MAPPING, inplace=True)

    # Ensure required columns exist
    required_columns = ['FirstName', 'LastName', 'Address1', 'City', 'State', 'ZIP']
    missing_columns = [col for col in required_columns if col not in df.columns]

    if missing_columns:
        print(f"Warning: Missing columns after standardization: {missing_columns}")
        print(f"Available columns: {list(df.columns)}")

    return df


def remove_duplicates(df):
    """
    Remove duplicate rows based on matching Address1 + FirstName + LastName

    Args:
        df: pandas DataFrame

    Returns:
        DataFrame with duplicates removed
    """
    # Create a copy to avoid modifying the original
    df = df.copy()

    # Check if required columns exist for deduplication
    dedupe_columns = ['Address1', 'FirstName', 'LastName']
    available_dedupe_columns = [col for col in dedupe_columns if col in df.columns]

    if not available_dedupe_columns:
        print("Warning: Cannot remove duplicates - no address/name columns found")
        return df

    # Normalize values for comparison (strip whitespace, lowercase)
    for col in available_dedupe_columns:
        if df[col].dtype == 'object':  # String columns
            df[col] = df[col].astype(str).str.strip()

    # Count duplicates before removal
    initial_count = len(df)

    # Remove duplicates keeping the first occurrence
    df = df.drop_duplicates(subset=available_dedupe_columns, keep='first')

    duplicates_removed = initial_count - len(df)
    print(f"Removed {duplicates_removed} duplicate rows based on: {', '.join(available_dedupe_columns)}")

    return df


def clean_excel_file(input_file, output_file=None):
    """
    Main function to clean Excel file

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
        output_file = input_path.stem + "_cleaned.csv"
    output_path = Path(output_file)

    print(f"Reading Excel file: {input_file}")

    # Read Excel file
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        raise Exception(f"Error reading Excel file: {e}")

    print(f"Initial rows: {len(df)}")
    print(f"Initial columns: {list(df.columns)}")

    # Standardize column names
    print("\nStandardizing column names...")
    df = standardize_column_names(df)

    # Remove duplicates
    print("\nRemoving duplicates...")
    df = remove_duplicates(df)

    # Output final statistics
    print(f"\nFinal rows: {len(df)}")
    print(f"Final columns: {list(df.columns)}")

    # Save to CSV
    print(f"\nSaving cleaned data to: {output_path}")
    df.to_csv(output_path, index=False)

    print("Done!")
    return str(output_path)


def main():
    """Main entry point for command-line usage"""
    parser = argparse.ArgumentParser(
        description='Clean Excel data: remove duplicates and standardize column names'
    )
    parser.add_argument(
        'input_file',
        help='Path to input Excel file (.xlsx or .xls)'
    )
    parser.add_argument(
        '-o', '--output',
        dest='output_file',
        help='Path to output CSV file (default: <input>_cleaned.csv)',
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
