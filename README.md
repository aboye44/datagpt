# Excel Data Cleaning Script

A Python script that reads Excel files, removes duplicate rows based on matching address + name, standardizes column names, and outputs clean CSV files.

## Features

- **Reads Excel files** (.xlsx and .xls formats)
- **Removes duplicates** based on matching Address + FirstName + LastName
- **Standardizes column names** to: FirstName, LastName, Address1, City, State, ZIP
- **Outputs clean CSV** files for easy processing

## Installation

1. Clone this repository
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python clean_excel_data.py input.xlsx
```

This will create a cleaned CSV file named `input_cleaned.csv` in the same directory.

### Specify Output File

```bash
python clean_excel_data.py input.xlsx -o output.csv
```

### Help

```bash
python clean_excel_data.py --help
```

## Column Name Mapping

The script automatically recognizes and standardizes various column name formats:

| Standard Name | Recognized Variations |
|--------------|----------------------|
| **FirstName** | first_name, firstname, first, fname, given_name, givenname |
| **LastName** | last_name, lastname, last, lname, surname, family_name, familyname |
| **Address1** | address, address1, address_1, street, street_address, streetaddress, addr, addr1 |
| **City** | city, town, municipality |
| **State** | state, province, region, st |
| **ZIP** | zip, zipcode, zip_code, postal_code, postalcode, postcode |

## How It Works

1. **Reads** the Excel file using pandas
2. **Normalizes** column names (lowercase, removes spaces)
3. **Maps** column names to standard format
4. **Removes duplicates** by comparing Address1 + FirstName + LastName (case-insensitive, whitespace-trimmed)
5. **Outputs** clean CSV file

## Example

**Input Excel file (contacts.xlsx):**

| first name | last name | street address | city | state | zip code |
|------------|-----------|----------------|------|-------|----------|
| John | Doe | 123 Main St | Boston | MA | 02101 |
| Jane | Smith | 456 Oak Ave | Austin | TX | 78701 |
| John | Doe | 123 Main St | Boston | MA | 02101 |

**Command:**

```bash
python clean_excel_data.py contacts.xlsx
```

**Output (contacts_cleaned.csv):**

| FirstName | LastName | Address1 | City | State | ZIP |
|-----------|----------|----------|------|-------|-----|
| John | Doe | 123 Main St | Boston | MA | 02101 |
| Jane | Smith | 456 Oak Ave | Austin | TX | 78701 |

## Requirements

- Python 3.7+
- pandas >= 2.0.0
- openpyxl >= 3.1.0

## License

MIT License
