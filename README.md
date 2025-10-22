# Excel Data Cleaning Script for BCC Bulk Mailer

A Python script that transforms messy Excel mailing lists into clean, validated CSV files ready for **BCC Bulk Mailer Business or Professional** import. Handles combined data fields, removes duplicates, validates addresses, and outputs USPS-compliant mailing lists.

## Features

### Data Cleaning
- **Parses combined address fields** - Handles data like "John Doe, 123 Main St, Boston, MA 02101" in a single cell
- **Splits names** - Intelligently handles "Last, First" and "First Last" formats
- **Removes duplicates** - Based on matching Address + Name (case-insensitive)
- **Validates ZIP codes** - Extracts and formats 5-digit ZIP and ZIP+4 extensions
- **Normalizes states** - Converts full state names to USPS 2-letter abbreviations
- **Proper capitalization** - Title cases addresses while preserving abbreviations (PO, NW, etc.)
- **Flags incomplete records** - Marks records missing required address fields in Notes column

### BCC Bulk Mailer Compliance
- **Intelligent field mapping** - Recognizes 60+ column name variations
- **USPS-compliant output** - Follows postal standards for addressing
- **UTF-8 CSV format** - Ready for direct import into BCC Bulk Mailer
- **Clean headers** - Automatic field mapping in BCC without manual configuration

## Output Format

The script produces CSV files with these columns in BCC-compatible format:

| Column | Required | Description |
|--------|----------|-------------|
| **FirstName** | Optional* | First name (auto-split from FullName if needed) |
| **LastName** | Optional* | Last name (auto-split from FullName if needed) |
| **FullName** | Recommended | Combined name field |
| **Company** | Optional | Business or organization name |
| **Address1** | Required | Primary street address (delivery line) |
| **Address2** | Optional | Apt/Suite/Unit (improves CASS validation) |
| **City** | Required | City name |
| **State** | Required | 2-letter USPS abbreviation |
| **ZIP** | Required | 5-digit ZIP code |
| **ZIP4** | Optional | 4-digit ZIP extension (improves presort) |
| **Email** | Optional | Email address |
| **Notes** | Optional | Flags/comments (auto-populated for incomplete records) |

*FirstName/LastName are optional because FullName can be used instead

## Installation

1. Clone or download this repository
2. Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python clean_excel_data.py contacts.xlsx
```

This creates `contacts_BCC_cleaned.csv` ready for BCC Bulk Mailer import.

### Custom Output File

```bash
python clean_excel_data.py messy_data.xlsx -o mailing_list_2025.csv
```

### Help

```bash
python clean_excel_data.py --help
```

## How It Works

The script performs 5 automated steps:

1. **Standardize Column Names** - Maps 60+ variations to BCC standard fields
2. **Parse Combined Data** - Extracts separate fields from combined cells using regex
3. **Clean & Validate** - Normalizes states, validates ZIPs, proper capitalization
4. **Remove Duplicates** - Case-insensitive matching on Address + Name
5. **Format for BCC** - Arranges columns and flags incomplete records

## Handling Messy Data

### Combined Address in One Cell

**Input Excel:**
```
Column A
-----------------
John Doe, 123 Main St, Boston, MA 02101
Jane Smith, 456 Oak Ave, Apt 2B, Austin, TX 78701
```

**Output CSV:**
```csv
FirstName,LastName,FullName,Company,Address1,Address2,City,State,ZIP,ZIP4,Email,Notes
John,Doe,John Doe,,123 Main St,,Boston,MA,02101,,,,
Jane,Smith,Jane Smith,,456 Oak Ave,Apt 2B,Austin,TX,78701,,,,
```

### Mixed Column Names

The script recognizes variations like:
- Names: `first name`, `fname`, `First`, `given_name` → **FirstName**
- Address: `street address`, `addr1`, `address_line_1` → **Address1**
- State: `province`, `region`, `st` → **State**
- ZIP: `postal_code`, `zipcode`, `zip5` → **ZIP**

### Full State Names

```
Input: Massachusetts, New York, Texas
Output: MA, NY, TX
```

### Name Formats

```
Input: "Doe, John" or "John Doe" or "John Q. Doe"
Output: FirstName="John", LastName="Doe", FullName="John Doe"
```

### ZIP Code Formats

```
Input: "02101-1234" or "02101" or "ZIP: 02101"
Output: ZIP="02101", ZIP4="1234"
```

## Real-World Example

**Messy Input Excel (all in one column):**
```
Contact Info
--------------------------------
John Smith, 123 main street, boston, massachusetts 02101
JANE DOE, 456 ELM AVE, APT 3, SPRINGFIELD, IL, 62704-5678
Bob Johnson, 789 Oak Blvd, Chicago Illinois 60601
John Smith, 123 Main Street, Boston, MA 02101
```

**Clean Output CSV:**
```csv
FirstName,LastName,FullName,Company,Address1,Address2,City,State,ZIP,ZIP4,Email,Notes
John,Smith,John Smith,,123 Main Street,,Boston,MA,02101,,,,
Jane,Doe,Jane Doe,,456 Elm Ave,Apt 3,Springfield,IL,62704,5678,,
Bob,Johnson,Bob Johnson,,789 Oak Blvd,,Chicago,IL,60601,,,,
```

**Processing Notes:**
- Removed 1 duplicate (John Smith)
- Normalized state names (massachusetts → MA, Illinois → IL)
- Proper capitalization (main street → Main Street)
- Extracted ZIP+4 extension (62704-5678 → ZIP: 62704, ZIP4: 5678)
- Parsed apartment/suite info into Address2

## Column Name Recognition

The script automatically recognizes these column name patterns (case-insensitive):

### Name Fields
- **FirstName**: first_name, firstname, first, fname, given_name, givenname
- **LastName**: last_name, lastname, last, lname, surname, family_name, familyname
- **FullName**: name, full_name, fullname, contact, contact_name

### Address Fields
- **Address1**: address, address1, street, street_address, addr, addr1, address_line_1
- **Address2**: address2, apt, suite, unit, address_line_2
- **City**: city, town, municipality
- **State**: state, province, region, st
- **ZIP**: zip, zipcode, zip_code, postal_code, postcode, zip5
- **ZIP4**: zip4, zip_4, plus4

### Other Fields
- **Company**: company, business, organization, org, company_name, business_name
- **Email**: email, email_address, e-mail, mail
- **Notes**: notes, note, comments, comment, remarks, memo

## Importing into BCC Bulk Mailer

1. Run the script on your Excel file
2. Open BCC Bulk Mailer Business/Professional
3. Use the import function - BCC will auto-detect headers
4. Field mapping happens automatically due to standard column names
5. Proceed with CASS validation and presort

The output is specifically designed for BCC's intelligent field mapping system, eliminating manual configuration.

## Data Validation Features

- **ZIP Code Format**: Validates against `^\d{5}(-\d{4})?$` pattern
- **State Abbreviations**: Converts to USPS standard 2-letter codes
- **Address Capitalization**: Proper title case with exceptions (PO, NW, etc.)
- **Duplicate Detection**: Case-insensitive, whitespace-trimmed comparison
- **Incomplete Address Flagging**: Marks records missing Address1, City, State, or ZIP

Incomplete records are marked in the **Notes** column with `INCOMPLETE_ADDRESS` for manual review before CASS processing.

## Requirements

- Python 3.7+
- pandas >= 2.0.0
- openpyxl >= 3.1.0

## Technical Details

### Parsing Algorithm

1. **ZIP extraction** - Most reliable anchor point (regex: `\b\d{5}(-\d{4})?\b`)
2. **State detection** - 2-letter codes validated against USPS list
3. **Address identification** - Lines containing numbers
4. **Name inference** - Lines without numbers before address
5. **City extraction** - Remaining text after state/ZIP removal

### Duplicate Logic

Duplicates are identified by case-insensitive, whitespace-trimmed comparison of:
- Address1 + FullName (if available)
- Address1 + FirstName + LastName (fallback)

First occurrence is kept, subsequent duplicates are removed.

## Troubleshooting

### "Missing columns after standardization"
Your Excel file has column names the script doesn't recognize. Check the column name recognition table above or add custom mappings.

### "Cannot remove duplicates - no address/name columns found"
The data format is very unusual. The script will still output data but cannot deduplicate.

### "Warning: X records have incomplete address data"
These records are flagged in the Notes column. Review them manually or they may fail CASS validation in BCC.

### Empty output file
Check that your Excel file has data in the first sheet and that at least one column contains address information.

## License

MIT License
