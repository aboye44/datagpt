# Excel Data Cleaning Script for BCC Bulk Mailer - Enhanced Edition

A powerful Python script that transforms messy Excel mailing lists into clean, validated CSV files ready for **BCC Bulk Mailer Business or Professional** import. Perfect for teams handling diverse data sources with varying quality levels.

## 🎯 Key Features

### Advanced Data Processing
- **🔍 usaddress Library Integration** - Professional-grade address parsing (handles complex formats)
- **🎯 Fuzzy Duplicate Matching** - Catches typos and variations (e.g., "123 Main St" vs "123 Main Street")
- **📊 Detailed Reporting** - Know exactly what changed with automatic report generation
- **📞 Phone Number Formatting** - Standardizes to (XXX) XXX-XXXX format
- **✉️ Email Validation** - Validates format and flags invalid addresses
- **🏢 Business Name Detection** - Automatically identifies and separates business vs. personal names

### Data Cleaning Capabilities
- **Parses combined address fields** - Handles "John Doe, 123 Main St, Boston, MA 02101" in a single cell
- **Splits names** - Intelligently handles "Last, First", "First Last", titles (Dr., Mr.), and suffixes (Jr., Sr.)
- **Removes duplicates** - Exact matching + fuzzy matching for similar records
- **Validates ZIP codes** - Extracts and formats 5-digit ZIP and ZIP+4 extensions
- **Normalizes states** - Converts full state names to USPS 2-letter abbreviations
- **Proper capitalization** - Title cases addresses while preserving abbreviations (PO, NW, etc.)
- **Flags incomplete records** - Marks records missing required address fields

### BCC Bulk Mailer Compliance
- **Intelligent field mapping** - Recognizes 70+ column name variations
- **USPS-compliant output** - Follows postal standards for addressing
- **UTF-8 CSV format** - Ready for direct import into BCC Bulk Mailer
- **Clean headers** - Automatic field mapping in BCC without manual configuration

## 📋 Output Format

The script produces CSV files with these columns in BCC-compatible format:

| Column | Required | Description |
|--------|----------|-------------|
| **FirstName** | Optional* | First name (auto-split from FullName if needed) |
| **LastName** | Optional* | Last name (auto-split from FullName if needed) |
| **FullName** | Recommended | Combined name field |
| **Company** | Optional | Business or organization name (auto-detected) |
| **Address1** | Required | Primary street address (delivery line) |
| **Address2** | Optional | Apt/Suite/Unit (improves CASS validation) |
| **City** | Required | City name |
| **State** | Required | 2-letter USPS abbreviation |
| **ZIP** | Required | 5-digit ZIP code |
| **ZIP4** | Optional | 4-digit ZIP extension (improves presort) |
| **Phone** | Optional | Phone number in (XXX) XXX-XXXX format |
| **Email** | Optional | Email address (validated) |
| **Notes** | Optional | Flags/comments (auto-populated for incomplete/invalid records) |

*FirstName/LastName are optional because FullName can be used instead

## ⚙️ Installation

1. Clone or download this repository
2. Install dependencies:

```bash
pip install -r requirements.txt
```

**Dependencies:**
- pandas - Data processing
- openpyxl - Excel file reading
- usaddress - Intelligent address parsing
- rapidfuzz - Fuzzy string matching for duplicates

## 🚀 Usage

### Basic Usage

```bash
python clean_excel_data.py contacts.xlsx
```

This creates `contacts_BCC_cleaned.csv` and `contacts_BCC_cleaned.txt` (detailed report).

### Custom Output File

```bash
python clean_excel_data.py messy_data.xlsx -o mailing_list_2025.csv
```

### Adjust Fuzzy Matching Sensitivity

```bash
# More strict (fewer fuzzy matches) - 90% similarity required
python clean_excel_data.py data.xlsx --fuzzy-threshold 90

# More lenient (more fuzzy matches) - 80% similarity required
python clean_excel_data.py data.xlsx --fuzzy-threshold 80
```

### Skip Report File

```bash
python clean_excel_data.py data.xlsx --no-report
```

### Help

```bash
python clean_excel_data.py --help
```

## 📊 Sample Report Output

```
======================================================================
BCC BULK MAILER DATA CLEANING REPORT
======================================================================
Generated: 2025-10-22 14:30:15
Processing Time: 2.34 seconds

INPUT/OUTPUT
──────────────────────────────────────────────────────────────────────
Input File:  contacts.xlsx
Output File: contacts_BCC_cleaned.csv
Initial Rows: 500
Final Rows:   450

DATA CLEANING SUMMARY
──────────────────────────────────────────────────────────────────────
✓ Removed 35 exact duplicate(s)
✓ Removed 15 fuzzy duplicate(s) (similar addresses/names)
✓ Parsed 0 combined address field(s)
✓ Normalized 127 state abbreviation(s)
✓ Formatted 89 ZIP code(s)
✓ Formatted 234 phone number(s)
✓ Identified 12 business name(s)

VALIDATION RESULTS
──────────────────────────────────────────────────────────────────────
⚠ 5 record(s) with incomplete addresses (flagged in Notes)
⚠ 3 invalid email address(es) (flagged in Notes)

WARNINGS (8 total)
──────────────────────────────────────────────────────────────────────
  • Row 45: Incomplete address (missing required fields)
  • Row 78: Invalid email: john@example
  • Row 92: Incomplete address (missing required fields)
  ...

======================================================================
✓ File ready for BCC Bulk Mailer import!
======================================================================
```

## 🔧 How It Works

The script performs 5 automated steps:

1. **Standardize Column Names** - Maps 70+ variations to BCC standard fields
2. **Parse Combined Data** - Uses usaddress library + regex to extract separate fields
3. **Clean & Validate** - Normalizes states, validates ZIPs/emails/phones, detects businesses
4. **Remove Duplicates** - Exact matching + fuzzy matching (catches typos/variations)
5. **Format for BCC** - Arranges columns and generates detailed report

## 💡 Handling Messy Data

### Combined Address in One Cell

**Input Excel:**
```
Column A
-----------------
John Doe, 123 Main St, Boston, MA 02101
Jane Smith, 456 Oak Ave, Apt 2B, Austin, TX 78701, (555) 123-4567, jane@email.com
Acme Corp LLC, 789 Business Blvd, Chicago, IL 60601
```

**Output CSV:**
```csv
FirstName,LastName,FullName,Company,Address1,Address2,City,State,ZIP,ZIP4,Phone,Email,Notes
John,Doe,John Doe,,123 Main St,,Boston,MA,02101,,,
Jane,Smith,Jane Smith,,456 Oak Ave,Apt 2B,Austin,TX,78701,,(555) 123-4567,jane@email.com,
,,,Acme Corp LLC,789 Business Blvd,,Chicago,IL,60601,,,
```

### Fuzzy Duplicate Detection

The script catches duplicates that exact matching would miss:

```
Row 1: John Smith, 123 Main Street, Boston, MA 02101
Row 2: Jon Smith, 123 Main St, Boston, MA 02101  ← Caught as duplicate!
Row 3: John Smith, 123 Main St., Boston MA 02101  ← Caught as duplicate!
```

### Phone Number Normalization

```
Input:                     Output:
555-123-4567         →    (555) 123-4567
(555) 123-4567       →    (555) 123-4567
555.123.4567         →    (555) 123-4567
15551234567          →    (555) 123-4567
```

### Business Name Detection

```
Input (FullName):          Output:
Acme Corp LLC        →    Company: "Acme Corp LLC", FullName: (empty)
Smith & Associates   →    Company: "Smith & Associates", FullName: (empty)
John Smith           →    FullName: "John Smith"
```

### Name Format Handling

```
Input:                     Output:
"Doe, John"          →    FirstName: "John", LastName: "Doe"
"Dr. John Smith Jr." →    FirstName: "John", LastName: "Smith"
"John Q. Doe"        →    FirstName: "John", LastName: "Doe"
```

## 🎯 Real-World Example

**Messy Input Excel (various formats):**
```
Name/Address                                           Phone          Email
-----------------------------------------------------------------------
John Smith, 123 main street, boston, massachusetts     555-123-4567   john@example.com
JANE DOE, 456 ELM AVE, APT 3, SPRINGFIELD, IL 62704   (555)234-5678
Acme Corp Inc, 789 Oak Blvd, Chicago Illinois 60601                  info@acme.com
Jon Smith, 123 Main St, Boston, MA 02101              5551234567
```

**Clean Output CSV:**
```csv
FirstName,LastName,FullName,Company,Address1,Address2,City,State,ZIP,ZIP4,Phone,Email,Notes
John,Smith,John Smith,,123 Main Street,,Boston,MA,02101,,(555) 123-4567,john@example.com,
Jane,Doe,Jane Doe,,456 Elm Ave,Apt 3,Springfield,IL,62704,,(555) 234-5678,,
,,,Acme Corp Inc,789 Oak Blvd,,Chicago,IL,60601,,,info@acme.com,
```

**Processing Notes:**
- Removed 1 fuzzy duplicate (Jon Smith vs John Smith at same address)
- Normalized states (massachusetts → MA, Illinois → IL)
- Proper capitalization (main street → Main Street, ELM AVE → Elm Ave)
- Detected and separated business name (Acme Corp Inc)
- Formatted all phone numbers to (XXX) XXX-XXXX
- Validated emails

## 📖 Column Name Recognition

The script automatically recognizes these column name patterns (case-insensitive):

### Name Fields
- **FirstName**: first_name, firstname, first, fname, given_name
- **LastName**: last_name, lastname, last, lname, surname, family_name
- **FullName**: name, full_name, fullname, contact, contact_name

### Business Fields
- **Company**: company, business, organization, org, company_name, business_name, firm

### Address Fields
- **Address1**: address, address1, street, street_address, addr, address_line_1
- **Address2**: address2, apt, suite, unit, address_line_2
- **City**: city, town, municipality
- **State**: state, province, region, st
- **ZIP**: zip, zipcode, zip_code, postal_code, postcode, zip5
- **ZIP4**: zip4, zip_4, plus4

### Contact Fields
- **Phone**: phone, phone_number, telephone, tel, mobile, cell, contact_number
- **Email**: email, email_address, e-mail, e_mail
- **Notes**: notes, note, comments, comment, remarks, memo

## 🎯 Importing into BCC Bulk Mailer

1. Run the script on your Excel file
2. Open BCC Bulk Mailer Business/Professional
3. Use the import function - BCC will auto-detect headers
4. Field mapping happens automatically due to standard column names
5. Proceed with CASS validation and presort
6. Review the generated report for any warnings

## 🔍 Advanced Features

### Fuzzy Duplicate Matching

The fuzzy matching algorithm uses the Levenshtein distance to calculate similarity between records. Default threshold is 85% (highly recommended).

**Threshold Guide:**
- **95-100%**: Very strict - only catches near-identical duplicates
- **85-94%**: Recommended - good balance
- **75-84%**: Lenient - may catch some false positives
- **Below 75%**: Not recommended - many false positives

### Business Name Detection

Automatically detects business entities by looking for:
- Legal suffixes: LLC, Inc, Corp, Ltd, LP, LLP, PC, PA, PLLC
- Business keywords: Shop, Store, Market, Center, Services, Group, Associates, Solutions

### Email Validation

Uses regex pattern to validate email format:
- Must have @ symbol
- Must have domain with at least 2-letter TLD
- Flags invalid emails in Notes column
- Converts all emails to lowercase

### Phone Number Parsing

Handles these formats:
- (XXX) XXX-XXXX
- XXX-XXX-XXXX
- XXX.XXX.XXXX
- XXXXXXXXXX
- 1-XXX-XXX-XXXX (strips leading 1)

## 📝 Data Validation

- **ZIP Code Format**: Validates against `^\d{5}(-\d{4})?$` pattern
- **State Abbreviations**: Converts to USPS standard 2-letter codes (all 50 states + territories)
- **Address Capitalization**: Title case with exceptions (PO, NW, APT, STE, etc.)
- **Duplicate Detection**: Case-insensitive, whitespace-trimmed + fuzzy matching
- **Incomplete Address Flagging**: Marks records missing Address1, City, State, or ZIP

Incomplete or invalid records are marked in the **Notes** column with flags like:
- `INCOMPLETE_ADDRESS` - Missing required address fields
- `INVALID_EMAIL` - Email format is invalid

## 🛠️ For Your Team (25 People)

### Best Practices

1. **Standardize Input**: Ask data sources to use consistent column names when possible
2. **Review Reports**: Always check the generated .txt report for warnings
3. **Adjust Fuzzy Threshold**: Start with 85%, increase if too many false positives
4. **Check Incomplete Records**: Filter Notes column for "INCOMPLETE_ADDRESS" before import
5. **Save Original Files**: Keep a copy of the original Excel file

### Batch Processing Multiple Files

```bash
# Process all Excel files in a folder (Windows)
for %f in (*.xlsx) do python clean_excel_data.py "%f"

# Process all Excel files in a folder (Mac/Linux)
for file in *.xlsx; do python clean_excel_data.py "$file"; done
```

### Common Team Workflows

**Scenario 1: Event Registration List**
```bash
python clean_excel_data.py event_registrations.xlsx -o mailing_list.csv
# Review mailing_list.txt report
# Import mailing_list.csv into BCC
```

**Scenario 2: Customer Database Export**
```bash
# More strict fuzzy matching for customer data
python clean_excel_data.py customers.xlsx --fuzzy-threshold 90
```

**Scenario 3: Mixed Personal + Business Contacts**
```bash
# Script automatically separates businesses into Company column
python clean_excel_data.py mixed_contacts.xlsx
# Check report for how many businesses were detected
```

## ⚠️ Troubleshooting

### "Missing columns after standardization"
Your Excel file has column names the script doesn't recognize. Check the column name recognition table above or modify the COLUMN_MAPPING in the script.

### "Warning: usaddress library not available"
Run `pip install usaddress`. The script will work without it but use less accurate parsing.

### "Warning: rapidfuzz library not available"
Run `pip install rapidfuzz`. The script will work but won't perform fuzzy duplicate matching.

### "Cannot remove duplicates - no address/name columns found"
The data format is very unusual. The script will still output data but cannot deduplicate.

### "Warning: X records have incomplete address data"
These records are flagged in the Notes column. Review them manually - they may fail CASS validation in BCC.

### High number of fuzzy duplicates
Your threshold might be too low. Try increasing it: `--fuzzy-threshold 90`

### Empty output file
- Check that your Excel file has data in the first sheet
- Verify at least one column contains address information
- Look for error messages in the console output

## 📦 Requirements

- Python 3.7+
- pandas >= 2.0.0
- openpyxl >= 3.1.0
- usaddress >= 1.0.0
- rapidfuzz >= 3.0.0

## 🆕 What's New in Enhanced Edition

- ✨ usaddress library integration for professional address parsing
- ✨ Fuzzy duplicate matching (catches typos and variations)
- ✨ Detailed processing reports saved automatically
- ✨ Phone number parsing and formatting
- ✨ Email validation with regex
- ✨ Business name detection and separation
- ✨ Comprehensive statistics tracking
- ✨ Warning system for data quality issues
- ✨ Command-line options for fuzzy threshold control

## 📄 License

MIT License

## 🤝 Support

For issues or questions:
1. Check the troubleshooting section above
2. Review the generated .txt report for specific warnings
3. Ensure all dependencies are installed: `pip install -r requirements.txt`
