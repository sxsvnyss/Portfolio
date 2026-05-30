# Excel Formulas Used in Data Cleaning

## Text Functions

### 1. TRIM - Remove Extra Spaces

**Purpose**: Removes leading, trailing, and extra spaces between words.

```excel
=TRIM(text)
```

**Example**:
```excel
=TRIM("  Hello   World  ")    # Returns: "Hello World"
```

**Use Case**: Clean user input, imported data with inconsistent spacing.

---

### 2. CLEAN - Remove Non-Printable Characters

**Purpose**: Removes characters that can't be printed (ASCII 0-31).

```excel
=CLEAN(text)
```

**Example**:
```excel
=CLEAN(A2 & CHAR(10))    # Removes line feed character
```

**Use Case**: Data from mainframes, web scrapes, or legacy systems.

---

### 3. SUBSTITUTE - Replace Text

**Purpose**: Replaces specific text strings.

```excel
=SUBSTITUTE(text, old_text, new_text, [instance_num])
```

**Examples**:
```excel
=SUBSTITUTE(A2, "-", "/")                    # Replace all hyphens with slashes
=SUBSTITUTE(A2, "USA", "United States")      # Standardize country names
=SUBSTITUTE(A2, CHAR(160), " ")              # Replace non-breaking spaces
```

**Use Case**: Standardizing formats, fixing encoding issues.

---

### 4. TEXT - Format Numbers and Dates

**Purpose**: Converts values to text in a specific format.

```excel
=TEXT(value, format_code)
```

**Examples**:
```excel
=TEXT(A2, "mm/dd/yyyy")     # Date format
=TEXT(A2, "$#,##0.00")      # Currency format
=TEXT(A2, "00000")          # Pad with zeros (ZIP codes)
```

**Use Case**: Creating consistent display formats.

---

### 5. VALUE - Convert Text to Number

**Purpose**: Converts text representations of numbers to actual numbers.

```excel
=VALUE(text)
```

**Example**:
```excel
=VALUE("$1,234.56")    # Returns: 1234.56
```

**Use Case**: Numbers stored as text from imports.

---

### 6. PROPER/UPPER/LOWER - Case Conversion

**Purpose**: Change text case.

```excel
=PROPER(text)    # First letter of each word capitalized
=UPPER(text)     # All uppercase
=LOWER(text)     # All lowercase
```

**Examples**:
```excel
=PROPER("john doe")      # Returns: "John Doe"
=LOWER("EMAIL@EXAMPLE.COM")  # Returns: "email@example.com"
=UPPER("us")             # Returns: "US"
```

**Use Case**: Standardizing names, emails, codes.

---

### 7. CONCATENATE / & - Join Text

**Purpose**: Combine text from multiple cells.

```excel
=CONCATENATE(text1, text2, ...)
=text1 & text2 & ...
```

**Example**:
```excel
=A2 & " " & B2           # Join first and last name
=A2 & ", " & B2 & ", " & C2   # Create address string
```

**Use Case**: Creating full names, addresses, unique IDs.

---

### 8. LEFT/RIGHT/MID - Extract Text

**Purpose**: Extract portions of text strings.

```excel
=LEFT(text, num_chars)
=RIGHT(text, num_chars)
=MID(text, start_num, num_chars)
```

**Examples**:
```excel
=LEFT(A2, 3)             # First 3 characters
=RIGHT(A2, 4)            # Last 4 characters
=MID(A2, 4, 5)           # 5 characters starting at position 4
```

**Use Case**: Extracting area codes, postal codes, ID prefixes.

---

### 9. FIND/SEARCH - Locate Text

**Purpose**: Find position of text within a string.

```excel
=FIND(find_text, within_text, [start_num])    # Case-sensitive
=SEARCH(find_text, within_text, [start_num])  # Not case-sensitive
```

**Example**:
```excel
=FIND("@", A2)           # Find @ symbol in email
=SEARCH("inc", A2)       # Find "inc" in company name
```

**Use Case**: Validating emails, parsing structured text.

---

## Logical Functions

### 10. IFERROR - Handle Errors

**Purpose**: Return alternative value if formula results in error.

```excel
=IFERROR(value, value_if_error)
```

**Example**:
```excel
=IFERROR(A2/B2, 0)       # Return 0 if division by zero
=IFERROR(VLOOKUP(...), "Not Found")
```

**Use Case**: Preventing error values in cleaned data.

---

### 11. IF/IFS - Conditional Logic

**Purpose**: Return different values based on conditions.

```excel
=IF(logical_test, value_if_true, [value_if_false])
=IFS(test1, value1, test2, value2, ...)
```

**Examples**:
```excel
=IF(A2="", "Missing", A2)                    # Handle blanks
=IF(ISNUMBER(A2), A2, "Invalid")             # Validate numbers
=IFS(A2<60,"F", A2<70,"D", A2<80,"C", A2<90,"B", TRUE,"A")
```

**Use Case**: Categorization, validation, default values.

---

### 12. SWITCH - Multiple Value Matching

**Purpose**: Match against multiple values and return corresponding result.

```excel
=SWITCH(expression, value1, result1, value2, result2, ..., [default])
```

**Example**:
```excel
=SWITCH(A2,
    "CA", "California",
    "NY", "New York",
    "TX", "Texas",
    "Other")
```

**Use Case**: Standardizing abbreviations, category mapping.

---

## Information Functions

### 13. ISBLANK - Check for Empty Cells

**Purpose**: Test if a cell is empty.

```excel
=ISBLANK(value)
```

**Example**:
```excel
=IF(ISBLANK(A2), "Missing", "Present")
```

**Use Case**: Identifying missing data.

---

### 14. ISNUMBER/ISTEXT - Check Data Types

**Purpose**: Verify data type.

```excel
=ISNUMBER(value)
=ISTEXT(value)
```

**Example**:
```excel
=IF(ISNUMBER(A2), "Valid", "Needs Conversion")
```

**Use Case**: Data validation, quality checks.

---

### 15. LEN - Count Characters

**Purpose**: Count number of characters in text.

```excel
=LEN(text)
```

**Example**:
```excel
=LEN(A2)                 # Total length
=LEN(A2)-LEN(TRIM(A2))   # Count of extra spaces
```

**Use Case**: Validating field lengths, detecting issues.

---

## Date Functions

### 16. DATEVALUE - Convert Text to Date

**Purpose**: Convert date stored as text to serial number.

```excel
=DATEVALUE(date_text)
```

**Example**:
```excel
=DATEVALUE("01/15/2024")
=DATEVALUE(A2)
```

**Use Case**: Fixing dates imported as text.

---

### 17. DATE - Create Date from Components

**Purpose**: Build date from year, month, day.

```excel
=DATE(year, month, day)
```

**Example**:
```excel
=DATE(2024, 1, 15)
=DATE(MID(A2,7,4), LEFT(A2,2), MID(A2,4,2))   # Parse MMDDYYYY
```

**Use Case**: Reconstructing dates from split columns.

---

### 18. TODAY/NOW - Current Date/Time

**Purpose**: Get current date or date/time.

```excel
=TODAY()
=NOW()
```

**Use Case**: Age calculations, freshness checks.

---

## Lookup Functions

### 19. XLOOKUP - Modern Lookup

**Purpose**: Search for value and return corresponding result.

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])
```

**Example**:
```excel
=XLOOKUP(A2, Codes!A:A, Codes!B:B, "Not Found")
```

**Use Case**: Standardizing values using reference tables.

---

### 20. VLOOKUP - Vertical Lookup

**Purpose**: Look up value in first column, return from specified column.

```excel
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

**Example**:
```excel
=VLOOKUP(A2, ReferenceTable, 2, FALSE)
```

**Use Case**: Legacy compatibility, simple lookups.

---

## Combination Examples

### Email Validation
```excel
=IF(AND(ISNUMBER(FIND("@",A2)), ISNUMBER(FIND(".",A2))), "Valid", "Invalid")
```

### Phone Number Formatting
```excel
=TEXT(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(A2,"-",""),"(",""),")",""), "(000) 000-0000")
```

### Full Name Construction
```excel
=TRIM(PROPER(IF(ISBLANK(B2), A2, A2 & " " & B2)))
```

### Duplicate Detection
```excel
=IF(COUNTIF($A$2:A2, A2)>1, "Duplicate", "Unique")
```

---

[Back to Project](../README.md)
