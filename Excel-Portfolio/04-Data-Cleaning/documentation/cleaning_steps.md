# Data Cleaning Steps

## Overview

This document provides a detailed, step-by-step guide to the data cleaning process used in this project.

## Phase 1: Initial Assessment

### Step 1.1: Data Profiling

Before cleaning, assess the current state of the data:

- **Row Count**: Total number of records
- **Column Count**: Number of fields
- **Data Types**: Identify expected vs. actual types
- **Completeness**: Calculate % of missing values per column
- **Uniqueness**: Check for duplicates

### Step 1.2: Issue Identification

Common issues to look for:

| Issue Type | How to Detect | Impact |
|------------|--------------|--------|
| Extra spaces | =LEN(A1) vs =LEN(TRIM(A1)) | Matching errors |
| Inconsistent case | Manual review | Grouping errors |
| Wrong data types | ISNUMBER, ISTEXT functions | Calculation errors |
| Duplicates | Conditional formatting | Inflated counts |
| Invalid formats | Error values, #VALUE! | Processing failures |

---

## Phase 2: Preparation

### Step 2.1: Create Backup

```
1. Save original file as "raw_data_BACKUP.xlsx"
2. Never modify the original backup
3. Work on a copy named "data_cleaning_WORK.xlsx"
```

### Step 2.2: Document Starting Point

Record initial metrics:
- Total rows: _______
- Missing values by column: _______
- Known issues: _______

---

## Phase 3: Cleaning Execution

### Step 3.1: Remove Extra Spaces

**Formula**:
```excel
=TRIM(A2)
```

**Process**:
1. Insert new column next to text column
2. Apply TRIM formula
3. Copy results as values
4. Replace original column
5. Delete helper column

### Step 3.2: Remove Non-Printable Characters

**Formula**:
```excel
=CLEAN(A2)
```

**Use Case**: Data imported from mainframes or web sources

### Step 3.3: Standardize Text Case

**Formulas**:
```excel
=PROPER(A2)     # Proper Case (First Letter Capitalized)
=UPPER(A2)      # ALL CAPS
=LOWER(A2)      # all lowercase
```

**Application**:
- Names: PROPER
- Email addresses: LOWER
- Country codes: UPPER

### Step 3.4: Fix Date Formats

**Identify Issues**:
```excel
=ISDATE(A2)     # Returns TRUE if valid date
```

**Convert Text to Date**:
```excel
=DATEVALUE(A2)
=DATE(MID(A2,7,4), LEFT(A2,2), MID(A2,4,2))  # For MMDDYYYY format
```

### Step 3.5: Convert Text Numbers to Actual Numbers

**Formula**:
```excel
=VALUE(A2)
```

**Alternative Methods**:
- Text to Columns feature
- Multiply by 1: `=A2*1`
- Paste Special → Multiply

### Step 3.6: Handle Missing Values

**Options**:

1. **Remove Rows** (if < 5% missing):
   - Filter blanks
   - Delete rows

2. **Fill with Default**:
   ```excel
   =IF(ISBLANK(A2), "Unknown", A2)
   ```

3. **Fill with Average/Median**:
   ```excel
   =IF(ISBLANK(A2), AVERAGE($A$2:$A$1000), A2)
   ```

4. **Forward Fill**:
   - Sort by date
   - Use formula to reference previous non-blank

### Step 3.7: Remove Duplicates

**Method 1 - Excel Feature**:
```
1. Select data range
2. Data tab → Remove Duplicates
3. Choose columns to check
4. Click OK
```

**Method 2 - Formula Identification**:
```excel
=COUNTIF($A$2:A2, A2)>1  # Marks duplicates after first occurrence
```

### Step 3.8: Standardize Categories

**Problem**: "USA", "U.S.A", "United States", "US" all mean same thing

**Solution**:
```excel
=SWITCH(A2,
    "USA", "United States",
    "U.S.A", "United States",
    "US", "United States",
    "America", "United States",
    A2)
```

---

## Phase 4: Validation

### Step 4.1: Data Type Verification

Check each column:
```excel
=ISNUMBER(A2)    # Should be TRUE for numeric columns
=ISTEXT(A2)      # Should be TRUE for text columns
=ISLOGICAL(A2)   # Should be TRUE for boolean columns
```

### Step 4.2: Range Checks

Verify values are within expected ranges:
```excel
=AND(A2>=MIN_EXPECTED, A2<=MAX_EXPECTED)
```

### Step 4.3: Consistency Checks

Cross-validate related fields:
```excel
=IF(EndDate < StartDate, "ERROR", "OK")
```

### Step 4.4: Completeness Check

Final missing value count:
```excel
=COUNTBLANK(Range)
```

---

## Phase 5: Documentation

### Step 5.1: Record Transformations

| Step | Action | Formula Used | Rows Affected |
|------|--------|--------------|---------------|
| 1 | Trimmed spaces | =TRIM() | 1,234 |
| 2 | Fixed dates | =DATEVALUE() | 567 |
| 3 | Removed duplicates | Built-in tool | 89 |

### Step 5.2: Save Final Version

```
Filename: cleaned_final_data_YYYYMMDD.xlsx
Location: /data/ folder
Documentation: Update this file
```

---

## Quality Checklist

- [ ] All text trimmed and cleaned
- [ ] Consistent text case applied
- [ ] Dates in proper format
- [ ] Numbers stored as numbers, not text
- [ ] Missing values handled appropriately
- [ ] Duplicates removed
- [ ] Categories standardized
- [ ] Data validation rules applied
- [ ] No error values (#N/A, #VALUE!, etc.)
- [ ] Final row count documented

---

[Back to Project](../README.md)
