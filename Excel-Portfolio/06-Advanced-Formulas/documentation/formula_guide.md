# Advanced Excel Formula Guide

## Overview

This comprehensive guide explains all the advanced formulas used in this collection, with examples and use cases.

---

## 1. Dynamic Array Functions (Excel 365)

### FILTER Function

**Purpose**: Filter data based on criteria.

```excel
=FILTER(array, include, [if_empty])
```

**Examples**:

```excel
// Filter sales > 1000
=FILTER(A2:C100, C2:C100 > 1000, "No results")

// Multiple criteria (AND)
=FILTER(A2:C100, (C2:C100 > 1000) * (B2:B100 = "West"), "No results")

// Multiple criteria (OR)
=FILTER(A2:C100, (C2:C100 > 1000) + (C2:C100 < 100), "No results")
```

---

### SORT Function

**Purpose**: Sort data dynamically.

```excel
=SORT(array, [sort_index], [sort_order], [by_col])
```

**Examples**:

```excel
// Sort by column 3 ascending
=SORT(A2:C100, 3, 1)

// Sort by column 3 descending
=SORT(A2:C100, 3, -1)

// Multi-level sort (col 3 desc, then col 2 asc)
=SORT(A2:C100, {3,2}, {-1,1})

// Sort filtered results
=SORT(FILTER(A2:C100, C2:C100 > 1000), 3, -1)
```

---

### UNIQUE Function

**Purpose**: Extract unique values.

```excel
=UNIQUE(array, [by_col], [exactly_once])
```

**Examples**:

```excel
// Unique values from a column
=UNIQUE(A2:A100)

// Unique rows from multiple columns
=UNIQUE(A2:C100)

// Values that appear exactly once
=UNIQUE(A2:A100, FALSE, TRUE)
```

---

### SEQUENCE Function

**Purpose**: Generate number sequences.

```excel
=SEQUENCE(rows, [columns], [start], [step])
```

**Examples**:

```excel
// Numbers 1 to 10
=SEQUENCE(10)

// 5 rows × 3 columns
=SEQUENCE(5, 3)

// From 100 to 91 (descending)
=SEQUENCE(10, 1, 100, -1)

// Create ID numbers
="ID-" & TEXT(SEQUENCE(100), "0000")
```

---

### RANDARRAY Function

**Purpose**: Generate random numbers.

```excel
=RANDARRAY(rows, [columns], [min], [max], [whole_number])
```

**Examples**:

```excel
// Random decimals between 0 and 1
=RANDARRAY(10, 5)

// Random integers 1-100
=RANDARRAY(10, 1, 1, 100, TRUE)

// Random dates in 2024
=DATE(2024,1,1) + RANDARRAY(10, 1, 0, 364, TRUE)
```

---

## 2. XLOOKUP Function

**Purpose**: Modern replacement for VLOOKUP/HLOOKUP.

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

**Examples**:

```excel
// Basic lookup
=XLOOKUP(A2, ProductIDs, ProductNames, "Not found")

// Multiple criteria lookup
=XLOOKUP(1, (A2:A100=A2) * (B2:B100=B2), C2:C100, "Not found")

// Return entire row
=XLOOKUP(A2, IDs, B2:D100, "Not found")

// Last match (search from bottom)
=XLOOKUP(A2, LookupCol, ReturnCol, , , -1)

// Wildcard match
=XLOOKUP("*" & A2 & "*", TextCol, ReturnCol, , 2)
```

---

## 3. INDEX/MATCH Combinations

**Purpose**: Flexible lookups (legacy alternative to XLOOKUP).

```excel
=INDEX(return_range, MATCH(lookup_value, lookup_range, 0))
```

**Examples**:

```excel
// Basic INDEX/MATCH
=INDEX(C2:C100, MATCH(A2, B2:B100, 0))

// Two-way lookup (matrix)
=INDEX(B2:D100, MATCH(G2, A2:A100, 0), MATCH(H2, B1:D1, 0))

// Multiple criteria
=INDEX(D2:D100, MATCH(1, (A2:A100=F2) * (B2:B100=G2), 0))
```

---

## 4. LET Function

**Purpose**: Define variables within formulas for readability.

```excel
=LET(name1, value1, name2, value2, ..., calculation)
```

**Examples**:

```excel
// Calculate bonus with named variables
=LET(
    sales, B2,
    target, 10000,
    rate, 0.05,
    bonus, MAX(0, sales - target) * rate,
    bonus
)

// Complex calculation with intermediate steps
=LET(
    data, FILTER(A2:C100, C2:C100 > 1000),
    sorted, SORT(data, 3, -1),
    top10, TAKE(sorted, 10),
    AVERAGE(INDEX(top10, , 3))
)
```

---

## 5. LAMBDA Function

**Purpose**: Create custom reusable functions.

```excel
=LAMBDA(parameter1, parameter2, ..., calculation)
```

**Examples**:

```excel
// Simple average function
=LAMBDA(x, y, (x + y) / 2)(A2, B2)

// Tax calculator (define once, use everywhere)
=LAMBDA(amount, IF(amount > 10000, amount * 0.25, amount * 0.15))(B2)

// Named LAMBDA (store in Name Manager)
// Name: AreaCircle
// Refers to: =LAMBDA(r, PI() * r^2)
// Use: =AreaCircle(A2)
```

---

## 6. Logical Functions

### IFS Function

```excel
=IFS(test1, value1, test2, value2, ...)
```

```excel
=IFS(
    A2 >= 90, "A",
    A2 >= 80, "B",
    A2 >= 70, "C",
    A2 >= 60, "D",
    TRUE, "F"
)
```

### SWITCH Function

```excel
=SWITCH(expression, value1, result1, value2, result2, ..., default)
```

```excel
=SWITCH(A2,
    "CA", "California",
    "NY", "New York",
    "TX", "Texas",
    "Other State"
)
```

### Nested IF with AND/OR

```excel
=IF(AND(A2 > 100, B2 = "Yes"), "Approved", 
    IF(OR(A2 > 50, C2 = "VIP"), "Review", "Denied"))
```

---

## 7. Text Functions

### TEXTJOIN

```excel
=TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
```

```excel
// Combine names with comma
=TEXTJOIN(", ", TRUE, A2:A100)

// Create email list
=TEXTJOIN("; ", TRUE, B2:B100 & "@company.com")
```

### CONCAT vs CONCATENATE

```excel
=CONCAT(A2, " ", B2)     // Modern
=A2 & " " & B2           // Alternative
```

### TEXTSPLIT (Excel 365)

```excel
=TEXTSPLIT(text, col_delimiter, [row_delimiter])
```

```excel
// Split "John,Doe,Manager" into 3 columns
=TEXTSPLIT(A2, ",")
```

---

## 8. Date Functions

### WORKDAY and NETWORKDAYS

```excel
=WORKDAY(start_date, days, [holidays])
=NETWORKDAYS(start_date, end_date, [holidays])
```

```excel
// Project end date (30 business days)
=WORKDAY(TODAY(), 30, HolidayRange)

// Business days between dates
=NETWORKDAYS(StartDate, EndDate, HolidayRange)
```

### EOMONTH

```excel
=EOMONTH(start_date, months)
```

```excel
// End of current month
=EOMONTH(TODAY(), 0)

// End of next quarter
=EOMONTH(TODAY(), 3 - MOD(MONTH(TODAY())-1, 3))
```

### DATEDIF

```excel
=DATEDIF(start_date, end_date, unit)
```

```excel
// Years between dates
=DATEDIF(A2, B2, "Y")

// Complete months
=DATEDIF(A2, B2, "M")

// Days ignoring years (anniversary)
=DATEDIF(A2, B2, "YD")
```

---

## 9. Aggregation Functions

### SUMIFS / COUNTIFS / AVERAGEIFS

```excel
=SUMIFS(sum_range, criteria_range1, criteria1, ...)
=COUNTIFS(criteria_range1, criteria1, ...)
=AVERAGEIFS(avg_range, criteria_range1, criteria1, ...)
```

```excel
// Sum sales for West region
=SUMIFS(SalesAmount, Region, "West")

// Count orders > $1000
=COUNTIFS(OrderAmount, ">1000")

// Average by product and region
=AVERAGEIFS(Sales, Product, "Widget", Region, "East")
```

### SUBTOTAL

```excel
=SUBTOTAL(function_num, ref1, ...)
```

```excel
// Sum visible rows only (ignores filtered out)
=SUBTOTAL(109, C2:C100)

// Count visible rows
=SUBTOTAL(103, A2:A100)
```

---

## 10. Financial Functions

### NPV and XNPV

```excel
=NPV(rate, value1, [value2], ...)
=XNPV(rate, values, dates)
```

```excel
// Net present value
=NPV(0.1, B2:B10) + B1  // B1 is initial investment (negative)

// XNPV for irregular cash flows
=XNPV(0.1, CashFlows, Dates)
```

### PMT

```excel
=PMT(rate, nper, pv, [fv], [type])
```

```excel
// Monthly loan payment
=PMT(0.05/12, 360, -200000)
```

---

## 11. Array Formula Patterns

### Legacy CSE Array Formulas

```excel
{=SUM(IF(A2:A100 > 100, B2:B100, 0))}  // Enter with Ctrl+Shift+Enter
```

### Modern Dynamic Arrays

```excel
=SUM(FILTER(B2:B100, A2:A100 > 100, 0))  // No CSE needed
```

### SEQUENCE for Row Numbers

```excel
=SEQUENCE(COUNTA(A2:A100))  // Dynamic row numbers
```

---

## Best Practices

1. **Use LET** for complex formulas to improve readability
2. **Prefer XLOOKUP** over VLOOKUP for flexibility
3. **Use dynamic arrays** when available (Excel 365)
4. **Name ranges** for better formula documentation
5. **Avoid volatile functions** (OFFSET, INDIRECT) in large datasets
6. **Test with small data** before applying to large ranges
7. **Document complex formulas** with comments or helper cells

---

[Back to Project](../README.md)
