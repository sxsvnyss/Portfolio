# Excel Functions Reference Guide

## Comprehensive Function Catalog by Category

---

## Text Functions

| Function | Description | Syntax | Example |
|----------|-------------|--------|---------|
| **TRIM** | Remove extra spaces | `TRIM(text)` | `=TRIM(" Hello ")` → "Hello" |
| **CLEAN** | Remove non-printable chars | `CLEAN(text)` | `=CLEAN(A1)` |
| **CONCAT** | Join text strings | `CONCAT(text1, text2, ...)` | `=CONCAT(A1," ",B1)` |
| **TEXTJOIN** | Join with delimiter | `TEXTJOIN(delimiter, ignore_empty, text1, ...)` | `=TEXTJOIN(", ",TRUE,A1:A10)` |
| **LEFT** | Extract from left | `LEFT(text, num_chars)` | `=LEFT("Excel",2)` → "Ex" |
| **RIGHT** | Extract from right | `RIGHT(text, num_chars)` | `=RIGHT("Excel",3)` → "cel" |
| **MID** | Extract from middle | `MID(text, start, num_chars)` | `=MID("Excel",2,3)` → "xce" |
| **LEN** | Count characters | `LEN(text)` | `=LEN("Excel")` → 5 |
| **FIND** | Find text (case-sensitive) | `FIND(find_text, within_text)` | `=FIND("c","Excel")` → 3 |
| **SEARCH** | Find text (not case-sensitive) | `SEARCH(find_text, within_text)` | `=SEARCH("C","Excel")` → 3 |
| **SUBSTITUTE** | Replace text | `SUBSTITUTE(text, old, new)` | `=SUBSTITUTE(A1,"old","new")` |
| **REPLACE** | Replace by position | `REPLACE(old_text, start, num, new)` | `=REPLACE("Excel",1,1,"E")` |
| **UPPER** | Convert to uppercase | `UPPER(text)` | `=UPPER("excel")` → "EXCEL" |
| **LOWER** | Convert to lowercase | `LOWER(text)` | `=LOWER("EXCEL")` → "excel" |
| **PROPER** | Capitalize first letters | `PROPER(text)` | `=PROPER("john doe")` → "John Doe" |
| **TEXT** | Format as text | `TEXT(value, format_code)` | `=TEXT(TODAY(),"mm/dd/yyyy")` |
| **VALUE** | Convert text to number | `VALUE(text)` | `=VALUE("$100")` → 100 |
| **TEXTSPLIT** | Split text (365) | `TEXTSPLIT(text, col_delim)` | `=TEXTSPLIT("A,B,C",",")` |

---

## Logical Functions

| Function | Description | Syntax | Example |
|----------|-------------|--------|---------|
| **IF** | Conditional test | `IF(logical_test, value_if_true, value_if_false)` | `=IF(A1>10,"Yes","No")` |
| **IFS** | Multiple conditions (2019+) | `IFS(test1, value1, test2, value2, ...)` | `=IFS(A1>90,"A",A1>80,"B",TRUE,"C")` |
| **SWITCH** | Match against values | `SWITCH(expression, val1, result1, ..., default)` | `=SWITCH(A1,1,"One",2,"Two","Other")` |
| **AND** | All conditions true | `AND(logical1, logical2, ...)` | `=AND(A1>0,A1<100)` |
| **OR** | Any condition true | `OR(logical1, logical2, ...)` | `=OR(A1="Yes",A1="Maybe")` |
| **NOT** | Reverse logic | `NOT(logical)` | `=NOT(A1=TRUE)` |
| **IFERROR** | Handle errors | `IFERROR(value, value_if_error)` | `=IFERROR(A1/B1,0)` |
| **IFNA** | Handle #N/A only | `IFNA(value, value_if_na)` | `=IFNA(VLOOKUP(...),"Not found")` |
| **XOR** | Exclusive OR | `XOR(logical1, logical2, ...)` | `=XOR(TRUE,FALSE)` → TRUE |

---

## Lookup & Reference Functions

| Function | Description | Syntax | Example |
|----------|-------------|--------|---------|
| **XLOOKUP** | Modern lookup (365) | `XLOOKUP(lookup, lookup_array, return_array, [if_not_found])` | `=XLOOKUP(A1,B:B,C:C,"Not found")` |
| **VLOOKUP** | Vertical lookup | `VLOOKUP(lookup_value, table, col_index, [range_lookup])` | `=VLOOKUP(A1,B2:D100,3,FALSE)` |
| **HLOOKUP** | Horizontal lookup | `HLOOKUP(lookup_value, table, row_index, [range_lookup])` | `=HLOOKUP(A1,B2:Z10,5,FALSE)` |
| **INDEX** | Return value at position | `INDEX(array, row_num, [col_num])` | `=INDEX(A1:C100,5,3)` |
| **MATCH** | Find position | `MATCH(lookup_value, lookup_array, [match_type])` | `=MATCH(A1,B:B,0)` |
| **OFFSET** | Reference offset | `OFFSET(reference, rows, cols, [height], [width])` | `=OFFSET(A1,5,0)` |
| **INDIRECT** | Text to reference | `INDIRECT(ref_text, [a1])` | `=INDIRECT("A"&ROW())` |
| **CHOOSE** | Choose from list | `CHOOSE(index_num, value1, value2, ...)` | `=CHOOSE(2,"A","B","C")` → "B" |
| **TRANSPOSE** | Flip orientation | `TRANSPOSE(array)` | `=TRANSPOSE(A1:C5)` |

---

## Math & Statistical Functions

| Function | Description | Syntax | Example |
|----------|-------------|--------|---------|
| **SUM** | Add numbers | `SUM(number1, number2, ...)` | `=SUM(A1:A100)` |
| **SUMIF** | Sum with one condition | `SUMIF(range, criteria, [sum_range])` | `=SUMIF(A:A,">100",B:B)` |
| **SUMIFS** | Sum with multiple conditions | `SUMIFS(sum_range, criteria_range1, criteria1, ...)` | `=SUMIFS(C:C,A:A,">100\",B:B,\"West\")` |
| **COUNT** | Count numbers | `COUNT(value1, value2, ...)` | `=COUNT(A1:A100)` |
| **COUNTA** | Count non-blank | `COUNTA(value1, value2, ...)` | `=COUNTA(A1:A100)` |
| **COUNTBLANK** | Count empty cells | `COUNTBLANK(range)` | `=COUNTBLANK(A1:A100)` |
| **COUNTIF** | Count with condition | `COUNTIF(range, criteria)` | `=COUNTIF(A:A,">100\")` |
| **COUNTIFS** | Count with multiple conditions | `COUNTIFS(criteria_range1, criteria1, ...)` | `=COUNTIFS(A:A,">100\",B:B,\"West\")` |
| **AVERAGE** | Calculate mean | `AVERAGE(number1, number2, ...)` | `=AVERAGE(A1:A100)` |
| **AVERAGEIF** | Average with condition | `AVERAGEIF(range, criteria, [average_range])` | `=AVERAGEIF(A:A,">100\",B:B)` |
| **AVERAGEIFS** | Average with multiple conditions | `AVERAGEIFS(avg_range, crit_range1, crit1, ...)` | `=AVERAGEIFS(C:C,A:A,">100\")` |
| **MIN** | Minimum value | `MIN(number1, number2, ...)` | `=MIN(A1:A100)` |
| **MAX** | Maximum value | `MAX(number1, number2, ...)` | `=MAX(A1:A100)` |
| **MEDIAN** | Middle value | `MEDIAN(number1, number2, ...)` | `=MEDIAN(A1:A100)` |
| **MODE.SNGL** | Most frequent value | `MODE.SNGL(number1, number2, ...)` | `=MODE.SNGL(A1:A100)` |
| **STDEV.S** | Standard deviation (sample) | `STDEV.S(number1, number2, ...)` | `=STDEV.S(A1:A100)` |
| **VAR.S** | Variance (sample) | `VAR.S(number1, number2, ...)` | `=VAR.S(A1:A100)` |
| **RANK.EQ** | Rank a number | `RANK.EQ(number, ref, [order])` | `=RANK.EQ(A1,$A$1:$A$100)` |
| **LARGE** | Nth largest | `LARGE(array, k)` | `=LARGE(A1:A100,3)` → 3rd largest |
| **SMALL** | Nth smallest | `SMALL(array, k)` | `=SMALL(A1:A100,3)` → 3rd smallest |

---

## Date & Time Functions

| Function | Description | Syntax | Example |
|----------|-------------|--------|---------|
| **TODAY** | Current date | `TODAY()` | `=TODAY()` |
| **NOW** | Current date/time | `NOW()` | `=NOW()` |
| **DATE** | Create date | `DATE(year, month, day)` | `=DATE(2024,1,15)` |
| **TIME** | Create time | `TIME(hour, minute, second)` | `=TIME(14,30,0)` |
| **YEAR** | Extract year | `YEAR(serial_number)` | `=YEAR(TODAY())` |
| **MONTH** | Extract month | `MONTH(serial_number)` | `=MONTH(TODAY())` |
| **DAY** | Extract day | `DAY(serial_number)` | `=DAY(TODAY())` |
| **HOUR** | Extract hour | `HOUR(serial_number)` | `=HOUR(NOW())` |
| **MINUTE** | Extract minute | `MINUTE(serial_number)` | `=MINUTE(NOW())` |
| **SECOND** | Extract second | `SECOND(serial_number)` | `=SECOND(NOW())` |
| **WEEKDAY** | Day of week | `WEEKDAY(serial_number, [return_type])` | `=WEEKDAY(TODAY())` |
| **WORKDAY** | Workdays from date | `WORKDAY(start_date, days, [holidays])` | `=WORKDAY(TODAY(),30)` |
| **NETWORKDAYS** | Workdays between dates | `NETWORKDAYS(start, end, [holidays])` | `=NETWORKDAYS(A1,B1)` |
| **EDATE** | Months from date | `EDATE(start_date, months)` | `=EDATE(TODAY(),3)` → 3 months later |
| **EOMONTH** | End of month | `EOMONTH(start_date, months)` | `=EOMONTH(TODAY(),0)` |
| **DATEDIF** | Difference between dates | `DATEDIF(start, end, unit)` | `=DATEDIF(A1,B1,"Y")` → years |
| **YEARFRAC** | Fraction of year | `YEARFRAC(start_date, end_date)` | `=YEARFRAC(A1,B1)` |

---

## Financial Functions

| Function | Description | Syntax | Example |
|----------|-------------|--------|---------|
| **PV** | Present value | `PV(rate, nper, pmt, [fv], [type])` | `=PV(0.05,10,-1000)` |
| **FV** | Future value | `FV(rate, nper, pmt, [pv], [type])` | `=FV(0.05,10,-1000)` |
| **PMT** | Payment amount | `PMT(rate, nper, pv, [fv], [type])` | `=PMT(0.05/12,360,-200000)` |
| **NPV** | Net present value | `NPV(rate, value1, value2, ...)` | `=NPV(0.1,B1:B10)` |
| **XNPV** | NPV for irregular cash flows | `XNPV(rate, values, dates)` | `=XNPV(0.1,B1:B10,A1:A10)` |
| **IRR** | Internal rate of return | `IRR(values, [guess])` | `=IRR(B1:B10)` |
| **XIRR** | IRR for irregular cash flows | `XIRR(values, dates, [guess])` | `=XIRR(B1:B10,A1:A10)` |
| **RATE** | Interest rate per period | `RATE(nper, pmt, pv, [fv], [type])` | `=RATE(360,-1000,200000)` |
| **NPER** | Number of periods | `NPER(rate, pmt, pv, [fv], [type])` | `=NPER(0.05,-1000,20000)` |
| **SLN** | Straight-line depreciation | `SLN(cost, salvage, life)` | `=SLN(10000,1000,5)` |

---

## Dynamic Array Functions (Excel 365)

| Function | Description | Syntax | Example |
|----------|-------------|--------|---------|
| **FILTER** | Filter data | `FILTER(array, include, [if_empty])` | `=FILTER(A2:C100,C2:C100>1000)` |
| **SORT** | Sort data | `SORT(array, [sort_index], [sort_order])` | `=SORT(A2:C100,3,-1)` |
| **UNIQUE** | Unique values | `UNIQUE(array, [by_col], [exactly_once])` | `=UNIQUE(A2:A100)` |
| **SEQUENCE** | Generate sequence | `SEQUENCE(rows, [cols], [start], [step])` | `=SEQUENCE(10,1,1,1)` |
| **RANDARRAY** | Random array | `RANDARRAY([rows], [cols], [min], [max], [whole])` | `=RANDARRAY(10,1,1,100,TRUE)` |
| **SORTBY** | Sort by array | `SORTBY(array, by_array, [sort_order])` | `=SORTBY(A2:B100,B2:B100,-1)` |
| **TAKE** | Take rows/cols | `TAKE(array, rows, [cols])` | `=TAKE(A2:C100,10)` → first 10 rows |
| **DROP** | Drop rows/cols | `DROP(array, rows, [cols])` | `=DROP(A2:C100,-10)` → drop last 10 |
| **TOCOL** | To column | `TOCOL(array, [ignore], [scan_by_column])` | `=TOCOL(A1:C10)` |
| **TOROW** | To row | `TOROW(array, [ignore], [scan_by_column])` | `=TOROW(A1:C10)` |
| **WRAPROWS** | Wrap to rows | `WRAPROWS(vector, wrap_count, [pad_with])` | `=WRAPROWS(A1:A100,5)` |
| **WRAPCOLS** | Wrap to columns | `WRAPCOLS(vector, wrap_count, [pad_with])` | `=WRAPCOLS(A1:A100,5)` |
| **VSTACK** | Stack vertically | `VSTACK(array1, array2, ...)` | `=VSTACK(A1:C10,E1:G20)` |
| **HSTACK** | Stack horizontally | `HSTACK(array1, array2, ...)` | `=HSTACK(A1:C10,E1:G10)` |
| **EXPAND** | Expand array | `EXPAND(array, rows, [cols], [pad_with])` | `=EXPAND(A1:C5,10,5)` |

---

## Information Functions

| Function | Description | Syntax | Example |
|----------|-------------|--------|---------|
| **ISBLANK** | Check if blank | `ISBLANK(value)` | `=ISBLANK(A1)` |
| **ISNUMBER** | Check if number | `ISNUMBER(value)` | `=ISNUMBER(A1)` |
| **ISTEXT** | Check if text | `ISTEXT(value)` | `=ISTEXT(A1)` |
| **ISERROR** | Check if error | `ISERROR(value)` | `=ISERROR(A1)` |
| **ISERR** | Check if error (not #N/A) | `ISERR(value)` | `=ISERR(A1)` |
| **ISNA** | Check if #N/A | `ISNA(value)` | `=ISNA(A1)` |
| **ISEVEN** | Check if even | `ISEVEN(number)` | `=ISEVEN(A1)` |
| **ISODD** | Check if odd | `ISODD(number)` | `=ISODD(A1)` |
| **ISLOGICAL** | Check if logical | `ISLOGICAL(value)` | `=ISLOGICAL(A1)` |
| **CELL** | Cell information | `CELL(info_type, [reference])` | `=CELL("address",A1)` |
| **INFO** | System information | `INFO(type_text)` | `=INFO("directory")` |
| **TYPE** | Value type | `TYPE(value)` | `=TYPE(A1)` |

---

## New Functions (Excel 365)

| Function | Description | Syntax | Example |
|----------|-------------|--------|---------|
| **LET** | Define variables | `LET(name1, value1, calculation)` | `=LET(x,A1,x^2)` |
| **LAMBDA** | Custom function | `LAMBDA(parameter, calculation)` | `=LAMBDA(x,x^2)(5)` |
| **XMATCH** | Enhanced MATCH | `XMATCH(lookup, lookup_array, [match_mode])` | `=XMATCH(A1,B:B)` |
| **TEXTBEFORE** | Text before delimiter | `TEXTBEFORE(text, delimiter)` | `=TEXTBEFORE("a-b-c","-")` → "a" |
| **TEXTAFTER** | Text after delimiter | `TEXTAFTER(text, delimiter)` | `=TEXTAFTER("a-b-c","-")` → "b-c" |
| **TEXTSPLIT** | Split text | `TEXTSPLIT(text, col_delim, [row_delim])` | `=TEXTSPLIT("a,b,c",",")` |
| **TOROW/TOCOL** | Transform arrays | `TOROW(array)`, `TOCOL(array)` | `=TOCOL(A1:C5)` |
| **TAKE/DROP** | Select/drop rows | `TAKE(array, n)`, `DROP(array, n)` | `=TAKE(A1:C100,10)` |

---

## Tips for Using Functions

1. **Use Formula AutoComplete**: Start typing `=` and Excel suggests functions
2. **Press F1**: Get help on selected function
3. **Use Evaluate Formula**: Formulas tab → Evaluate Formula to debug
4. **Named Ranges**: Make formulas more readable
5. **Function Wizard**: Shift+F3 to open Insert Function dialog

---

[Back to Resources](../README.md)
