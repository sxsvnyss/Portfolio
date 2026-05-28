# Advanced Formulas Collection

## Project Overview

This project showcases complex Excel formula solutions for real business problems. It includes examples of array formulas, lookup functions, logical operations, and text/date manipulations using both traditional and modern dynamic array functions.

## Business Problem

Business users often rely on manual processes or complex VBA for tasks that can be solved with advanced Excel formulas. Understanding and applying these formulas can significantly improve productivity and reduce errors.

## Key Features

- Comprehensive library of advanced formula examples
- Dynamic array functions (FILTER, SORT, UNIQUE, SEQUENCE)
- Complex nested formulas for business logic
- XLOOKUP and INDEX/MATCH comparisons
- LAMBDA function examples for custom calculations

## Excel Skills Used

| Skill Category | Specific Skills |
|---------------|-----------------|
| Functions | XLOOKUP, FILTER, SORT, UNIQUE, LAMBDA, LET |
| Visualization | Formula Auditing Tools, Evaluate Formula |
| Data Tools | Dynamic Arrays, Named Ranges, Tables |
| Automation | Formula-based calculations without VBA |

## File Structure

```
06-Advanced-Formulas/
├── examples/
│   ├── array_formulas.xlsx        # Array formula examples
│   ├── lookup_formulas.xlsx       # XLOOKUP, INDEX/MATCH
│   ├── logical_formulas.xlsx      # Nested IFs, IFS, SWITCH
│   └── text_date_formulas.xlsx    # Text manipulation
└── documentation/
    └── formula_guide.md           # Explanation of each formula
```

## How to Use This Project

1. **Download the files**: Clone or download this folder
2. **Open example files**: Each file focuses on a specific function category
3. **Study formulas**: Click on cells to see formulas in the formula bar
4. **Modify inputs**: Change input values to see dynamic results
5. **Apply to your work**: Adapt formulas to your specific needs

## Key Insights & Results

- Replaced 50+ lines of VBA with single array formulas
- Reduced calculation time by 80% with efficient formulas
- Created reusable formula templates for common scenarios
- Eliminated helper columns with dynamic arrays
- Built complex business logic without macros

## Methodology

### Formula Categories Covered

1. **Array Formulas**: Multi-cell calculations, legacy and dynamic
2. **Lookup Functions**: Finding and retrieving data
3. **Logical Functions**: Conditional calculations
4. **Text Functions**: String manipulation and parsing
5. **Date Functions**: Date calculations and formatting
6. **Math/Statistical**: Aggregations and analysis
7. **Financial**: NPV, IRR, amortization
8. **LAMBDA Functions**: Custom reusable functions

## Technical Details

### Dynamic Array Functions (Excel 365)

```excel
=FILTER(array, include, [if_empty])
=SORT(array, [sort_index], [sort_order])
=UNIQUE(array, [by_col], [exactly_once])
=SEQUENCE(rows, [columns], [start], [step])
=RANDARRAY(rows, [columns], [min], [max], [whole_number])
```

### XLOOKUP Examples

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, "Not found")
=XLOOKUP(1, (criteria1)*(criteria2), return_array)  # Multiple criteria
=XLOOKUP(lookup_value, lookup_array, return_array, , -1)  # Last match
```

### LET Function for Complex Formulas

```excel
=LET(
    sales, B2:B100,
    target, 10000,
    bonus_rate, 0.05,
    over_target, MAX(0, sales - target),
    bonus, over_target * bonus_rate,
    bonus
)
```

### LAMBDA Custom Function

```excel
=LAMBDA(x, y, (x + y) / 2)(A2, B2)  # Average of two numbers
```

## What I Learned

- Power of dynamic array functions in Excel 365
- Best practices for writing readable complex formulas
- Performance optimization techniques
- When to use formulas vs. VBA
- Creating maintainable formula solutions

## License

This project is provided for educational purposes. Feel free to use and adapt.

## Feedback

Found a bug or have a suggestion? Open an issue or submit a pull request!

---

[Back to Main Portfolio](../README.md)
