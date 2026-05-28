# Power Query ETL Steps Documentation

## Overview

This document details each step in the Power Query ETL process, including the M code and transformation logic.

---

## Query 1: Load Source Data 1

### Connection Setup

1. **Data Tab** → Get Data → From File → From Excel Workbook
2. Select `source_data_1.xlsx`
3. Choose the worksheet or table
4. Click **Transform Data**

### Applied Steps

| Step | Operation | Description |
|------|-----------|-------------|
| 1 | Source | Import from Excel file |
| 2 | Promoted Headers | Use first row as headers |
| 3 | Changed Type | Set correct data types |
| 4 | Removed Columns | Remove unnecessary columns |
| 5 | Filtered Rows | Exclude invalid records |
| 6 | Trimmed Text | Clean text fields |

### M Code

```powerquery
let
    Source = Excel.Workbook(File.Contents("C:\Data\source_data_1.xlsx"), null, true),
    Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{
        {"ID", Int64.Type},
        {"Date", type date},
        {"Amount", Currency.Type},
        {"Category", type text}
    }),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type",{"Unused Column"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Columns", each [Amount] <> null),
    #"Trimmed Text" = Table.TransformColumns(#"Filtered Rows",{{"Category", Text.Trim, type text}})
in
    #"Trimmed Text"
```

---

## Query 2: Load Source Data 2

### Connection Setup

Same process as Query 1, connecting to `source_data_2.xlsx`

### Applied Steps

| Step | Operation | Description |
|------|-----------|-------------|
| 1 | Source | Import from Excel file |
| 2 | Promoted Headers | Use first row as headers |
| 3 | Changed Type | Set correct data types |
| 4 | Renamed Columns | Standardize column names |
| 5 | Replaced Values | Fix inconsistent values |

### M Code

```powerquery
let
    Source = Excel.Workbook(File.Contents("C:\Data\source_data_2.xlsx"), null, true),
    Data_Sheet = Source{[Item="Data",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Data_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{
        {"CustomerID", Int64.Type},
        {"Name", type text},
        {"Email", type text},
        {"Phone", type text}
    }),
    #"Renamed Columns" = Table.RenameColumns(#"Changed Type",{
        {"CustomerID", "ID"},
        {"Name", "Customer Name"}
    }),
    #"Replaced Values" = Table.ReplaceValue(#"Renamed Columns","N/A",null,Replacer.ReplaceText,{"Email"})
in
    #"Replaced Values"
```

---

## Query 3: Load Source Data 3

### Connection Setup

Connecting to `source_data_3.xlsx`

### Applied Steps

| Step | Operation | Description |
|------|-----------|-------------|
| 1 | Source | Import from Excel file |
| 2 | Promoted Headers | Use first row as headers |
| 3 | Changed Type | Set correct data types |
| 4 | Pivoted Column | Transform rows to columns |
| 5 | Filled Down | Fill null values from above |

### M Code

```powerquery
let
    Source = Excel.Workbook(File.Contents("C:\Data\source_data_3.xlsx"), null, true),
    Products_Sheet = Source{[Item="Products",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(Products_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{
        {"ProductID", Int64.Type},
        {"ProductName", type text},
        {"Attribute", type text},
        {"Value", type text}
    }),
    #"Pivoted Column" = Table.Pivot(#"Changed Type", 
        List.Distinct(#"Changed Type"[Attribute]), 
        "Attribute", 
        "Value"
    ),
    #"Filled Down" = Table.FillDown(#"Pivoted Column",{"Price", "Cost"})
in
    #"Filled Down"
```

---

## Query 4: Merge All Sources

### Merge Operation

Combines all three source queries into a single dataset.

### Applied Steps

| Step | Operation | Description |
|------|-----------|-------------|
| 1 | Merged Queries | Join Query 1 and Query 2 on ID |
| 2 | Expanded Table | Expand merged columns |
| 3 | Merged Again | Join with Query 3 on ProductID |
| 4 | Expanded Again | Expand final merged columns |
| 5 | Reordered Columns | Arrange in logical order |

### M Code

```powerquery
let
    Source1 = #"Load Source Data 1",
    Source2 = #"Load Source Data 2",
    Source3 = #"Load Source Data 3",
    
    // First merge: Sales + Customer
    #"Merged Queries" = Table.NestedJoin(Source1, {"CustomerID"}, Source2, {"ID"}, "CustomerInfo", JoinKind.LeftOuter),
    #"Expanded CustomerInfo" = Table.ExpandTableColumn(#"Merged Queries", "CustomerInfo", 
        {"Customer Name", "Email", "Phone"}, 
        {"Customer Name", "Email", "Phone"}
    ),
    
    // Second merge: Add Product info
    #"Merged Queries1" = Table.NestedJoin(#"Expanded CustomerInfo", {"ProductID"}, Source3, {"ProductID"}, "ProductInfo", JoinKind.LeftOuter),
    #"Expanded ProductInfo" = Table.ExpandTableColumn(#"Merged Queries1", "ProductInfo", 
        {"ProductName", "Price", "Cost"}, 
        {"ProductName", "Price", "Cost"}
    ),
    
    // Calculate profit
    #"Added Custom" = Table.AddColumn(#"Expanded ProductInfo", "Profit", each [Amount] - [Cost]),
    #"Added Margin" = Table.AddColumn(#"Added Custom", "Margin %", each [Profit] / [Amount]),
    
    // Final cleanup
    #"Reordered Columns" = Table.ReorderColumns(#"Added Margin",{
        "ID", "Date", "Customer Name", "Email", "Phone",
        "ProductID", "ProductName", "Amount", "Cost", "Profit", "Margin %"
    })
in
    #"Reordered Columns"
```

---

## Query 5: Final Transformations

### Data Quality Enhancements

Final cleaning and validation before loading.

### Applied Steps

| Step | Operation | Description |
|------|-----------|-------------|
| 1 | Removed Duplicates | Eliminate duplicate records |
| 2 | Validated Email | Check email format |
| 3 | Conditional Column | Create status flags |
| 4 | Sorted Rows | Order by date |
| 5 | Buffered Query | Improve performance |

### M Code

```powerquery
let
    Source = #"Merge All Sources",
    
    // Remove exact duplicates
    #"Removed Duplicates" = Table.Distinct(Source),
    
    // Validate email format
    #"Added Email Valid" = Table.AddColumn(#"Removed Duplicates", "Email Valid", 
        each Text.Contains([Email], "@") and Text.Contains([Email], ".")),
    
    // Create status flag
    #"Added Conditional" = Table.AddColumn(#"Added Email Valid", "Status", 
        each if [Margin %] > 0.3 then "High Margin"
             else if [Margin %] > 0.1 then "Normal"
             else "Low Margin"),
    
    // Sort by date descending
    #"Sorted Rows" = Table.Sort(#"Added Conditional",{{"Date", Order.Descending}}),
    
    // Buffer for performance
    #"Buffered Table" = Table.Buffer(#"Sorted Rows")
in
    #"Buffered Table"
```

---

## Loading Options

### Load To Worksheet

```powerquery
// In Excel: Home → Close & Load To → Table
// Creates a new worksheet with the query results
```

### Load To Data Model

```powerquery
// In Excel: Home → Close & Load To → Only Create Connection
// Check "Add to Data Model" for Power Pivot integration
```

### Load As Connection Only

```powerquery
// For use as a data source for other queries
// No output to worksheet
```

---

## Refresh Process

### Manual Refresh

1. Right-click query result table
2. Select **Refresh**
3. Or: Data tab → Refresh All

### Scheduled Refresh

1. File → Options → Data
2. Enable "Enable automatic refresh on file open"
3. Set refresh interval (if using Power BI Gateway)

### VBA Auto-Refresh

```vba
Sub RefreshAllQueries()
    ThisWorkbook.RefreshAll
    Application.CalculateFull
End Sub
```

---

## Performance Optimization Tips

1. **Use Query Folding**: Keep transformations simple to enable folding
2. **Filter Early**: Apply filters as early as possible in the pipeline
3. **Remove Columns**: Eliminate unused columns before merging
4. **Disable Load**: For intermediate queries, disable load to worksheet
5. **Use Parameters**: Make file paths dynamic for portability

---

[Back to Project](../README.md)
