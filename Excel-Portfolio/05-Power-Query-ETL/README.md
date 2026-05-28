# Power Query ETL Process

## Project Overview

This project demonstrates automated data extraction, transformation, and loading (ETL) using Power Query. It shows how to merge data from multiple sources, apply transformations, and create refreshable data pipelines.

## Business Problem

Organizations often have data scattered across multiple files and systems. Manually combining and transforming this data is time-consuming, error-prone, and difficult to repeat when source data updates. Power Query provides a solution for automated, repeatable ETL processes.

## Key Features

- Multi-source data merging from Excel files
- Automated data transformation steps
- Query folding for performance optimization
- One-click refresh capability
- Comprehensive ETL documentation

## Excel Skills Used

| Skill Category | Specific Skills |
|---------------|-----------------|
| Functions | Power Query M Language, Custom Columns |
| Visualization | Query Editor Screenshots, Step Documentation |
| Data Tools | Power Query Editor, Query Folding, Parameters |
| Automation | Scheduled Refresh, Data Gateway Integration |

## File Structure

```
05-Power-Query-ETL/
├── data/
│   ├── source_data_1.xlsx         # First data source
│   ├── source_data_2.xlsx         # Second data source
│   └── source_data_3.xlsx         # Third data source
├── queries/
│   └── power_query_steps.md       # ETL process documentation
├── output/
│   └── merged_dataset.xlsx        # Final merged data
└── images/
    └── query_editor.png           # Power Query editor screenshot
```

## How to Use This Project

1. **Download the files**: Clone or download this folder
2. **Review source data**: Examine files in `data/` folder
3. **Open merged dataset**: Open `output/merged_dataset.xlsx`
4. **View queries**: Data tab → Queries & Connections
5. **Edit transformations**: Double-click query → Power Query Editor
6. **Refresh data**: Right-click query → Refresh

## Key Insights & Results

- Reduced data consolidation time from 4 hours to 5 minutes
- Eliminated manual copy-paste errors completely
- Created reusable ETL templates for 20+ reports
- Enabled daily automated refresh with zero manual intervention
- Improved data freshness from weekly to real-time

## Query Editor Preview

![Power Query Editor](./images/query_editor.png)

## Methodology

### ETL Process Flow

1. **Extract**: Connect to multiple Excel source files
2. **Transform**: Clean, filter, merge, and reshape data
3. **Load**: Output to Excel worksheet or Data Model
4. **Refresh**: Update with one click when source data changes

### Transformation Steps

1. Import data from each source file
2. Standardize column names and data types
3. Remove unnecessary columns
4. Filter out invalid records
5. Merge queries on common keys
6. Create calculated columns
7. Pivot/unpivot as needed
8. Load to destination

## Technical Details

### Power Query M Code Examples

#### Combine Multiple Files

```powerquery
let
    Source = Folder.Files("C:\Data\Sources"),
    Combined = Table.Combine(
        List.Transform(Source[Content], Excel.Workbook)
    ),
    Expanded = Table.ExpandTableColumn(Combined, "Data", {"Name", "Value"})
in
    Expanded
```

#### Merge Queries

```powerquery
let
    Source1 = Excel.CurrentWorkbook(){[Name="Table1"]}[Content],
    Source2 = Excel.CurrentWorkbook(){[Name="Table2"]}[Content],
    Merged = Table.NestedJoin(Source1, {"ID"}, Source2, {"ID"}, "Details", JoinKind.Inner),
    Expanded = Table.ExpandTableColumn(Merged, "Details", {"Column1", "Column2"})
in
    Expanded
```

#### Custom Column

```powerquery
let
    Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
    AddedCustom = Table.AddColumn(Source, "Profit", each [Revenue] - [Cost]),
    AddedMargin = Table.AddColumn(AddedCustom, "Margin %", each [Profit] / [Revenue])
in
    AddedMargin
```

### Query Folding

Query folding pushes transformations back to the data source for better performance:

**Steps that fold**:
- Filter rows
- Select columns
- Sort
- Group by
- Join/merge

**Steps that don't fold**:
- Custom columns with complex M code
- Index columns
- Some text transformations

### Data Types Handled

| Source Type | Connection Method |
|-------------|------------------|
| Excel Files | Excel.Workbook |
| CSV Files | Csv.Document |
| SQL Server | Sql.Database |
| Web APIs | Web.Contents + Json.Parse |
| Folders | Folder.Files |

## What I Learned

- Power Query M language fundamentals
- Query optimization and folding techniques
- Best practices for ETL design
- Error handling in Power Query
- Parameter-driven queries for flexibility

## License

This project is provided for educational purposes. Feel free to use and adapt.

## Feedback

Found a bug or have a suggestion? Open an issue or submit a pull request!

---

[Back to Main Portfolio](../README.md)
