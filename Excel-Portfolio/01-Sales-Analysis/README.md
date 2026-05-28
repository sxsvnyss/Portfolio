# Sales Analysis Dashboard

## Project Overview

This project demonstrates comprehensive sales performance tracking and trend analysis using Excel. It provides an interactive dashboard that enables stakeholders to monitor key sales metrics, compare performance across regions and time periods, and identify growth opportunities.

## Business Problem

Sales teams need a centralized, visual way to track performance metrics, identify trends, and make data-driven decisions. Manual reporting is time-consuming and prone to errors, making it difficult to respond quickly to market changes.

## Key Features

- Interactive dashboard with slicers for filtering by region, product, and time period
- Year-over-year comparison charts showing growth trends
- Regional breakdown with map visualization
- Top performers leaderboard with dynamic ranking
- Automated data refresh capability

## Excel Skills Used

| Skill Category | Specific Skills |
|---------------|-----------------|
| Functions | XLOOKUP, SUMIFS, AVERAGEIFS, IFERROR, TEXT |
| Visualization | Pivot Charts, Slicers, Conditional Formatting, Sparklines |
| Data Tools | Pivot Tables, Power Pivot, Data Models |
| Automation | Power Query for data refresh |

## File Structure

```
01-Sales-Analysis/
├── data/
│   ├── raw_sales_data.xlsx          # Original dataset
│   └── cleaned_sales_data.xlsx      # Processed dataset
├── analysis/
│   ├── sales_dashboard.xlsx         # Final deliverable
│   ├── pivot_tables.xlsx            # Supporting pivot analysis
│   └── calculations.xlsx            # Formula calculations
├── images/
│   ├── dashboard_screenshot.png     # Visual preview
│   └── charts_preview.png           # Chart examples
└── documentation/
    ├── methodology.md               # Approach explanation
    └── insights.md                  # Key findings
```

## How to Use This Project

1. **Download the files**: Clone or download this folder
2. **Open the data files**: Review the raw and cleaned data in the `data/` folder
3. **Explore the dashboard**: Open `analysis/sales_dashboard.xlsx` and enable all features
4. **Enable macros** (if applicable): Click "Enable Content" in the security warning
5. **Interact with slicers**: Filter data by date, category, region, etc.
6. **Review formulas**: Check the formula bar to understand the logic

## Key Insights & Results

- Identified 23% increase in Q4 sales compared to Q3
- Discovered top 10 customers generate 65% of total revenue
- Regional analysis showed 40% growth in emerging markets
- Product category analysis revealed underperforming segments
- Automated report generation saved 5 hours per week

## Dashboard Preview

![Dashboard Preview](./images/dashboard_screenshot.png)

## Methodology

1. **Data Collection**: Gathered sales transaction data from ERP system
2. **Data Cleaning**: Removed duplicates, fixed formatting, handled missing values using Power Query
3. **Data Modeling**: Created relationships between sales, customer, and product tables in Power Pivot
4. **Visualization**: Built interactive dashboard with key KPIs and slicers
5. **Validation**: Verified calculations against source reports

## Technical Details

### Excel Functions Used

```excel
=XLOOKUP(lookup_value, lookup_array, return_array, "Not found")
=SUMIFS(sum_range, criteria_range1, criteria1, criteria_range2, criteria2)
=IFERROR(formula, "Error message")
=TEXT(value, format_code)
```

### Power Query Steps

1. Import data from Excel source
2. Remove top 3 rows (headers in row 4)
3. Promote headers
4. Change data types
5. Remove duplicates
6. Merge queries with customer table
7. Load to Data Model

## What I Learned

- Advanced my understanding of dynamic array formulas
- Learned best practices for dashboard design and user experience
- Improved data modeling skills with Power Pivot
- Understood importance of data validation and error handling

## License

This project is provided for educational purposes. Feel free to use and adapt.

## Feedback

Found a bug or have a suggestion? Open an issue or submit a pull request!

---

[Back to Main Portfolio](../README.md)
