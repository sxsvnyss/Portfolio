# Financial Modeling Kit

## Project Overview

This project showcases comprehensive financial modeling capabilities including budget forecasting, discounted cash flow (DCF) valuation, and sensitivity analysis. It provides a complete toolkit for financial planning and investment analysis.

## Business Problem

Finance teams and analysts need robust, flexible models to forecast future performance, evaluate investment opportunities, and support strategic decision-making. Manual calculations are error-prone and difficult to update when assumptions change.

## Key Features

- DCF model with automated valuation calculations
- Budget forecasting with variance analysis
- Sensitivity tables for scenario planning
- What-if analysis with data tables
- Dynamic charts showing financial projections

## Excel Skills Used

| Skill Category | Specific Skills |
|---------------|-----------------|
| Functions | NPV, IRR, XNPV, XIRR, PV, FV, PMT |
| Visualization | Waterfall Charts, Scenario Charts, Conditional Formatting |
| Data Tools | Data Tables, Scenario Manager, Goal Seek |
| Automation | Named Ranges, Dynamic Arrays |

## File Structure

```
02-Financial-Modeling/
├── data/
│   └── financial_data.xlsx          # Historical financial data
├── models/
│   ├── budget_model.xlsx            # Budget forecasting
│   ├── dcf_model.xlsx               # Discounted cash flow valuation
│   └── sensitivity_analysis.xlsx    # What-if analysis
├── images/
│   └── model_screenshot.png         # Visual preview
└── documentation/
    └── assumptions.md               # Model assumptions
```

## How to Use This Project

1. **Download the files**: Clone or download this folder
2. **Review historical data**: Open `data/financial_data.xlsx`
3. **Explore models**: Open individual model files in `models/` folder
4. **Modify assumptions**: Change input values in highlighted cells
5. **View scenarios**: Use Scenario Manager to compare different outcomes
6. **Analyze results**: Review output tables and charts

## Key Insights & Results

- DCF valuation shows 15% undervaluation compared to market price
- Budget variance analysis identified 12% cost overrun in operations
- Sensitivity analysis reveals NPV is most sensitive to revenue growth
- 5-year projection shows 22% CAGR under base case scenario
- Break-even analysis indicates 18 months to profitability

## Model Preview

![Financial Model Preview](./images/model_screenshot.png)

## Methodology

1. **Data Collection**: Gathered historical financial statements
2. **Assumption Development**: Created realistic growth and margin assumptions
3. **Model Building**: Built integrated 3-statement financial model
4. **Valuation**: Applied DCF methodology with WACC calculation
5. **Scenario Analysis**: Created best/worst/base case scenarios
6. **Validation**: Cross-checked outputs with manual calculations

## Technical Details

### Excel Functions Used

```excel
=NPV(rate, value1, [value2], ...)
=IRR(values, [guess])
=XNPV(rate, values, dates)
=XIRR(values, dates, [guess])
=PV(rate, nper, pmt, [fv], [type])
```

### Key Model Assumptions

- Revenue Growth: 8-15% annually
- Operating Margin: 18-22%
- WACC: 9.5%
- Terminal Growth Rate: 3%
- Tax Rate: 25%

### Scenario Analysis

| Scenario | Revenue Growth | Operating Margin | NPV |
|----------|---------------|------------------|-----|
| Bear Case | 5% | 15% | $45M |
| Base Case | 10% | 18% | $68M |
| Bull Case | 15% | 22% | $95M |

## What I Learned

- Mastered advanced financial functions in Excel
- Learned best practices for financial model design
- Understood importance of clear assumption documentation
- Improved scenario planning and sensitivity analysis skills
- Developed error-checking and validation techniques

## License

This project is provided for educational purposes. Feel free to use and adapt.

## Feedback

Found a bug or have a suggestion? Open an issue or submit a pull request!

---

[Back to Main Portfolio](../README.md)
