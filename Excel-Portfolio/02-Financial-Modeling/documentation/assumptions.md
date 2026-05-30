# Financial Model Assumptions

## Overview

This document details all key assumptions used in the financial modeling kit. These assumptions drive the projections and valuations in our models.

## Macroeconomic Assumptions

| Factor | Assumption | Source |
|--------|------------|--------|
| GDP Growth Rate | 2.5% annually | Federal Reserve Projections |
| Inflation Rate | 2.0% annually | Central Bank Target |
| Risk-Free Rate | 3.5% | 10-Year Treasury Yield |
| Market Risk Premium | 6.0% | Historical Average |
| Corporate Tax Rate | 25% | Current Tax Code |

## Company-Specific Assumptions

### Revenue Growth

| Period | Growth Rate | Rationale |
|--------|-------------|-----------|
| Year 1 | 12% | New product launches |
| Year 2 | 15% | Market expansion |
| Year 3 | 10% | Market maturation |
| Year 4 | 8% | Stable growth |
| Year 5 | 6% | Mature market |

### Margin Assumptions

| Metric | Current | Target | Timeline |
|--------|---------|--------|----------|
| Gross Margin | 45% | 48% | 3 years |
| Operating Margin | 18% | 22% | 3 years |
| Net Margin | 13% | 17% | 3 years |

### Capital Structure

| Component | Weight | Cost |
|-----------|--------|------|
| Debt | 40% | 5.5% (after-tax) |
| Equity | 60% | 12.5% |
| **WACC** | **100%** | **9.5%** |

## DCF Model Assumptions

### Discount Rate Calculation

```
Risk-Free Rate: 3.5%
Beta: 1.2
Market Risk Premium: 6.0%
Cost of Equity: 3.5% + (1.2 × 6.0%) = 10.7%

Cost of Debt: 7.0%
Tax Rate: 25%
After-Tax Cost of Debt: 7.0% × (1 - 0.25) = 5.25%

WACC = (0.60 × 10.7%) + (0.40 × 5.25%) = 9.5%
```

### Terminal Value

- **Terminal Growth Rate**: 3.0% (in line with long-term inflation + GDP growth)
- **Exit Multiple**: 12x EBITDA (based on comparable company analysis)
- **Perpetuity Method**: Used for primary valuation

### Projection Period

- **Explicit Forecast**: 5 years
- **Rationale**: Sufficient to capture near-term growth initiatives and reach stable growth phase

## Sensitivity Analysis Parameters

### Variables Tested

1. **Revenue Growth**: ±3% from base case
2. **Operating Margin**: ±300 basis points
3. **WACC**: ±150 basis points
4. **Terminal Growth**: ±1%

### Scenario Definitions

| Scenario | Revenue Growth | Margin | WACC |
|----------|---------------|--------|------|
| Bear Case | Base - 3% | Base - 3% | Base + 1.5% |
| Base Case | As projected | As projected | As calculated |
| Bull Case | Base + 3% | Base + 3% | Base - 1.5% |

## Working Capital Assumptions

| Metric | Days | Industry Benchmark |
|--------|------|-------------------|
| DSO (Receivables) | 45 days | 40-50 days |
| DPO (Payables) | 60 days | 55-65 days |
| DIO (Inventory) | 30 days | 30-40 days |

## Capital Expenditure Assumptions

| Category | % of Revenue | Purpose |
|----------|-------------|---------|
| Maintenance CapEx | 3% | Replace existing assets |
| Growth CapEx | 5% | Expand capacity |
| **Total** | **8%** | |

## Key Risks & Considerations

1. **Market Risk**: Economic downturn could impact revenue growth
2. **Competition**: New entrants may pressure margins
3. **Regulatory**: Changes in tax policy could affect valuations
4. **Execution**: Ability to achieve operational improvements

## Assumption Review Schedule

- **Quarterly**: Review actual vs. projected performance
- **Annually**: Update long-term assumptions
- **Event-Driven**: Revise when material changes occur

---

[Back to Project](../README.md)
