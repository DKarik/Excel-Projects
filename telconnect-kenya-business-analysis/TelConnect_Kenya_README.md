# TelConnect Kenya — Telecommunications Business Performance Analysis

## Project overview
This project analyses a fictional Kenyan telecommunications company using Microsoft Excel. It explores completed recharge value, regional performance, subscription plans, payment methods, customer segments and transaction outcomes.

- **Analysis period:** January–December 2025
- **Currency:** Kenyan shillings (KES)
- **Tool:** Microsoft Excel
- **Synthetic dataset:** 250 customers, 1,500 recharge transactions, 400 support tickets, 120 network incidents, 8 plans and 10 regions/counties.

> **Important measurement note:** Completed `Recharge_Amount_KES` is used as a proxy for transaction revenue in this exercise. It is not audited company revenue, recognised accounting revenue or total telecommunications revenue. All data and the company are fictional.

## Business questions
1. How does completed recharge value change month by month?
2. Which counties contribute the most and least completed recharge value?
3. Which plans generate the highest completed recharge value?
4. Which payment methods contribute the most value?
5. Which customer segments contribute the most value?
6. How many transactions are completed, failed or reversed?
7. What are the average completed recharge value and transaction completion rate?
8. What actions could management consider based on the findings?

## Workbook contents
The supplied workbook includes:
- `START_HERE` — project introduction
- `PROJECT_TASKS` — guided tasks
- `Plans` — subscription plan reference data
- `Regions` — region and county reference data
- `Customers` — customer reference data
- `Usage_Data` — recharge and usage transactions
- `Support_Tickets` — customer support records
- `Network_Incidents` — network incident records
- `Analysis_Template` — transaction analysis table and formula columns
- `KPI_Guide` — KPI definitions
- `DATA_DICTIONARY` — field descriptions

Additional worksheets to create during the project:
- `Business_Analysis`
- `PT_Monthly_Performance`
- `PT_County_Performance`
- `PT_Plan_Performance`
- `PT_Payment_Performance`
- `PT_Segment_Performance`
- `PT_Transaction_Status`
- `Dashboard`
- `Business_Insights`

## Excel skills and techniques
This project practises:
- Data familiarisation and data-quality checks
- Excel Tables
- `VLOOKUP` with exact matching
- Absolute and relative references
- `IF`, nested `IF` and `IFERROR`
- `SUMIFS`, `COUNTIFS` and `AVERAGEIFS`
- PivotTables and PivotCharts
- Slicers and report connections
- KPI calculations
- Dashboard design
- Business interpretation and recommendations

## Analysis methodology

### 1. Data familiarisation and quality checks
Review the source worksheets, headers, field types, categories and relationships. Check for duplicate transaction IDs, missing key values, invalid dates, unexpected negative or zero amounts, inconsistent category labels, unmatched reference IDs and formula errors. Document only checks actually performed and issues actually identified.

### 2. Data enrichment with VLOOKUP
The `Analysis_Template` sheet enriches transactions with customer, plan and regional attributes. Examples of the lookup approach:

```excel
=VLOOKUP(C2,Customers!$A$2:$I$251,5,FALSE)
=VLOOKUP(D2,Plans!$A$2:$D$9,2,FALSE)
=VLOOKUP(E2,Regions!$A$2:$C$11,2,FALSE)
```

Use the exact ranges and column positions in the final workbook. `FALSE` requests an exact match; absolute references keep the lookup range fixed when formulas are filled down.

### 3. Conditional logic
Example recharge band formula:

```excel
=IF(J2<100,"Low",IF(J2<500,"Medium","High"))
```

Example transaction inclusion formula:

```excel
=IF(M2="Completed","Count","Exclude")
```

### 4. Conditional aggregation
Example total completed recharge value:

```excel
=SUMIFS(Analysis_Template!J:J,Analysis_Template!M:M,"Completed")
```

Use `SUMIFS`, `COUNTIFS` and `AVERAGEIFS` for the required county, payment method and transaction-status calculations.

### 5. PivotTable and dashboard analysis
Create PivotTables for monthly performance, county, plan, payment method, customer segment and transaction status. Use completed transactions for completed recharge value comparisons. Keep all statuses in the transaction-status analysis. Create PivotCharts and slicers, then connect slicers to the relevant PivotTables.

## KPI definitions

| KPI | Definition |
|---|---|
| Completed recharge value | Sum of recharge amounts where status is `Completed` |
| Total transactions | Count of all transaction records |
| Completed transactions | Count of records marked `Completed` |
| Failed transactions | Count of records marked `Failed` |
| Reversed transactions | Count of records marked `Reversed` |
| Average completed recharge | Average recharge amount for completed records |
| Completion rate | Completed transactions ÷ all transactions |
| Unsuccessful transaction rate | (Failed + reversed transactions) ÷ all transactions |

The completion and unsuccessful transaction rates above use all transaction records as the denominator. Clearly state any alternative definition if you use one.

## Findings and recommendations
**Complete this section using your actual workbook results. Do not publish placeholders or invented findings as real results.**

### Monthly performance
- **Finding:** [Highest/lowest month and actual completed recharge values.]
- **Business implication:** [What the pattern may indicate.]
- **Recommendation:** [Evidence-based action.]

### County performance
- **Finding:** [Highest/lowest counties and actual values.]
- **Business implication:** [What the regional contribution pattern may indicate.]
- **Recommendation:** [Practical action or further investigation.]

### Plan performance
- **Finding:** [Best- and lowest-performing plans by completed recharge value.]
- **Business implication:** [What this may suggest.]
- **Recommendation:** [Plan, pricing or customer research action.]

### Payment methods
- **Finding:** [Leading payment method and contribution.]
- **Business implication:** [What the payment pattern may indicate.]
- **Recommendation:** [Action to improve convenience or reliability.]

### Customer segments
- **Finding:** [Leading customer segment by completed recharge value.]
- **Business implication:** [What its contribution may indicate.]
- **Recommendation:** [Targeted engagement or retention action.]

### Transaction status
- **Finding:** [Completed, failed and reversed counts and unsuccessful transaction rate.]
- **Business implication:** [Possible impact on customer experience and transaction value.]
- **Recommendation:** [Monitoring, root-cause analysis or payment/network investigation.]

### Three overall recommendations
1. [Recommendation supported by a specific result.]
2. [Recommendation supported by a specific result.]
3. [Recommendation supported by a specific result.]

## Dashboard contents
The final dashboard should include:
- Completed recharge value
- Total transactions
- Completed transactions
- Failed transactions
- Average completed recharge
- Transaction completion rate
- Monthly completed recharge trend
- Completed recharge value by county and plan
- Payment method and customer segment contribution
- Interactive slicers where appropriate

## Screenshots
Save genuine, readable screenshots from the completed Excel workbook in a folder named `screenshots/`.

| Filename | Evidence |
|---|---|
| `01_data_structure.png` | Analysis data structure and enriched columns |
| `02_advanced_formulas.png` | Formula bar showing lookup and conditional/aggregation formulas |
| `03_pivot_tables.png` | Completed PivotTables |
| `04_pivot_charts.png` | PivotCharts |
| `05_final_dashboard.png` | Full dashboard with KPI cards and charts |
| `06_business_insights.png` | Findings, implications and recommendations |
| `07_dashboard_interaction.png` | Dashboard after changing a slicer selection |

On Windows, press `Win + Shift + S`, select a rectangular snip, capture the relevant Excel area and save as PNG.

## Limitations
- All data is synthetic and the company is fictional.
- Completed recharge value is a proxy for transaction revenue, not audited revenue or a complete measure of telecommunications income.
- The analysis describes patterns in the supplied dataset; it does not establish causation.
- Recharge value alone does not prove customer profitability, customer lifetime value, market share or network quality.
- Do not make claims about support tickets or network incidents unless you have analysed those records and linked them appropriately.

## How to use the workbook
1. Review `START_HERE` and `PROJECT_TASKS`.
2. Familiarise yourself with the source data and data dictionary.
3. Complete and document data-quality checks.
4. Populate the lookup and conditional formula columns in `Analysis_Template`.
5. Complete the conditional calculations in `Business_Analysis`.
6. Build and verify the PivotTables.
7. Create PivotCharts and connect slicers.
8. Build the dashboard.
9. Write evidence-based insights and recommendations.
10. Save the final workbook and add screenshots to the repository.

## Suggested GitHub repository structure

```text
telconnect-kenya-business-analysis/
├── README.md
├── TelConnect_Kenya_Business_Analysis_Final.xlsx
└── screenshots/
    ├── 01_data_structure.png
    ├── 02_advanced_formulas.png
    ├── 03_pivot_tables.png
    ├── 04_pivot_charts.png
    ├── 05_final_dashboard.png
    ├── 06_business_insights.png
    └── 07_dashboard_interaction.png
```

## Skills demonstrated
Microsoft Excel • Data quality assurance • VLOOKUP • Conditional formulas • SUMIFS • COUNTIFS • AVERAGEIFS • PivotTables • PivotCharts • Slicers • KPI reporting • Dashboard development • Data interpretation • Business recommendations

## Disclaimer
This is an educational portfolio project using synthetic data. TelConnect Kenya Ltd. is fictional. The results do not represent an actual telecommunications provider, its customers or the Kenyan telecommunications market.
