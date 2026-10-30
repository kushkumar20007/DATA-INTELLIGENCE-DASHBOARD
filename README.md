<div align="center">

# 📊 DATA INTELLIGENCE DASHBOARD

**Transform raw sales data into clear business insights with Microsoft Excel.**

A portfolio-ready analytics project featuring formula-driven analysis, customer segmentation, What-If modeling, regression, interactive PivotCharts, and an executive dashboard.


</div>

---

## 📌 Project Overview

This project demonstrates how Microsoft Excel can be used to analyze sales transactions, evaluate customer and product performance, explore business scenarios, and communicate findings through an interactive dashboard.

The workbook combines modern Excel functions, conditional formatting, What-If Analysis, the Analysis ToolPak, PivotTables, PivotCharts, slicers, and timeline filters.

> **Important:** Replace example field names with the columns in your own dataset. Add numerical findings only after validating them against the actual workbook.

## 🎯 Project Objectives

- Analyze sales, profit, quantity, order dates, customers, products, and regions.
- Identify high-value customers and frequently purchased products.
- Use formulas to filter, compare, summarize, and classify records.
- Evaluate business assumptions using Scenario Manager and Goal Seek.
- Explore relationships between variables with linear regression.
- Build an interactive dashboard with KPIs, charts, slicers, and a timeline.
- Present findings through concise, evidence-based business storytelling.

## 🧰 Tools & Features

| Tool / Feature | Purpose |
|---|---|
| Microsoft Excel Tables | Structured, expandable source data |
| Excel formulas | Date/time, filtering, matching, ranking, and aggregation |
| Conditional Formatting | Visual alerts, icons, arrows, and KPI status |
| What-If Analysis | Scenario Manager and Goal Seek |
| Analysis ToolPak | Linear regression and statistical output |
| PivotTables & PivotCharts | Summarize sales and performance |
| Slicers & Timelines | Interactive dashboard filtering |
| Markdown / GitHub | Project documentation and presentation |

## 🗂️ Recommended Workbook Structure

| Worksheet | Contents |
|---|---|
| `Raw_Data` | Original dataset; retain an untouched copy |
| `Clean_Data` | Validated and standardized records |
| `Formula_Analysis` | Date, text, filter, matching, and KPI formulas |
| `Customer_Analysis` | Customer totals, rankings, and segments |
| `What_If` | Scenario Manager and Goal Seek models |
| `Regression` | Regression inputs, output, and interpretation |
| `Pivot_Analysis` | PivotTables and supporting charts |
| `Dashboard` | Interactive KPIs and visual summaries |
| `Final_Report` | Findings, recommendations, and limitations |

## 🧾 Suggested Dataset Fields

Convert the source range to an Excel Table using **Ctrl + T** and name it `SalesData`.

| Field | Description |
|---|---|
| `OrderID` | Order or transaction identifier |
| `OrderDate` | Purchase date |
| `CustomerName` | Customer name |
| `ProductName` | Purchased product |
| `Category` | Product category |
| `Quantity` | Units sold |
| `Sales` | Revenue, according to the dataset's definition |
| `Discount` | Discount rate or amount |
| `Cost` | Cost, according to the dataset's definition |
| `Profit` | Profit amount |
| `Region` | Sales region |

Column names and business definitions may differ in your dataset. Confirm whether Sales is before or after discount and how Profit is calculated before interpreting results.

---
## 🎥 Video Demonstration

[![Watch Video](https://img.shields.io/badge/🎥-Watch_Video-red?style=for-the-badge)](https://drive.google.com/file/d/1NKcjtYavHwiJtay88BrI2yzv3vH23oG9/view?usp=sharing)
## 🧮 Key Excel Implementations

### 1. Date & Time Analysis

| Requirement | Formula | Purpose |
|---|---|---|
| Current date | `=TODAY()` | Returns today's date |
| Current date and time | `=NOW()` | Returns the current date/time at recalculation |
| Completed years | `=DATEDIF([@OrderDate],TODAY(),"Y")` | Calculates full years since the order date |
| Month-end date | `=EOMONTH([@OrderDate],0)` | Returns the last day of the order month |
| Next month's end | `=EOMONTH(TODAY(),1)` | Returns the end of next month |

Ensure date fields contain valid Excel date values, not text.

### 2. Dynamic Filtering with `FILTER`

Return all records for the customer selected in cell `B2`:

```excel
=FILTER(SalesData,SalesData[CustomerName]=B2,"No matches")
```

Return records with positive profit:

```excel
=FILTER(SalesData,SalesData[Profit]>0,"No profitable rows")
```

Return records from the selected region:

```excel
=FILTER(SalesData,SalesData[Region]=B2,"No matches")
```

**Compatibility:** `FILTER` requires a supported Excel version, such as Microsoft 365 or Excel 2021 and later.

### 3. Conditional Formatting & Icon Indicators

Apply conditional formatting to highlight performance and exceptions.

**Suggested implementation**
1. Select the Profit or KPI range.
2. Open **Home → Conditional Formatting**.
3. Choose rules, color scales, data bars, or **Icon Sets**.
4. Open **Manage Rules → Edit Rule** to customize thresholds.
5. Document the thresholds and test positive, negative, and boundary values.

Example interpretation:
- 🟢 Green / upward arrow: meeting or exceeding the target.
- 🟡 Yellow / sideways arrow: close to the target.
- 🔴 Red / downward arrow: below the target.

Set thresholds based on a meaningful business target rather than choosing colors arbitrarily.

### 4. Timestamp Creation

A dynamic timestamp:

```excel
=NOW()
```

This value can change when Excel recalculates; it is not a permanent entry timestamp.

For a fixed timestamp using iterative calculation, suppose `A2` is the trigger field and `K2` stores the timestamp:

```excel
=IF(A2="","",IF(K2="",NOW(),K2))
```

To test this approach:
1. Save a copy of the workbook.
2. Open **File → Options → Formulas**.
3. Enable **Enable iterative calculation** and set **Maximum Iterations** to `1`.
4. Test changes, clearing behavior, and recalculation carefully.

Circular-reference timestamps depend on workbook calculation settings. For dependable audit logs, consider an approved VBA or Power Automate solution.

### 5. High-Value Customer Analysis with Formulas

Assume a customer name is in cell `A2`.

**Total sales per customer**

```excel
=SUMIFS(SalesData[Sales],SalesData[CustomerName],A2)
```

**Total profit per customer**

```excel
=SUMIFS(SalesData[Profit],SalesData[CustomerName],A2)
```

**Number of transaction rows**

```excel
=COUNTIF(SalesData[CustomerName],A2)
```

**Rank customers by total sales**  
If customer sales totals are in `B2:B100`:

```excel
=RANK.EQ(B2,$B$2:$B$100,0)
```

Define “high-value” transparently—for example, the top 10 customers by profit, or customers above a documented sales threshold. Compare sales and profit separately because high revenue does not always mean high profitability.

> **Data-quality note:** `COUNTIF` counts matching rows, not necessarily unique orders. If orders contain multiple product rows, calculate distinct OrderIDs separately.

### 6. Most Frequently Purchased Product

Count the number of rows for a product listed in `A2`:

```excel
=COUNTIF(SalesData[ProductName],A2)
```

If product names are in `A2:A100` and their counts are in `B2:B100`, find the highest frequency:

```excel
=MAX(B2:B100)
```

Return the first product with that frequency:

```excel
=INDEX(A2:A100,MATCH(MAX(B2:B100),B2:B100,0))
```

This measures transaction-row frequency, not total units sold. To identify the product with the greatest quantity, summarize `Quantity` by `ProductName` instead. Check for ties when reporting the most frequent product.

### 7. Compare Two Lists & Extract Matches

Suppose List 1 is in `A2:A100` and List 2 is in `D2:D100`.

Check whether a name from List 1 appears in List 2:

```excel
=IF(COUNTIF($D$2:$D$100,A2)>0,"Match","No match")
```

Extract matching names:

```excel
=FILTER(A2:A100,COUNTIF(D2:D100,A2:A100)>0,"No matches")
```

Return unique matching names in supported Excel versions:

```excel
=UNIQUE(FILTER(A2:A100,COUNTIF(D2:D100,A2:A100)>0,"No matches"))
```

For inconsistent source lists, standardize spelling and remove unwanted spaces before comparing.

### 8. Create Abbreviations with Text Functions

For a two-word name in `A2`, return the first initial of each word:

```excel
=UPPER(LEFT(A2,1)&MID(A2,FIND(" ",A2)+1,1))
```

Return the first three characters:

```excel
=UPPER(LEFT(A2,3))
```

For a multiword label in Excel versions supporting `TEXTSPLIT`:

```excel
=TEXTJOIN("",TRUE,LEFT(TEXTSPLIT(A2," "),1))
```

The first formula assumes two words separated by a space. Test text formulas against blank cells, extra spaces, and names with more than two words.

---

## 📉 What-If Analysis

### Scenario Manager

Build a formula-driven model that calculates profit from relevant inputs. Compare clearly defined assumptions, such as:

- **Base case:** current discount and demand assumptions.
- **Higher-discount case:** increased discount and a stated demand assumption.
- **Lower-discount case:** reduced discount and a stated demand assumption.

**Steps**
1. Identify the input cells that drive the result.
2. Open **Data → What-If Analysis → Scenario Manager**.
3. Add each scenario and enter its input values.
4. Create a Scenario Summary using the Profit formula cell as the result.
5. Compare outcomes and explain the assumptions behind each scenario.

### Goal Seek

Use Goal Seek to find the input value required to reach a target profit.

1. Create a formula-driven Profit result cell.
2. Open **Data → What-If Analysis → Goal Seek**.
3. **Set cell:** the Profit formula cell.
4. **To value:** the target profit.
5. **By changing cell:** the selected input, such as discount rate or units sold.
6. Check whether the resulting input is realistic.

Goal Seek changes one input cell at a time. It does not prove that a target is achievable under all business constraints.

---

## 📈 Linear Regression — Data Analysis ToolPak

Example question: **How is Sales statistically associated with Profit in this dataset?**

### Setup

1. Enable the add-in through **File → Options → Add-ins → Excel Add-ins → Go → Analysis ToolPak**.
2. Arrange the chosen outcome (Y) and predictor (X) in numeric columns.
3. Open **Data → Data Analysis → Regression**.
4. Set **Input Y Range** to the outcome, such as Profit.
5. Set **Input X Range** to the predictor, such as Sales.
6. Select **Labels** if the ranges include headers.
7. Choose an output range or new worksheet and run the analysis.
8. Review R Square, coefficients, p-values, sample size, and residuals.

<details>
<summary><strong>How to interpret the output</strong></summary>

- **Coefficient:** estimated change in the outcome associated with a one-unit change in the predictor, under the fitted model.
- **R Square:** proportion of variation in the outcome explained by the fitted model in the sample.
- **P-value:** evidence against the specified null hypothesis, interpreted alongside assumptions and context.
- **Residuals:** differences between observed and model-predicted outcomes.

Regression does not establish causation. Investigate outliers, missing values, non-linearity, data leakage, duplicate observations, and whether the model assumptions are reasonable. Avoid treating transaction rows as independent observations when the data structure makes that assumption inappropriate.
</details>

<LearningViz type_id="LEAST_SQUARE_REGRESSION" initial_values={{"fitMode":"with_intercept"}} />

---

## 📊 Interactive Dashboard

Create a dedicated `Dashboard` worksheet designed for quick business decisions.

### Recommended KPI Cards

| KPI | Example calculation |
|---|---|
| Total Sales | `=SUM(SalesData[Sales])` |
| Total Profit | `=SUM(SalesData[Profit])` |
| Profit Margin | `=IFERROR(SUM(SalesData[Profit])/SUM(SalesData[Sales]),0)` |
| Total Quantity | `=SUM(SalesData[Quantity])` |

Add unique customers and unique orders using a distinct count in the Data Model or a suitable formula for your Excel version.

### PivotTables & PivotCharts

| Analysis | Suggested visualization |
|---|---|
| Sales and profit by month | Line chart |
| Sales by product or category | Bar/column chart |
| Sales share by category | Pie chart with a small number of categories |
| Profit by region | Bar/column chart |
| Top 10 customers | Sorted bar chart |
| Quantity by product | Bar chart |

**Build the charts**
1. Select a cell in `SalesData`.
2. Choose **Insert → PivotTable**.
3. Place fields into Rows, Columns, Values, and Filters.
4. Insert a PivotChart from the PivotTable.
5. Use clear chart titles, consistent units, and appropriate number formats.
6. Sort ranking charts from highest to lowest.

### Slicers & Timeline

1. Select a PivotTable and choose **PivotTable Analyze → Insert Slicer**.
2. Add slicers for fields such as Region, Category, or Customer.
3. Choose **Insert Timeline** for a valid date field.
4. Use **Report Connections** to connect slicers to compatible PivotTables.
5. Test the filters and refresh PivotTables after changing the source data.

### Dashboard Design Principles

- Place the most important KPIs at the top.
- Use consistent number formats and units.
- Keep charts focused on one question each.
- Avoid unnecessary decoration and overcrowded visuals.
- Add a visible date range and a note describing the metric definitions.
- Verify that slicers and timelines filter all intended visuals.

---

## 🧠 Data Analysis & Business Storytelling

Use the following process for each finding:

1. **Business question:** What decision needs support?
2. **Metric:** Define the measure and reporting period.
3. **Evidence:** Calculate the result and cite the relevant table or chart.
4. **Interpretation:** Explain what the result may indicate.
5. **Recommendation:** Suggest a specific, practical action.
6. **Limitation:** State relevant data gaps or uncertainty.

### Questions to Investigate

- Which month had the highest sales and profit?
- Which products generate high sales but low profit?
- Which customers contribute the most profit?
- How do discount levels relate to sales and margins?
- Which regions or categories are underperforming?
- Which products are purchased most frequently?
- How do profit outcomes change across the modeled scenarios?

### Insight Template

> **Finding:** [Verified result and reporting period]  
> **Evidence:** [Relevant KPI, table, formula, or chart]  
> **Interpretation:** [What the result may suggest]  
> **Recommended action:** [Specific next step]  
> **Limitation:** [Data or analytical caveat]

Do not invent findings or claim that one factor caused another without suitable evidence.

## 📝 Final Report Structure

1. Executive Summary
2. Dataset Overview & Data Quality
3. KPI Summary
4. Monthly Sales & Profit Trends
5. Product and Category Performance
6. Regional Performance
7. High-Value Customer Analysis
8. Frequently Purchased Products
9. What-If Scenario Results
10. Regression Results & Limitations
11. Recommendations
12. Appendix: Metric Definitions and Formulas

## ✅ Project Completion Checklist

- [ ] Source data converted to an Excel Table and validated.
- [ ] Missing values, duplicates, and inconsistent labels reviewed.
- [ ] Date/time formulas and dynamic filters tested.
- [ ] Conditional formatting thresholds documented.
- [ ] Timestamp behavior tested on a workbook copy.
- [ ] Scenario Manager and Goal Seek outputs explained.
- [ ] Regression results interpreted with appropriate caveats.
- [ ] High-value customer rankings validated.
- [ ] Product frequency and matching-list formulas checked.
- [ ] Abbreviation formulas tested on real examples.
- [ ] PivotTables, PivotCharts, slicers, and timeline connected correctly.
- [ ] Dashboard KPIs reconcile with the underlying data.
- [ ] Final report includes evidence-based insights and recommendations.

## ⚙️ Compatibility Notes

- Microsoft Excel for Windows is recommended for the complete ToolPak and What-If Analysis workflow.
- `FILTER`, `UNIQUE`, and `TEXTSPLIT` require supported Excel versions.
- Formula argument separators may be commas or semicolons depending on regional settings.
- PivotTable features can vary by Excel version and source configuration.
- Keep an untouched copy of the original dataset and document important assumptions.

## 🚀 How to Use This Project

1. Place the source data in `Raw_Data`.
2. Convert it to an Excel Table and verify the field names.
3. Clean and validate the data.
4. Complete the formula and customer analyses.
5. Run the What-If and regression analyses.
6. Build and test the PivotTables and dashboard.
7. Write the final report using verified results.
8. Capture genuine screenshots of the completed workbook if you want to showcase it in your portfolio.

## 👤 Author

**Kush Kumar**  
Excel Data Analysis | Business Intelligence | Data Storytelling

Add your GitHub repository link and a genuine screenshot of the finished dashboard here.

---

<div align="center">

**Built with Microsoft Excel · Focused on data-driven decisions**

</div>
