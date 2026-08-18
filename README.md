# Sales Performance Analysis

An end-to-end Excel project that cleans raw retail sales data, builds calculated financial metrics, and turns them into an interactive, dynamic dashboard for data-driven decision-making.

![Final Dashboard](assets/final-dashboard.jpeg)

---

## 📌 Overview

This project analyzes retail sales transactions to answer key business questions around revenue, cost, and profit performance, order status trends, and product-level profitability. The end result is a fully automated Excel dashboard — clean, interactive, and ready to use without any further intervention from a data professional.

## ❓ Business Questions

- Which products generate the highest revenue and profit, and what are their associated costs?
- What are the trends in financial metrics such as revenue, cost, and profit over time?
- How do order statuses (Completed, Pending, Returned) evolve, and what patterns can we identify?

## 🗂 Data Source

The dataset was provided by my class tutor as part of a cohort learning assignment. It contains retail transaction records including order details, product category, quantity, unit price, order/delivery dates, status, country, and payment method.

## 🛠 Tools Used

- **Microsoft Excel** — Tables, Formulas, PivotTables, Charts, Slicers, Conditional Formatting
- **XLOOKUP** — for dynamic field lookups
- **Descriptive Statistics (Data Analysis ToolPak)** — for anomaly detection

## 🔍 Project Workflow

**1. Data Preparation**
- Duplicated the source sheet to preserve the original raw data before making any changes.
- Converted the dataset into an Excel Table (`Ctrl + T`) to enable structured referencing and dynamic ranges.
- Removed duplicate records to ensure data integrity.
- Applied conditional formatting to flag empty/missing cells in red for quick visual identification during cleaning.

![Raw Sales Data](assets/raw-data.png)

**2. Calculated Fields**

To analyze profitability, I needed **Revenue**, **Cost**, and **Profit** fields, none of which existed in the raw data. I built a reference table of cost percentages per product on a separate sheet, then used `XLOOKUP` to pull the correct cost percentage into the main table based on product name:

```excel
=XLOOKUP([@[Product Name]], Cost[Product Name], Cost[Cost Percentage], "NA")
```

From there, Revenue, Cost, and Profit were derived for each transaction.

![Cost Per Unit Reference Sheet](assets/cost-per-unit-sheet.png)
![Calculated Fields with XLOOKUP](assets/calculated-fields-xlookup.png)

**3. Data Validation**

Ran a Descriptive Statistics summary on Revenue, Cost, and Profit to check for outliers or anomalies before building visuals. No significant anomalies were found, confirming the data was reliable for analysis.

**4. Analysis**

Built PivotTables to break down the data and directly address each business question — revenue/profit by product and category, trends over time, and order status distribution.

**5. Dashboard & Interactivity**

Combined the PivotTable outputs into a single dashboard, with Slicers added for filtering by product category, country, payment method, and order status — allowing anyone to explore the data without touching a formula.

## 📊 Dashboard Preview

The final dashboard brings together:
- Revenue, Cost & Profit by product category
- Revenue, Cost & Profit by city
- Order status breakdown (Completed / Pending / Returned)
- Orders by payment method
- Interactive slicers for on-the-fly filtering

![Final Dashboard](assets/final-dashboard.jpeg)

## 🚀 How to Use

1. Download `sales_data.xlsx` from this repository.
2. Open in Excel and go to the **Dashboard** sheet.
3. Use the slicers to filter by product category, country, payment method, or order status.
4. Explore the **Analysis** and **Work** sheets to see the underlying PivotTables and calculated fields.

## 💡 Skills Demonstrated

- Data cleaning and structuring (Excel Tables, duplicate removal, conditional formatting)
- Lookup functions (XLOOKUP) for cross-sheet data referencing
- Descriptive statistical analysis
- PivotTables for business-question-driven analysis
- Dashboard design with Slicers for interactivity

---

*This project was completed as part of a data analytics cohort assignment.*
