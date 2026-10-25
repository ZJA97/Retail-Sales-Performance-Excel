# Retail Sales Performance Dashboard (Excel)

An end-to-end Excel analysis of almost 10,000 rows of US retail sales data. The raw data was cleaned and enriched using Excel Tables and calculated columns, summarised with Pivot Tables and Pivot Charts, and brought together in an interactive dashboard with slicers.

![Dashboard overview showing total sales, total profit, profit margin and the monthly sales trend](images/Overview.JPG)

## Key results

| Metric | Value |
|---|---|
| Total sales | $2,297,200.86 |
| Total profit | $286,397.02 |
| Profit margin | 12.47% |
| Period covered | January 2014 - December 2017 |

### Findings

- **Strong seasonality.** Sales peak in November (about $352k) and December (about $325k), followed by September (about $308k). February is the weakest month (about $60k). Monthly figures combine all four years.
- **December is the most profitable month** (about $43k profit), ahead of September and November.
- **East and West drive sales.** West ($725k) and East ($679k) together account for about 61% of total sales, and each sells $178k-$334k more than Central ($501k) and South ($392k).
- **Categories are fairly balanced.** Technology leads on sales ($836k), then Furniture ($742k) and Office Supplies ($719k).
- **A few products and customers stand out.** The Canon imageCLASS 2200 Advanced Copier is the top-selling product at about $61.6k, more than double the next product. The top customer by sales is Sean Miller at about $25k.

## Dataset

- **Source:** Superstore sales dataset from Kaggle (https://www.kaggle.com/datasets/vivek468/superstore-dataset-final)
<!-- TODO: confirm this is the exact Kaggle upload you used before publishing -->
- **Size:** 9,994 order lines, 21 columns (5,009 unique orders, 793 customers, 1,862 products)
- **Coverage:** United States, 4 regions, 3 customer segments, 3 product categories, 17 sub-categories
- **Note:** each row is one order line, not one order.

## What I did

**1. Data preparation**
- Loaded the data into an Excel Table (`SalesData`) so formulas use structured references and ranges update automatically.
- Order dates were a mix of true Excel dates and text strings, so I used `IF` with `ISNUMBER` to handle both and extract the date parts reliably.
- Added calculated columns:
  - **Profit Margin**: Profit / Sales
  - **Year**: `YEAR` (or extracted from text dates)
  - **Month**: `TEXT(date, "mmmm")` (or extracted from text dates)
  - **Quarter**: `ROUNDUP(MONTH / 3)`
  - **Sales Band**: High (500 or more), Medium (100-499.99), Low (under 100) using nested `IF`

**2. KPIs**
- Calculated total sales, total profit and profit margin with `SUM` and structured references on a separate Calculations sheet.

**3. Analysis**
- Built six Pivot Tables: sales by region, sales by category, monthly sales, monthly profit, top 10 products by sales, and top 10 customers by sales.
- Created Pivot Charts from them: monthly sales (line), monthly profit (line), sales by region (bar), sales by category (pie) and top products (bar).

**4. Interactivity**
- Added slicers for **Region**, **Category**, **Segment** and **Year**, connected to every Pivot Table and chart so the whole analysis can be filtered at once.
- Assembled a Dashboard sheet with the three KPI cards and four charts.

## Workbook structure

| Sheet | Contents |
|---|---|
| `in` | Raw data as an Excel Table (`SalesData`) plus calculated columns |
| `Calculations` | KPI formulas |
| `Pivot Tables` | The six Pivot Tables, Pivot Charts and the four slicers |
| `Dashboard` | KPI cards and charts |

## Tools and skills

Excel: Tables and structured references, calculated columns, `IF` / `ISNUMBER` / `YEAR` / `TEXT` / `ROUNDUP` / `SUM` / `AVERAGE`, Pivot Tables, Pivot Charts, slicers, dashboard design.

## Notes and limitations

- Monthly totals add together the same month across 2014-2017, so they show seasonality rather than a single year. Use the Year slicer to look at individual years.
- The dataset has no currency field. It is US retail data, so figures are shown in dollars.
- Possible extensions: profit by category and sub-category, the effect of discounts on profit, and year-on-year growth.

## Author

Zeyd Ali - [LinkedIn](https://www.linkedin.com/in/zeydali) | [GitHub](https://github.com/ZJA97)
