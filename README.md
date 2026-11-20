# 🍕 Pizza Sales Dashboard (Excel)

An interactive Excel dashboard that analyzes pizza sales: revenue, orders, quantity, and customer buying patterns by category, size, month, weekday, and time of day. Built with a cleaned data table, pivot tables, pivot charts, and slicers.

---

## 📌 Project Overview

| Item | Detail |
|------|--------|
| **Tool** | Microsoft Excel |
| **Techniques** | Data cleaning, data modelling, Pivot Tables, Pivot Charts, Slicers, Timeline filter, KPI cards |
| **Workbook sheets** | `Data` → `pivot table` → `Dashboard` |
| **Period analyzed** | 2015 |

**Goal:** turn raw pizza order data into a single-page dashboard that helps answer: *What sells? When does it sell? Which sizes and categories drive revenue?*

---

## 📊 Key KPIs

| KPI | Value |
|-----|-------|
| Total Sales | **$8,436** |
| Total Orders | **499** |
| Customers | **499** |
| Quantity Sold | **511** |

---

## 🧱 Step 1 – Data Modelling

The workbook follows a simple, flat-table model that feeds every pivot table.

1. **Import** the raw order data into the `Data` sheet.
2. **Clean** the data: remove duplicates and blanks, fix data types (dates, numbers, text), and standardize column names.
3. **Create helper columns** (used as pivot row labels):
   - `Month` – extracted from `order_date`
   - `Weekday` – extracted from `order_date`
   - `order time range` – bucket of the order time: *morning, afternoon, evening, night*
   - `Sales` – calculated as quantity × unit price
4. **Convert to an Excel Table** (`Ctrl + T`) so pivot tables refresh automatically when data changes.

### Main fields

| Field | Description |
|-------|-------------|
| `order_id` | Unique order identifier |
| `order_date` | Date of the order |
| `pizza_name` | Name of the pizza |
| `pizza_category` | Chicken / Classic / Supreme / Veggie |
| `pizza_size` | S / M / L / XL |
| `quantity` | Number of pizzas ordered |
| `Sales` | Revenue of the order line |
| `order time range` | Morning / Afternoon / Evening / Night |

> The source columns include Arabic labels (e.g., category, size, date, quantity), so the workbook supports bilingual field names.

---

## 🧮 Step 2 – Pivot Tables (`pivot table` sheet)

Each dashboard visual is backed by one pivot table.

| # | Pivot Table | Rows | Values | Output chart |
|---|-------------|------|--------|--------------|
| 1 | Sales by Category | `pizza_category` | Sum of Sales (% of Grand Total) | 3-D Pie |
| 2 | Sales by Size | `pizza_size` | Sum of Sales | Pie |
| 3 | Monthly Trend | `Month` | Sum of Sales | Column chart |
| 4 | Weekday Trend | `Weekday` | Sum of Sales | Line chart |
| 5 | Top Pizzas | `pizza_name` (Top N filter, sorted descending) | Sum of Quantity | Bar chart |
| 6 | Time of Day | `order time range` | Sum of Quantity (% of Grand Total) | 3-D Pie |

**How to build any of them**
1. Select the table in `Data` → **Insert → PivotTable** → place it in the `pivot table` sheet.
2. Drag the field to **Rows** and the measure to **Values**.
3. For percentages: *Value Field Settings → Show Values As → % of Grand Total*.
4. For Top N: *Row Labels filter → Value Filters → Top 10 → set the count*, then sort descending.
5. Insert a **PivotChart** from each pivot table.

---

## 🖥️ Step 3 – Building the Dashboard (`Dashboard` sheet)

1. **Layout** – a red title banner, an orange KPI strip, and a grid of charts.
2. **KPI cards** – Total Sales, Total Orders, Customers, Quantity (linked to cells with `SUM` / `COUNT` / `DISTINCTCOUNT`-style formulas).
3. **Charts** – cut each PivotChart from the `pivot table` sheet and paste it into `Dashboard`, then format titles, colors, and data labels.
4. **Slicers** – *PivotTable Analyze → Insert Slicer*:
   - `pizza_category`
   - `pizza_size`
5. **Timeline** – *Insert Timeline* on `order_date` (months view).
6. **Connect filters to all pivots** – *Slicer → Report Connections* and tick every pivot table so one click updates the whole dashboard.
7. **Final polish** – hide gridlines, remove field buttons from charts, align objects, and protect the layout.

### Source code (calculated column formulas)

```excel
' Month name from order date
=TEXT([@order_date], "mmm")

' Weekday name from order date
=TEXT([@order_date], "dddd")

' Sales per row (if not already in the data)
=[@quantity] * [@unit_price]

' Time-of-day bucket from order time (adjust the hours to your data)
=IF([@order_time]<TIME(12,0,0),"morning",
  IF([@order_time]<TIME(17,0,0),"afternoon",
  IF([@order_time]<TIME(21,0,0),"evening","night")))
```

KPI cells:

```excel
' Total Sales
=SUM(Data[Sales])

' Total Orders
=SUMPRODUCT(1/COUNTIF(Data[order_id],Data[order_id]))

' Quantity
=SUM(Data[quantity])
```

---

## 🔍 Key Insights

- **Category:** Veggie leads revenue at 28.04%, followed by Classic (26.96%), Supreme (24.32%), and Chicken (20.67%).
- **Size:** Large pizzas generate the most revenue (≈48%, $4,005), then Medium (27%), Small (22%), and XL (3%).
- **Month:** Sales peak in **June ($859)** and **August ($858)**; the lowest month is **March ($520)**.
- **Weekday:** **Friday** is the strongest day ($1,535); **Wednesday** is the weakest ($853).
- **Time of day:** **Afternoon (45%)** and **evening (41%)** account for most orders; morning is only 5%.
- **Best sellers:** *The Classic Deluxe Pizza* (30 units) leads, then *The Pepperoni Pizza* (26) and *The Four Cheese Pizza* (25). The Thai Chicken and Sicilian pizzas tie at 22 units, so the "Top 5" chart shows six items.

---

## 💡 Recommendations

- Promote afternoon/evening bundles, and test morning offers to grow the weakest time slot.
- Run mid-week (Wednesday) discounts to lift the slowest day.
- Keep Large as the hero size; review whether XL is worth keeping on the menu.
- Push Veggie and Classic items, which bring the highest revenue share.

---

## 📁 Repository Structure

```
pizza-sales-dashboard-excel/
│
├── Pizza_Sales_Dashboard.xlsx      # Workbook: Data, pivot table, Dashboard
├── README.md                       # Project documentation
└── images/
    ├── dashboard.png               # Dashboard screenshot
    └── pivot_tables.png            # Pivot tables screenshot
```

---

## 🖼️ Screenshots
<img width="1680" height="652" alt="Pizza Sales Dashboard" src="https://github.com/user-attachments/assets/96b579ae-8897-4b14-b339-d12543baecca" />


<img width="1492" height="659" alt="pivot table pizaa" src="https://github.com/user-attachments/assets/4ea0be4e-b0be-4315-9f4d-d622434501e4" />

---

## ▶️ How to Use

1. Download `Pizza_Sales_Dashboard.xlsx`.
2. Open it in Microsoft Excel (2016 or later recommended).
3. Go to the `Dashboard` sheet and use the slicers and timeline to filter.
4. After changing the `Data` sheet, click **Data → Refresh All**.

---

## 🛠️ Skills Demonstrated

Data cleaning · Data modelling · Pivot Tables & Charts · Slicers & Timelines · KPI design · Dashboard layout · Business insight storytelling

---

## 👤 Author
https://github.com/amalabdelkawy-web
https://www.linkedin.com/in/amal-abdelkawy-69453728a/

