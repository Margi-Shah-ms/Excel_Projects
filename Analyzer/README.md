# 📊 Advanced Retail Sales & Data Analytics Dashboard

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data_Analysis-0078D4?style=for-the-badge&logo=analysis&logoColor=white)
![Dashboard](https://img.shields.io/badge/Interactive_Dashboard-FF6F00?style=for-the-badge&logo=tableau&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge&logo=github)

An end-to-end data analytics project performing descriptive statistics, linear regression, sensitivity analysis, pivot table reporting, and executive KPI dashboarding on retail transaction datasets using Excel / WPS Office.

---
![Video Demo](https://img.shields.io/badge/🎥_Video_Demo-8A2BE2?style=for-the-badge&logoColor=white)  
https://drive.google.com/file/d/1V6DzIC5aJmpwZ3SmEzo_ZQAHmtDAI-sX/view?usp=sharing

## 🎯 Executive Overview

This project processes transactional e-commerce sales data to extract actionable business insights. By executing conditional logic, What-If scenario modeling, statistical regression analysis, dynamic data lookup formulas, and interactive dynamic charts, the model provides strategic recommendations for regional growth and product category optimizations.

---

## 🧮 Formulas & Excel Logic Implemented

The spreadsheet leverages a combination of key formulas and automated Excel functionalities across multiple worksheets:

| Analysis / Task | Key Formula / Syntax | Description & Purpose |
| :--- | :--- | :--- |
| **Top Spender Name Extraction** | `=INDEX(Sheet1!B:B, MATCH(I1, Sheet1!E:E, 0))` | Dynamic two-way lookup to pull `Customer_Name` based on the maximum sales value. |
| **Top Spender ID Extraction** | `=INDEX(Sheet1!A:A, MATCH(I1, Sheet1!E:E, 0))` | Dynamic lookup retrieving `Customer_ID` corresponding to top-tier transactions. |
| **Dynamic Timestamping** | `=NOW()` | Generates a dynamic real-time timestamp for automated audit trails on the analysis sheet. |
| **What-If Sensitivity Analysis** | `Data Table (=TABLE(row_input, col_input))` | Calculates profit variance under fluctuating discount bands ($0\%$ to $20\%$). |
| **Regional Sales Aggregation** | `=SUMIF(Sheet1!C:C, "West", Sheet1!E:E)` | Sums total transaction revenue dynamically filtered by specific regions. |
| **Product Category Sales** | `=SUMIFS(Sheet1!E:E, Sheet1!C:C, "West", Sheet1!D:D, "Books")` | Multi-criteria aggregation calculating total sales for specific product lines per region. |

---
<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=waving&color=auto&customColorList=9,12,20&height=110&section=header&text=Output%20Screenshots&fontSize=26&fontColor=ffffff&animation=fadeIn&fontAlignY=38" width="100%" />
</p>
<img width="666" height="541" alt="image" src="https://github.com/user-attachments/assets/ee52b52b-5ec2-4ccd-8c44-93ed13176395" />
<img width="746" height="590" alt="image" src="https://github.com/user-attachments/assets/03386976-dc99-4eb5-adf5-1892cf0c2078" />
<img width="458" height="338" alt="image" src="https://github.com/user-attachments/assets/3a9c8d50-14ca-444b-bad1-ebbff8c21549" />
<img width="829" height="590" alt="image" src="https://github.com/user-attachments/assets/a935a53b-1a31-491c-be34-5cd73006c184" />
<img width="741" height="582" alt="image" src="https://github.com/user-attachments/assets/5facddfa-b0ea-47b0-ba4b-18ad8648e660" />
<img width="1222" height="552" alt="image" src="https://github.com/user-attachments/assets/809a3864-51a7-48fc-a99e-e2fe4fa8dd46" />
<img width="690" height="587" alt="image" src="https://github.com/user-attachments/assets/d5c1c28b-ca2d-45d0-9d40-fe242e7d5f2a" />


## ⚙️ Key Technical Features & Workflows

### 📌 1. Customer Analysis & High-Value Segment Identification
- **Conditional Formatting:** Automated top 10 purchase identification across all regions.
- **Dynamic Lookup Models:** Combined `INDEX()` and `MATCH()` functions paired with custom column filtering to track top spender profiles (`Customer_ID: CUST015`, **$1,923.66** total purchase).

### 📌 2. What-If & Profit Sensitivity Analysis
- Evaluated sales-to-profit margins under fluctuating discount bands ($0\%$ to $20\%$).
- Formatted sensitivity tables to highlight optimized discount margins without eroding bottom-line profitability.

### 📌 3. Statistical Modeling & Predictive Analytics
- **Linear Regression (Sales vs Profit):** Applied standard regression modeling to estimate profit variance as a function of overall sales throughput.
- **Descriptive Statistics:** Extracted statistical parameters (mean, median, standard deviation, and variance) to understand purchasing patterns.
- **Dynamic Timestamping:** Embedded dynamic time trackers using `=NOW()`.

### 📌 4. Dynamic Pivot Reporting & Data Storytelling
- Generated multi-tiered Pivot Tables breaking down total revenue across **Regions** (*Central, East, North, South, West*) and **Product Categories** (*Books, Clothing, Electronics, Furniture, Office Supplies*).
- Integrated month-over-month trend analysis (`Order_Date` tracking).

### 📌 5. Executive Dashboard & Visual Storytelling
- **Bar Chart:** Regional product breakdown comparison.
- **Line Chart:** Monthly revenue trajectory and seasonality spikes.
- **Pie Chart:** Regional market share distribution.
- **Executive Summary Sheet:** Structural KPI summary table with actionable strategic recommendations.

---

## 📁 Repository & Sheet Structure
