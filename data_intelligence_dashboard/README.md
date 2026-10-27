# 📊 Data Intelligence & Sales Analytics Dashboard

![Microsoft Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Data Analysis](https://img.shields.io/badge/Data_Analysis-0078D4?style=for-the-badge&logo=analysis&logoColor=white)
![Dashboard](https://img.shields.io/badge/Interactive_Dashboard-FF6F00?style=for-the-badge&logo=tableau&logoColor=white)
![Status](https://img.shields.io/badge/Status-Completed-success?style=for-the-badge&logo=github)

An end-to-end data analytics project performing descriptive statistics, linear regression, sensitivity analysis, pivot table reporting, and executive KPI dashboarding on retail transaction datasets using Excel / WPS Office.

---
![Video Demo](https://img.shields.io/badge/🎥_Video_Demo-8A2BE2?style=for-the-badge&logoColor=white)  


## 🎯 Executive Overview

This project processes transactional e-commerce sales data to extract actionable business insights. By executing conditional logic, What-If scenario modeling, statistical regression analysis, dynamic data lookup formulas, and interactive dynamic charts, the model provides strategic recommendations for regional growth and product category optimizations.

---

## 🧮 Formulas & Excel Logic Implemented

The spreadsheet leverages a combination of key formulas and automated Excel functionalities across multiple worksheets:

| Analysis / Task | Key Formula / Syntax | Description & Purpose |
| :--- | :--- | :--- |
| **Top Performing Product** | `=INDEX(Sheet1!F2:F250, MODE(MATCH(Sheet1!F2:F250, Sheet1!F2:F250, 0)))` | Mode-based dynamic array formula extracting the most frequently purchased product name. |
| **Customer Name Matching** | `=IF(COUNTIF(Sheet1!$D$2:$D$250, Sheet1!D2)>1, Sheet1!D2, "Unique")` | Evaluates repeating customer names and extracts recurring high-value entries. |
| **Customer Initials / Abbrev.** | `=LEFT(D2, 1) & MID(D2, FIND(" ", D2) + 1, 1)` | Uses text string functions to generate customer name initials for structured reporting. |
| **Dynamic Timestamping** | `=IF(G3<>"", IF(H3="", NOW(), H3), "")` | Iterative logic generating real-time automated audit timestamps. |
| **Linear Regression ($R^2$)** | Data Analysis Toolpak ($R^2 = 0.6979$) | Evaluates the strong positive linear correlation between `Unit_Price` and `Total_Amount`. |
| **What-If Analysis** | Scenario Manager & Goal Seek | Evaluates revenue variance under varying order quantities ($Q=6, Q=3$) and pricing models. |

---
<p align="center">
  <img src="https://capsule-render.vercel.app/api?type=waving&color=004B50&height=110&section=header&text=Output%20Screenshots&fontSize=26&fontColor=ffffff&animation=fadeIn&fontAlignY=38" width="100%" />
</p>



## ⚙️ Key Technical Features & Workflows

### 📌 1. Customer Analysis & High-Value Segment Identification
- **Conditional Formatting:** Automated icon sets and directional arrows highlighting key performance bands.
- **Dynamic Text Processing:** String extraction methods (`LEFT`, `MID`, `FIND`) applied to `Customer_Name` for concise reporting.

### 📌 2. What-If & Scenario Modeling
- Implemented **Scenario Manager** and **Goal Seek** to simulate revenue variations based on unit pricing and order quantities.
- Back-calculated target pricing models to optimize bottom-line totals.

### 📌 3. Statistical Modeling & Predictive Analytics
- **Linear Regression:** Applied Data Analysis Toolpak regression models ($R^2 = 0.6979$) to evaluate pricing dependencies.
- **Temporal Functions:** Configured date-time computations using `TODAY()`, `NOW()`, `DATEDIF()`, and `EOMONTH()`.

### 📌 4. Dynamic Pivot Reporting & Data Storytelling
- Generated multi-tiered Pivot Tables breaking down total revenue across **Regions** (*Central, North, South, West*) and **Product Categories** (*Appliances, Electronics, Furniture*).
- Incorporated payment method distribution analysis (`Cash`, `Credit Card`, `Debit Card`, `PayPal`).

### 📌 5. Executive Dashboard & Visual Storytelling
- **Bar Chart:** Category-wise revenue distribution comparison (highlighting *Electronics* as top revenue driver at ~$171k).
- **Line Chart:** Date-wise trend analysis tracking temporal performance across 2024-2025.
- **Pie Chart:** Payment method share breakdown.
- **Executive KPI Cards:** Top-level summary blocks tracking Total Revenue (`$2,29,192.47`), Total Transactions (`249`), Top Category Revenue (`$4,499.95`), and Average Order Value (`$916.77`).
