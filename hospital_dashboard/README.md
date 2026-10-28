# Hospital Management & Analytics Dashboard

## 📌 Overview
This repository / project showcases an interactive **Hospital Management & Analytics Dashboard** built in Microsoft Excel (`ShahMargi_ExcelDashboard.xlsx`). The dashboard tracks critical healthcare metrics, patient admissions, department-wise performance, and operational KPIs, providing hospital administrators with actionable insights through automated summaries, pivot tables, and dynamic visual charts.

---
## VIDEO LINK: 
https://drive.google.com/file/d/15CX0_NLyKl5pTsON6ughpiQ-BPxClBfA/view?usp=sharing

## 🏥 Key Metrics & Data Tracked
- **Patient Demographics & Admissions:** Tracking patient inflow, discharge status, length of stay, and age/gender distributions.
- **Departmental Performance:** Emergency, Cardiology, Orthopedics, Pediatrics, ICU, and General Surgery metrics.
---

## 📊 Structure of the Workbook
The workbook is structured into logical layers to ensure clean data management and powerful reporting:
1. **`Dashboard` Sheet:** 
   - The central command center featuring high-level KPI summary cards, interactive charts, and filtering slicers for executive review.
2. **`Summary / Pivot Tables` Sheet:** 
   - Houses aggregated pivot tables and summary metrics that power the visual charts on the main dashboard.
3. **`Raw Data` Sheet:** 
   - The primary database containing granular, row-level hospital transactional records (Patient IDs, Dates, Departments, Doctors, Status, and Costs).
---

## 🧮 Formulas Used
To automate calculations, summaries, and lookups across the workbook, standard and advanced Excel functions are implemented:
- **Aggregation & Conditional Counting/Summing:** `SUMIFS`, `COUNTIFS`, `AVERAGEIFS` (used extensively in summary blocks to calculate department-wise revenue, patient counts, and average wait times).
- **Lookups & Data Retrieval:** `XLOOKUP` / `VLOOKUP` and `INDEX/MATCH` for fetching patient records and doctor details dynamically.
- **Logical & Date Handling:** `IF`
---

## 📈 Charts & Visualizations
The dashboard incorporates a variety of clean, professionally formatted charts to visualize hospital data at a glance:
- **Bar / Column Charts:** Comparing patient volume and revenue across different medical departments.
- **Line Charts:** Tracking patient admission trends over time (monthly/weekly analysis).
- **Donut / Pie Charts:** Showing bed occupancy percentage and breakdown of patient categories.

---

## output: 
<img width="876" height="564" alt="image" src="https://github.com/user-attachments/assets/eb61c1e9-07a5-411e-9a89-a9921511a9a0" />
<img width="1191" height="540" alt="image" src="https://github.com/user-attachments/assets/84a36822-5be3-4cc6-9e3a-88845613178e" />
<img width="789" height="514" alt="image" src="https://github.com/user-attachments/assets/9e241dfb-a379-4490-a897-85b0708cd2d1" />
<img width="850" height="562" alt="image" src="https://github.com/user-attachments/assets/027ac86a-a1f8-484b-a7a8-6ef20f3abd8e" />
<img width="1420" height="549" alt="image" src="https://github.com/user-attachments/assets/9d8effec-5cb9-4b40-a129-0d3a98974390" />
<img width="858" height="589" alt="image" src="https://github.com/user-attachments/assets/d97e9432-232e-4e24-ae48-b8fd2bfcfeee" />
<img width="863" height="568" alt="image" src="https://github.com/user-attachments/assets/f1e47f39-1831-4ae6-8651-9acb77ac90dc" />


## 🚀 How to Use & Interact
1. **Open the Workbook:** Download and open `ShahMargi_ExcelDashboard.xlsx` in Microsoft Excel (Desktop application recommended for full Pivot Table and Slicer functionality).
2. **Use Slicers:** Click on the interactive slicers on the Dashboard tab to filter data instantly by Date Range, Department, or Patient Status.
3. **Refresh Data:** 
   - Add or update records in the **Raw Data** sheet.
   - Go to the **Data** tab on the Excel ribbon and click **Refresh All** (`Ctrl + Alt + F5`) to update all Pivot Tables and Charts automatically.

---

## 🛠️ Requirements
- **Software:** Microsoft Excel 2016, 2019, Excel 365, or Google Sheets.
- **Skills Demonstrated:** Advanced Excel, Data Modeling, Pivot Tables & Pivot Charts, Formula Automation, and Dashboard UI/UX Design.

---
