# 🚀 Fundamental Booster - Excel Practice & Project Workbook
<p align="center">
  <img src="https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Microsoft Excel Badge" />
  <img src="https://img.shields.io/badge/Status-Completed-success?style=for-the-badge" alt="Status Badge" />
  <img src="https://img.shields.io/badge/Level-Intermediate_to_Advanced-blue?style=for-the-badge" alt="Level Badge" />
</p>

> A comprehensive Excel project workbook designed to master core spreadsheet functions, dynamic formulas, conditional logic, and advanced data lookup techniques.

---

## 📌 Overview

**Fundamental Booster** is structured around practical Excel operations, ranging from basic cell formatting to advanced formulas like `XLOOKUP`, `INDEX/MATCH`, and multi-conditional criteria (`SUMIFS`, `COUNTIFS`). This repository contains templates and exercises built to sharpen analytical skills using real-world data simulations.

---

<p align="center">
  <img src="https://img.shields.io/badge/🧮_FORMULAS_&_EXPRESSIONS-APPLIED-107C41?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Excel Formulas Sticker" />
</p>

## 🧮 Comprehensive Formulas & Functions Reference

Below is the complete reference of all formulas, functions, and logic implemented across the three primary worksheets.

---

### 1. 🎓 Sheet 1: Students Grade

| Task / Analysis | Formula Applied | Function Category | Output / Description |
| :--- | :--- | :--- | :--- |
| **Total Marks** | `=SUM(C2:E2)` | Math | Sums Math, Science & English marks |
| **Grade Classification** | `=IF(G2>=250, "A", IF(G2>=200, "B", "C"))` | Nested IF | Classifies grades into A, B, or C based on total |
| **Pass/Fail Status** | `=IF(G2>=180, "Pass", "Fail")` | Logical IF | Checks pass status (Threshold >= 180) |
| **Score Threshold (>80)**| `=IF(G2>240, TRUE, FALSE)` | Boolean Logic | Returns `TRUE` if student scores above 80% average |
| **Average Score** | `=AVERAGE(C2:E2)` | Statistical | Calculates overall subject average per student |
| **Conditional Average** | `=AVERAGEIF(L2:L21, ">60")` | Conditional Avg | Calculates average score for students scoring above 60 |
| **Count Above 60** | `=COUNTIF(C2:C21, ">60")` | COUNTIF | Counts students with Math score > 60 |
| **Text Upper Case** | `=UPPER(B2)` | Text | Converts student name to uppercase (`STUDENT 1`) |
| **Text Lower Case** | `=LOWER(B2)` | Text | Converts student name to lowercase (`student 1`) |
| **Student Lookup** | `=VLOOKUP(26, A2:F21, 2, FALSE)` | VLOOKUP | Fetches student name dynamically based on ID |
| **Dynamic Reference** | `=INDIRECT(A33)` | INDIRECT | Dynamically references cell address stored in A33 (`Student 4`) |
| **Round Number** | `=ROUND(C2, 0)` | Math | Rounds off numerical scores |
| **Ceiling Function** | `=CEILING(C2, 10)` | Math | Rounds number UP to nearest multiple |
| **Floor Function** | `=FLOOR(C2, 10)` | Math | Rounds number DOWN to nearest multiple |
| **Filter Top Performers**| `=FILTER(A2:F21, C2:C21>80)` | Dynamic Array | Extracts complete list of top-performing students |
| **Date Difference** | `=DATEDIF(Start_Date, End_Date, "d")` | Date & Time | Calculates difference in days between two enrollment dates |
| **Age Calculation** | `=INT((TODAY() - DOB)/365.25)` | Date & Time | Computes age from Date of Birth |

---

### 2. 💼 Sheet 2: Sales Data

| Task / Analysis | Formula Applied | Function Category | Output / Description |
| :--- | :--- | :--- | :--- |
| **Discount Tier** | `=IF(E2>=20000, "20% Discount", "10% Discount")` | Logical IF | Assigns discount tiers based on Sales Amount |
| **Eligibility Check** | `=IF(AND(E2>=10000, C2="South"), "Eligible", "Not Eligible")` | IF + AND | Determines offer eligibility based on Region & Amount |
| **2D Matrix Summation** | `=SUMIFS($E$2:$E$21, $B$2:$B$21, $A26, $C$2:$C$21, B$25)` | SUMIFS Matrix | Creates cross-tabulation table for Total Sales by Product & Region |
| **VLOOKUP Price Search**| `=VLOOKUP(2003, A2:E21, 5, FALSE)` | VLOOKUP | Retrieves product price using Sales ID (`₹ 23,672.00`) |
| **2D Lookup Search** | `=INDEX(E2:E21, MATCH("Person 4", D2:D21, 0))` | INDEX + MATCH | Finds sales value for specific salesperson without column limits |
| **Position Search** | `=XMATCH("Mouse", B2:B21)` | Lookup & Pos | Finds relative row position of product in list |
| **Dynamic Cell Range** | `=INDIRECT("E" & 3)` | INDIRECT | Dynamically fetches sales amount from row index |
| **Dynamic Offset Range**| `=SUM(OFFSET(E2, 0, 0, 5, 1))` | OFFSET | Creates dynamic moving range for sales trend calculations (`96800`) |
| **Flexible Search** | `=XLOOKUP("Person 4", D2:D21, E2:E21)` | XLOOKUP | Modern flexible search for salesperson performance (`5264`) |
| **Sales Date Diff** | `=DATEDIF(F2, F3, "d")` | Date & Time | Calculates interval between transaction dates |

---

### 3. 👥 Sheet 3: Employee Data

| Task / Analysis | Formula Applied | Function Category | Output / Description |
| :--- | :--- | :--- | :--- |
| **Dynamic Dept Lookup** | `=INDEX(C2:C21, MATCH(3003, A2:A21, 0))` | INDEX + MATCH | Looks up employee department dynamically (`HR`) |
| **Salary Fetching** | `=XLOOKUP(3001, A2:A21, D2:D21)` | XLOOKUP | Fetches exact salary for Employee ID (`86766`) |
| **Tenure Calculation** | `=DATEDIF(E2, TODAY(), "Y")` | Date & Time | Calculates total completed years of service |
| **Joining Date Gap** | `=DATEDIF(E2, E3, "d")` | Date & Time | Computes difference in days between two employee joining dates (`8`) |
| **Indirect Emp Search** | `=INDIRECT(A25)` | INDIRECT | Dynamically references cell range (`Employee 4`) |

<p align="center">
  <img src="https://img.shields.io/badge/🖼️_WORKBOOK_SCREENSHOTS-OUTPUT-107C41?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Screenshots Sticker" />
</p>
### 🎓 Students Grade Worksheet

| 📊 Worksheet Output | 🧮 Formulas View |
| :---: | :---: |
| ![Students Grade Output](<img width="868" height="492" alt="image" src="https://github.com/user-attachments/assets/1987eaa7-a660-4cf0-9d21-756a9cf77563" />
) | ![Students Grade Output](<img width="473" height="316" alt="image" src="https://github.com/user-attachments/assets/9ee05448-f1f6-4a29-bba3-78bf5bdfdc77" />
) |

### 💼 Sales Data Worksheet

| 📊 Output View | 🧮 Formulas View |
| :---: | :---: |
| ![Sales Data Output](<img width="1060" height="432" alt="image" src="https://github.com/user-attachments/assets/542a7c63-d6df-470e-8192-5e32b159285a" />
) | ![Sales Data Output](<img width="1078" height="274" alt="image" src="https://github.com/user-attachments/assets/44a0adc4-7318-45b8-b0c1-05096c6ca5c8" />
) |
### 👥 Employee Data Worksheet

![Employee Data Output](<img width="1001" height="462" alt="image" src="https://github.com/user-attachments/assets/1f74a368-4902-4e1a-b0f5-d844028e7fff" />
)

## 🗂️ Topics Covered

* **References & Formatting:** Relative & Absolute Cell References, Custom Data Formatting, and Data Input.
* **Conditional Logic:** Standard `IF` formulas, Nested `IF` statements, and combinations with `AND`/`OR`.
* **Conditional Aggregations:** `COUNTIFS`, `SUMIFS`, and `AVERAGEIFS` for multi-criteria analysis.
* **Lookup & Reference Functions:** `VLOOKUP`, `XLOOKUP`, `XMATCH`, and dynamic searches using `INDEX/MATCH`.
* **String Manipulations:** Essential `TEXT` functions.
* **Advanced Dynamics:** Dynamic references with `INDIRECT` and `OFFSET`.
* **Data Extraction & Math:** Date/Time functions, Math operations, and the `FILTER` function.

---

## 📊 Dataset Structure

The workbook contains the following structured sheets:

1. **Students Grade**
   * Tracks student performance across subjects (**Math**, **Science**, **English**) along with their **Enrollment Date**.
   * **Columns:** `Student ID`, `Name`, `Math`, `Science`, `English`, `Enrollment Date`.

2. **Sales Data**
   * Monitors regional sales transactions, product types, and performance records.
   * **Columns:** `Sales ID`, `Product`, `Region`, `Salesperson`, `Amount`, `Date`.

3. **Employee Data**
   * Maintains internal workforce details, departmental divisions, salary structures, and tenure.
   * **Columns:** `Employee ID`, `Name`, `Department`, `Salary`, `Joining Date`.

---

## 🛠️ Main Tasks & Objectives

1. **Grade Classification:** Apply nested `IF` formulas to categorize student performance grades.
2. **Conditional Counting:** Use `COUNTIFS` to evaluate and count students scoring above specific thresholds (e.g., >60).
3. **Product Price Lookup:** Implement `VLOOKUP` to extract product prices and metrics from sales logs.
4. **Dynamic Salary Fetching:** Utilize `XLOOKUP` to retrieve employee salaries dynamically based on unique IDs.
5. **Date Analysis:** Format, sort, and analyze joining dates in the Employee dataset.
6. **Advanced Filtering:** Use the modern `FILTER` function to isolate and extract records (e.g., students scoring above 80).

---

## 🚀 Getting Started

1. Download or clone this repository to your local machine.
2. Open `Fundamental_Booster.xlsx` using Microsoft Excel or Google Sheets.
3. Complete the required worksheets by writing the designated formulas in their respective cells.

---

## 🤝 Contributing

Contributions, feedback, and suggestions to improve these exercises are always welcome! Feel free to fork this repository and submit a pull request.
