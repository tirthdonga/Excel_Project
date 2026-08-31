<div align="center">

# 📊 Excel Fundamental Booster
### *Excel Formulas, Lookup Functions, Dynamic References & Data Analysis Project*

![Excel](https://img.shields.io/badge/Microsoft%20Excel-Data%20Analysis-green?style=for-the-badge&logo=microsoftexcel&logoColor=white)
![Formulas](https://img.shields.io/badge/Excel-Formulas-blue?style=for-the-badge)
![Lookup](https://img.shields.io/badge/Lookup-VLOOKUP%20%7C%20XLOOKUP-orange?style=for-the-badge)
![Functions](https://img.shields.io/badge/Functions-Advanced-purple?style=for-the-badge)

</div>

---

# 📋 Table of Contents

- [📌 Overview](#-overview)
- [🎯 Objective](#-objective)
- [✨ Features](#-features)
- [🏗️ Workbook Structure](#️-workbook-structure)
- [🗂️ Worksheets](#️-worksheets)
- [🔍 Tasks & Excel Functions](#-tasks--excel-functions)
- [🔄 Lookup Functions](#-lookup-functions)
- [📊 Conditional & Aggregate Functions](#-conditional--aggregate-functions)
- [🔤 Text Functions](#-text-functions)
- [🔗 Dynamic Reference Functions](#-dynamic-reference-functions)
- [📅 Date, Time & Math Functions](#-date-time--math-functions)
- [🧮 Advanced Excel Functions](#-advanced-excel-functions)
- [🛠️ Tech Stack](#️-tech-stack)
- [📈 Learning Outcomes](#-learning-outcomes)
- [👤 Author](#-author)

---

# 📌 Overview

**Excel Fundamental Booster** is a practical Microsoft Excel project designed to build strong skills in formulas, data manipulation, lookup functions, conditional logic, dynamic references, date and time calculations, mathematical functions, and data filtering.

The workbook contains **Student, Sales Data, and Employee Data** worksheets and applies Excel functions to solve practical data-analysis problems.

---

# 🎯 Objective

The main objective of this project is to gain hands-on experience with Excel formulas and functions by applying them to real-world style datasets.

The project focuses on using:

- Relative & Absolute Cell References
- IF and Nested IF formulas
- AND / OR conditions
- COUNTIFS, SUMIFS, and AVERAGEIFS
- VLOOKUP, XLOOKUP, and XMATCH
- INDEX and MATCH
- Text functions
- INDIRECT and OFFSET
- Date and Time functions
- Mathematical functions
- FILTER function

---

# ✨ Features

- 🎓 Classify student grades using IF formulas
- 📊 Analyze student scores and averages
- 🔎 Search and retrieve data using lookup functions
- 💰 Analyze regional and product sales
- 👨‍💼 Retrieve employee names, departments, and salaries
- 🔤 Extract and manipulate text from names
- 🔗 Create dynamic cell and range references
- 📅 Calculate age and date differences
- 🧮 Perform financial calculations using ROUND, CEILING, and FLOOR
- 📈 Rank and filter top-performing students
- 🔄 Apply flexible and dynamic lookup criteria

---

# 🏗️ Workbook Structure

```text
📦 Excel Fundamental Booster
│
├── 📄 Student
│      ├── Student information
│      ├── Math & Science scores
│      ├── Average & Grade
│      ├── Conditional checks
│      ├── First Name extraction
│      ├── Age calculation
│      └── Student ranking / filtering
│
├── 📄 Sales Data
│      ├── Sales information
│      ├── Product & Region
│      ├── Sales Person
│      ├── Sales Amount
│      ├── Discount eligibility
│      ├── SUMIFS
│      ├── VLOOKUP
│      ├── XLOOKUP
│      ├── XMATCH
│      ├── INDIRECT
│      └── OFFSET
│
└── 📄 Employee Data
       ├── Employee information
       ├── Department
       ├── Salary
       ├── Joining Date
       ├── INDEX & MATCH
       ├── XLOOKUP
       ├── ROUND
       ├── CEILING
       └── FLOOR
```

---

# 🗂️ Worksheets

## 🎓 Student Worksheet

The **Student** worksheet contains student IDs, names, dates of birth, Math and Science scores, averages, grades, and additional calculated information.

### Main Operations

- Calculate student average scores
- Classify student grades
- Check whether students scored above 80 in both Math and Science
- Extract first names from full names
- Calculate student age
- Count students based on score conditions
- Retrieve student information using VLOOKUP
- Rank students based on scores
- Filter students based on performance

---

## 💰 Sales Data Worksheet

The **Sales Data** worksheet contains sales records with region, product, product code, salesperson, month, price, quantity, sales amount, date, and discount eligibility.

### Main Operations

- Calculate sales amount from price and quantity
- Determine discount eligibility
- Calculate regional and product sales
- Retrieve product prices using VLOOKUP
- Search salesperson performance using XLOOKUP
- Find product positions using XMATCH
- Create dynamic references using INDIRECT
- Create dynamic sales ranges using OFFSET

---

## 👨‍💼 Employee Data Worksheet

The **Employee Data** worksheet contains employee IDs, names, departments, salaries, and joining dates.

### Main Operations

- Retrieve employee names dynamically
- Retrieve employee departments
- Retrieve employee salaries
- Search employee salary using XLOOKUP
- Use INDEX and MATCH for dynamic employee lookup
- Round salaries
- Calculate ceiling salary values
- Calculate floor salary values

---

# 🔍 Tasks & Excel Functions

## 1. 📐 Relative & Absolute Cell References, Formatting & Data Input

- Enter and format data using bold, italic, currency, and custom date formats.
- Use relative cell references for normal calculations.
- Use absolute references such as `$A$1` and `$A$1:$B$10` when a reference must remain fixed.

---

## 2. 🔀 IF Formula & Nested IF Formulas

- Classify students' grades based on their scores.
- Compute discounts for sales based on price thresholds.
- Use nested conditions when multiple score or sales ranges are required.

---

## 3. 🔗 IF With AND/OR Formulas

- Identify students who scored above 80 in both Math and Science using `AND()`.
- Determine whether a product is eligible for a discount using `IF()` and a condition based on price.

---

## 4. 🔢 COUNTIFS, SUMIFS & AVERAGEIFS

- Count students who scored above 50 in Math.
- Calculate total sales for a specific region and product.
- Calculate average scores for students above a specified score.

---

## 5. 🔎 VLOOKUP Function

- Fetch student information based on Student ID from a separate lookup area.
- Retrieve product prices based on product codes.
- Use exact-match lookup with `FALSE`.

---

## 6. 🧩 INDEX & MATCH Functions

- Find employee details dynamically using Employee ID.
- Retrieve employee name, department, and salary.
- Use `MATCH()` to find the position and `INDEX()` to return the corresponding value.

---

## 7. 🔤 Excel TEXT Functions

- Extract the first name from a full name using `LEFT()` and `FIND()`.
- Convert text to uppercase using `UPPER()`.
- Convert text to lowercase using `LOWER()`.

---

## 8. 🔍 XLOOKUP Function

- Search salesperson performance using flexible search criteria.
- Retrieve employee salaries without requiring sorted data.
- Use an alternative return value when a match is not found.

---

## 9. 📍 XMATCH Function

- Find the position of a product in a sales list.
- Use exact matching to locate a product dynamically.

---

## 10. 🔗 INDIRECT Function

- Dynamically reference a cell range using a text-based reference.
- Use `INDIRECT()` with functions such as `SUM()` to calculate values from dynamically selected ranges.

---

## 11. ↔️ OFFSET Function

- Create dynamic ranges for sales trends.
- Build a range based on a starting cell and a specified number of rows or columns.

---

## 12. 📅 Date & Time Functions

- Calculate age from a date of birth.
- Find the difference in days between two dates.
- Work with joining dates and sales dates.
- Use functions such as `DATEDIF()`, `TODAY()`, and date calculations.

---

## 13. 🧮 Math Functions

- Round financial values using `ROUND()`.
- Round salary values upward using `CEILING.MATH()`.
- Round salary values downward using `FLOOR.MATH()`.

---

## 14. 🧹 Return Multiple Values Using FILTER Function

- Extract a list of top-performing students based on scores.
- Return multiple matching records from a dataset using dynamic arrays.

---

# 🔄 Lookup Functions

| Function | Purpose |
|----------|---------|
| `VLOOKUP()` | Search vertically and return matching information |
| `XLOOKUP()` | Flexible lookup with customizable match and not-found options |
| `INDEX()` | Return a value from a selected position |
| `MATCH()` | Find the position of a value |
| `XMATCH()` | Find the position of a value with flexible matching options |

### Lookup Workflow

```text
Lookup Value
     │
     ▼
Search Data Range
     │
     ▼
Find Matching Position
     │
     ├── VLOOKUP
     ├── XLOOKUP
     ├── MATCH
     └── XMATCH
     │
     ▼
Return Required Value
```

---

# 📊 Conditional & Aggregate Functions

### Conditional Functions

- `IF()`
- Nested `IF()`
- `AND()`
- `OR()`

### Conditional Aggregation

- `COUNTIFS()`
- `SUMIFS()`
- `AVERAGEIFS()`

These functions are used to analyze student performance and calculate sales based on multiple conditions.

---

# 🔤 Text Functions

The project uses text functions for manipulating names and other text values.

| Function | Purpose |
|----------|---------|
| `LEFT()` | Extract characters from the left side |
| `FIND()` | Find the position of specific text |
| `UPPER()` | Convert text to uppercase |
| `LOWER()` | Convert text to lowercase |

### Example

```text
Full Name
    │
    ▼
"Rahul Desai"
    │
    ├── LEFT()
    └── FIND()
    │
    ▼
"Rahul"
```

---

# 🔗 Dynamic Reference Functions

### INDIRECT

Creates a cell or range reference from text.

### OFFSET

Creates a dynamic range relative to a starting cell.

These functions are useful for creating flexible calculations and dynamic data ranges.

---

# 📅 Date, Time & Math Functions

### Date & Time

- `TODAY()`
- `DATEDIF()`
- Date subtraction

### Math

- `ROUND()`
- `CEILING.MATH()`
- `FLOOR.MATH()`

These functions are applied to age calculations, date analysis, and employee salary calculations.

---

# 🧮 Advanced Excel Functions

### FILTER

Returns multiple records that satisfy a condition.

### XLOOKUP

Provides flexible lookup functionality without requiring sorted data.

### XMATCH

Returns the relative position of a lookup value.

### INDEX + MATCH

Provides dynamic two-function lookup functionality.

---

# 🛠️ Tech Stack

| Technology | Purpose |
|------------|---------|
| 🟢 Microsoft Excel | Spreadsheet & Data Analysis |
| 📊 Excel Formulas | Calculations & Data Processing |
| 🔎 Lookup Functions | Searching & Retrieving Data |
| 📈 Dynamic Array Functions | Returning Multiple Values |
| 📅 Date Functions | Date & Age Calculations |

---

# 📈 Learning Outcomes

- ✅ Understanding relative and absolute cell references
- ✅ Applying professional data formatting
- ✅ Using IF and nested IF formulas
- ✅ Combining IF with AND and OR
- ✅ Applying COUNTIFS, SUMIFS, and AVERAGEIFS
- ✅ Using VLOOKUP for data retrieval
- ✅ Using XLOOKUP for flexible lookups
- ✅ Using INDEX and MATCH together
- ✅ Using XMATCH for position-based searches
- ✅ Manipulating text using Excel TEXT functions
- ✅ Creating dynamic references with INDIRECT
- ✅ Creating dynamic ranges with OFFSET
- ✅ Performing date and age calculations
- ✅ Applying ROUND, CEILING, and FLOOR
- ✅ Returning multiple values using FILTER
- ✅ Performing practical student, sales, and employee data analysis

---

# 👤 Author

<div align="center">

# Tirth Donga

[![GitHub](https://img.shields.io/badge/GitHub-Tirth_Donga-181717?style=for-the-badge&logo=github&logoColor=white)](https://github.com/tirthdonga)

Excel Fundamental Booster Project

---

### ⭐ Thank You For Visiting This Project ⭐

Made with ❤️ using Microsoft Excel

</div>
