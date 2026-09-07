# Excel VLOOKUP, INDEX & MATCH Practice

This project contains my hands-on practice with Excel lookup and reference functions.

The main goal of this practice was to understand how Excel can be used to search, retrieve, and dynamically reference data — important skills for data analysis and reporting.

## 📌 Topics Covered

* VLOOKUP
* MATCH
* INDEX
* VLOOKUP + MATCH
* Exact Match
* Dynamic column lookup
* Data retrieval from tables

## 🧠 What I Practiced

### 1. VLOOKUP

Used VLOOKUP to search for a value in the first column of a table and return the corresponding value from another column.

Example:

```excel
=VLOOKUP(A2,$A$2:$D$10,3,FALSE)
```

### 2. MATCH

Used MATCH to find the position of a specific value within a row or column.

Example:

```excel
=MATCH(B2,$B$1:$E$1,0)
```

### 3. INDEX

Used INDEX to return a value from a specific position within a range.

Example:

```excel
=INDEX($B$2:$E$10,3,2)
```

### 4. VLOOKUP + MATCH

Combined VLOOKUP and MATCH to dynamically identify the required column instead of manually entering the column number.

Example:

```excel
=VLOOKUP(A2,$A$2:$E$10,MATCH(B1,$A$1:$E$1,0),FALSE)
```

This makes the lookup more flexible when the position of the required column changes.

## 📊 Skills Practiced

* Microsoft Excel
* Lookup Functions
* Data Retrieval
* Formula Building
* Dynamic References
* Data Analysis Fundamentals

## 🎯 Learning Objective

The objective of this exercise was to strengthen my Excel fundamentals and understand how lookup functions can be applied to real-world data analysis tasks.

## 📁 Project File

The Excel workbook containing the practice exercises is available in the `Excel` folder.

## 🚀 Next Steps

I plan to continue improving my Excel skills by practicing:

* XLOOKUP
* IF / IFS
* SUMIF / SUMIFS
* COUNTIF / COUNTIFS
* Pivot Tables
* Conditional Formatting
* Data Cleaning
* Excel Dashboards

---

**Learning by practicing, one Excel formula at a time. 📊**
