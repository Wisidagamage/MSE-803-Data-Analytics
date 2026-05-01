# Week 3 – Activity 2: Data Cleaning and Data Visualization

**Course:** MSE-803 Data Analytics (2511-YCCIA-MSE)  
**Student:** Mukesh  
**Due:** 1 May 2026

---

## Overview

This activity demonstrates data cleaning and visualization on a messy employee dataset (`messy_dataset_Mukesh.csv`), including a Pearson correlation heatmap and outlier detection using the IQR method.

---

## Files

| File | Description |
|------|-------------|
| `messy_dataset_Mukesh.csv` | Original raw dataset (10 rows, multiple quality issues) |
| `data_cleaning_visualization.py` | Python script — full cleaning + visualization pipeline |
| `Data_Cleaning_Visualization_Mukesh.xlsx` | Final deliverable — 5-sheet workbook with charts |
| `charts/heatmap.png` | Pearson correlation heatmap |
| `charts/boxplots.png` | Box plots for outlier detection (Age & Salary) |
| `charts/salary_by_country.png` | Average salary by country bar chart |
| `charts/age_distribution.png` | Age distribution histogram |

---

## Data Issues Found (11 total)

| Row | Record | Column | Issue | Fix Applied |
|-----|--------|--------|-------|-------------|
| 2 | Bob | ID | Duplicate ID (same as row 3) | Merged rows, removed duplicate |
| 3 | Bob | Salary | Missing value | Removed (duplicate row) |
| 4 | Charlie | Join Date | Missing value | Marked as `Unknown` |
| 5 | David | Age | Non-numeric: `thirty-eight` | Converted to `38` |
| 6 | Eve | ID | Missing value | Assigned next available ID: `6` |
| 6 | Eve | Join Date | Invalid date: `2019-13-01` (month=13) | Marked as `Unknown` |
| 7 | — | Name | Missing value | Filled as `Unknown` |
| 7 | — | Salary | Non-numeric: `sixty five thousand` | Converted to `65000` |
| 8 | Grace | Country | Missing value | Marked as `Unknown` |
| 9 | Heidi | Age | Missing value | Filled with median (`29`) |
| 9 | Heidi | Salary | Missing value | Filled with median (`62000`) |

---

## Cleaned Dataset (9 rows)

| ID | Name | Age | Country | Salary | Join Date |
|----|------|-----|---------|--------|-----------|
| 1 | Alice | 25 | NZ | 55,000 | 15/01/2020 |
| 2 | Bob | 30 | NZ | 60,000 | 20/02/2020 |
| 4 | Charlie | 35 | AUS | 72,000 | Unknown |
| 5 | David | 38 | NZ | 68,000 | 1/11/2019 |
| 6 | Eve | 29 | AU | 59,000 | Unknown |
| 7 | Unknown | 40 | NZ | 65,000 | 30/05/2018 |
| 8 | Grace | 22 | Unknown | 64,000 | 25/07/2021 |
| 9 | Heidi | 29 | AUS | 62,000 | 25/07/2021 |
| 10 | Ivan | 27 | NZ | 58,000 | 15/03/2019 |

---

## Pearson Correlation Results

|  | ID | Age | Salary |
|--|-----|-----|--------|
| **ID** | 1.0000 | -0.0876 | 0.0700 |
| **Age** | -0.0876 | 1.0000 | **0.6285** |
| **Salary** | 0.0700 | **0.6285** | 1.0000 |

- **Age vs Salary (r = 0.629):** Moderate positive correlation — older employees tend to earn higher salaries.
- **ID vs Age / Salary:** Near-zero correlation — hire sequence does not predict age or salary.

---

## Outlier Detection (IQR Method)

| Column | Q1 | Q3 | IQR | Lower Fence | Upper Fence | Outliers |
|--------|----|----|-----|-------------|-------------|----------|
| Age | 27.0 | 35.0 | 8.0 | 15.0 | 47.0 | None |
| Salary | 59,000 | 65,000 | 6,000 | 50,000 | 74,000 | None |

No outliers were detected in the cleaned dataset using the 1.5×IQR rule.

---

## How to Run

```bash
python data_cleaning_visualization.py
```

**Dependencies:**
```bash
pip install pandas numpy matplotlib seaborn openpyxl
```

---

## Excel Workbook Structure

| Sheet | Contents |
|-------|----------|
| 1_Raw Data | Original data with red-highlighted problem cells |
| 2_Issues Log | All 11 issues with type, raw value, and fix taken |
| 3_Cleaned Data | Final clean dataset + summary statistics |
| 4_Correlation & Outliers | Pearson matrix, interpretations, IQR fence table |
| 5_Visualizations | 4 embedded charts |
