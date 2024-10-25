
# Big Data Sorter

This script performs a sales analysis on data from an Excel file. It generates an output file with insights, including sales trends, category-based analysis, and profit trends.

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Output](#output)
- [License](#license)

## Overview

The Sales Analysis Script reads sales data from an Excel file (`data.xlsx`), processes it using Python's Pandas library, and outputs a new Excel file (`sales_analysis_output.xlsx`) with multiple sheets containing various analyses. 

## Features

1. **Overall Sales Trend**
2. **Sales by Category**
3. **Sales by Sub-Category**
4. **Profit Analysis by Month**
5. **Profit by Customer Segment**
6. **Top 10 Products by Sales**
7. **Most Selling Products by Quantity**
8. **Most Preferred Ship Mode**
9. **Most Profitable Category and Sub-Category**

## Installation

1. **Clone the repository** or download the `sales_analysis.py` script.
2. **Install dependencies**: This script requires Python 3 and a few additional libraries.

   Use the following command to install the required libraries:

   ```bash
   pip install pandas openpyxl xlsxwriter
   ```

## Usage

1. **Place the Input File**:
   Ensure `data.xlsx` is in the same directory as `sales_analysis.py`. Modify `file_path` if your file has a different name or is in another location.

2. **Run the Script**:
   Execute the script by running:

   ```bash
   python sales_analysis.py
   ```

3. **Check Output**:
   Once executed, an Excel file named `sales_analysis_output.xlsx` will be generated in the same directory. This file contains analysis results on multiple sheets.

## Output

The output file `sales_analysis_output.xlsx` includes the following sheets:
- **Sales Trend**: Monthly sales trend
- **Sales by Category**: Total sales per category
- **Sales by Sub-Category**: Total sales per sub-category
- **Profit Trend**: Monthly profit trend
- **Profit by Segment**: Profit analysis by customer segment
- **Top 10 Products**: Top 10 products based on sales
- **Most Selling Products**: Top products based on quantity sold
- **Preferred Ship Mode**: Ship mode preference based on sales
- **Profitable Category**: Profit by category
- **Profitable Sub-Category**: Profit by sub-category

## License

This project is licensed under the MIT License.
