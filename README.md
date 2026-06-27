# Module 3 Assignment – Complex Data Management: What-If Analysis & Excel Pivot

This repository contains the completed assignment for **Module 3: Complex Data Management** under the professional Excel and Data Analytics course. The project demonstrates advanced data aggregation, automated formatting, dynamic analysis tools, and modern formula integration in Microsoft Excel.

## 🚀 Project Specifications & File Structure
* **File Name:** `module-3-pivot-analysis.xlsx`
* **Workbook Structure:**
  1. **Sales Raw Data:** Core dataset consisting of 31 realistic sales transactions across multiple regions and products with dynamic `Total Sales` calculation.
  2. **Pivot Report:** Advanced PivotTable formatted in **Tabular Layout** containing structured sorting, customized subtotals at the bottom of groups, an custom Calculated Field (`Average Sales per product`), and a Calculated Item (`Tech Devices`).
  3. **Goal Seek Result:** Isolated snapshot documenting the automated quantity shift required to reach a specific target value (৳500,000).
  4. **Scenario Summary:** Dedicated What-If Analysis sheet tracking performance indicators across three distinct business models: *Low Sales*, *Medium Sales*, and *High Sales*.
  5. **Pivotby Function:** Modern implementation of the dynamic `=PIVOTBY()` array formula designed for continuous summary reports.

---

## 🛠️ Step-by-Step Implementation Summary

### Step 1: Project Setup
* Initialized a fresh spreadsheet workbook and saved it using the official standard nomenclature: `module-3-pivot-analysis.xlsx`.
* Renamed the primary configuration sheet to `Sales Raw Data`.

### Step 2: Creating the Dataset for Pivot Analysis
* Outlined 7 explicit attribute headers: `Order Date`, `Region`, `Salesperson`, `Product`, `Quantity`, `Unit Price`, and `Total Sales`.
* Populated over 25 records distributed organically across dynamic months and regional demographics.
* Integrated the fundamental matrix formulation: `Total Sales = Quantity × Unit Price`.

### Step 3: Preparing Data and Creating PivotTable
* Verified semantic accuracy across columns (ensuring zero empty fields or isolated breaks) and established an independent target sheet named `Pivot Report`.
* Constructed a native cross-tabulation table filtering `Region` into rows, `Product` into columns, and routing `Total Sales` as the default summary metric.

### Step 4: Using Different PivotTable Report Layouts
* Modified structural viewports by migrating from Compact mode to a clear **Tabular Form**.
* Programmed operational settings to display group-level **Subtotals explicitly at the bottom of each region**.

### Step 5: Changing PivotTable Data Source
* Appended 5 additional logical sales rows at the base of the raw matrix to simulate live updates.
* Refreshed database connections and successfully extended the data source range grid from row 26 to **row 31**.

### Step 6: Using What-If Analysis with Goal Seek
* Targeted cell `G2` inside the data framework and configured **Goal Seek** settings.
* Set a target thresholds of `৳500,000` by allowing variance calculation against the primary item quantity (`E2`). Captured records neatly onto a dedicated sheet named `Goal Seek Result`.

### Step 7: Using Scenario Manager
* Generated 3 structural scenarios via the native Excel Scenario Framework:
  * **Low Sales Scenario**
  * **Medium Sales Scenario**
  * **High Sales Scenario**
* Adjusted quantitative variables globally and automatically exported an integrated `Scenario Summary` report.

### Step 8: Adding Calculated Field and Calculated Item
* **Calculated Field:** Introduced a custom tracking metric named `Average Sales per product` running the mathematical algorithm `= Total Sales / Quantity`.
* **Calculated Item:** Combined `Laptop` and `Monitor` categories under a single compound entity labelled `Tech Devices` inside the product array.

### Step 9: Grouping, Sorting, Filtering, and Slicer Usage
* Structured calendar parameters by grouping chronological entries into clean **Year and Month** segments.
* Administered descending index sorting protocols (**Highest Sales First**).
* Installed dynamic, highly interactive floating **Slicers for Region and Product attributes** for effortless filtering.

### Step 10: Using Pivotby Function and Final Review
* Deployed the newest analytical engine array tool: `=PIVOTBY('Sales Raw Data'!B1:B31, 'Sales Raw Data'!D1:D31, 'Sales Raw Data'!G1:G31, SUM)` on a fresh worksheet.
* *Note:* Users working on older localized desktop builds may encounter a `#NAME?` error warning flag due to version compatibility. The data maps and aggregates perfectly on modern cloud endpoints and deployment platforms supporting Microsoft 365.

---
*Developed as part of data analytics professional evaluation.*
