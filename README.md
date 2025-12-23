# Python-development---Project

# 📊 Excel Report Generator

## 📌 Project Overview

The **Excel Report Generator** is a Python-based desktop application that allows users to upload a CSV file and automatically generate a fully formatted Excel report. The tool creates **summary statistics, pivot tables, and charts** to provide actionable insights from raw data, eliminating the need for manual report creation.

---

## 🛠️ Technologies Used

* **Programming Language:** Python 3.x
* **GUI Framework:** Tkinter
* **Data Processing:** Pandas
* **Visualization:** Matplotlib
* **Excel Manipulation:** OpenPyXL

---

## 📁 Project Structure

### 1. CSV File Input

* Users can upload any CSV file with relevant data for report generation.
* Data is read and processed using **Pandas**.

### 2. Excel Report Generation

* Generates multiple sheets in the Excel file:

  * **Raw Data** – Original CSV data.
  * **Summary Stats** – Statistical summary using `describe()`.
  * **Pivot Table** – Category-wise aggregation (optional if `Category` exists).
  * **Chart** – Trend chart of the `Quantity` column.

### 3. Styling

* Headers in all sheets are formatted as **bold** using OpenPyXL for better readability.

### 4. Charts

* Matplotlib is used to create a **Quantity Trend chart**, which is embedded into the Excel file.

---

## ⚙️ Core Features

* ✅ Upload CSV files via a graphical interface
* ✅ Generate complete Excel reports automatically
* ✅ Pivot tables for categorical summaries
* ✅ Charts visualizing trends
* ✅ Styled Excel sheets with bold headers
* ✅ User notifications for successful operations or errors

---

## 📤 How to Use

1. Run the Python script `ExcelReportGenerator.py`.
2. Click **“Upload CSV File”** to select your CSV.
3. Click **“Generate Excel Report”** to create the Excel file.
4. Save the file in your preferred location.
5. Open the Excel report to view raw data, summary stats, pivot tables, and charts.

---

## 🧪 Sample Usage

```python
# Launch the application
ExcelReportGenerator()
```

* CSV file should have a `Category` column (optional) and `Quantity` column for pivot table and chart generation.
* The Excel file will include formatted sheets and embedded charts.

---

## 📦 Deliverables

1. Python script (`ExcelReportGenerator.py`)
2. GUI for uploading CSV and generating reports
3. Excel report with:

   * Raw Data
   * Summary Statistics
   * Pivot Table (if applicable)
   * Quantity Trend Chart

## Output screenshots

Below are the screenshots showing the application interface and generated Excel report.

<img width="753" height="670" alt="{EF27BBC9-8D8A-4006-864A-780394CD902E}" src="https://github.com/user-attachments/assets/bb1b4eb0-9672-4e62-ab0c-38dfffdcec7f" />


<img width="896" height="668" alt="{1AAEAEC7-73FF-404F-BDC0-0855F704FCB2}" src="https://github.com/user-attachments/assets/d6862c63-6bb1-42ca-8258-f3a52aa0f955" />


<img width="889" height="664" alt="{D9A126FB-6647-49F9-932A-8A9F0394AD83}" src="https://github.com/user-attachments/assets/e6f871bc-ce8c-4ee2-83d2-737cdb6136b5" />

