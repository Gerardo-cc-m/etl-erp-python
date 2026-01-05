# ERP Sales ETL Pipeline (Python)

## 📌 Project Overview
This project simulates a real-world ETL process used to extract commercial sales data from an ERP system, transform and enrich it using business rules, and load it into a structured repository ready for reporting and Business Intelligence analysis.

The workflow replicates production scenarios where source files contain non-standard headers, custom delimiters, and require data cleansing before analysis.

---

## 🔄 ETL Workflow
1. **Extract**
   - Read raw ERP-like sales files with pipe (`|`) delimiter
   - Skip non-data header rows
   - Handle encoding and formatting issues

2. **Transform**
   - Standardize column names
   - Parse and normalize dates
   - Convert sales amounts to numeric values
   - Apply business rules and derive time dimensions (year, month)

3. **Load**
   - Store processed data in a structured repository
   - Output datasets ready for Power BI, Excel, or further analytics

---

## 🧰 Technologies Used
- Python (pandas)
- YAML for configuration management
- CSV / TXT data sources
- Git & GitHub

---

## 📁 Project Structure
```text
etl-erp-python/
│── data_raw/
│   └── sales_data_raw.csv
│── data_processed/
│   └── sales_data_processed.csv
│── src/
│   └── etl_pipeline.py
│── config.yaml
│── README.md
