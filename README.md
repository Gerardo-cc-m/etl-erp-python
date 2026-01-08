# ERP Sales ETL Pipeline (Python)

This project is part of my Data & Business Intelligence portfolio:  
https://github.com/Gerardo-cc-m/portfolio_GerardoCasasCordero

---

## 📌 Project Overview

This project simulates a real-world ETL process used to extract commercial sales
data from an ERP system, transform and enrich it using business rules, and load
it into a structured repository ready for reporting and Business Intelligence
analysis.

The datasets included in this repository are fictitious and were created to
replicate the structure, complexity, and logic of real ERP exports used in
production environments.

---

## 🔄 ETL Workflow

### Extract
- Read ERP-like sales files with custom formats and headers
- Handle non-standard delimiters and encoding
- Skip non-data rows and normalize structure

### Transform
- Standardize column names
- Parse and normalize dates
- Convert quantities and sales values to numeric types
- Apply business rules and derive time dimensions
- Aggregate data at monthly level

### Load
- Store processed data in a structured repository
- Output datasets ready for Power BI, Excel, or further analytics

---

## 🧰 Technologies Used
- Python (pandas)
- YAML for configuration management
- CSV / TXT and Excel data sources
- Git & GitHub

---

## 📁 Project Structure

etl-erp-python/  
│── data_raw/  
│ └── sales_data_raw.csv # Fictitious ERP-like input data  
│── data_processed/  
│ └── sales_data_processed.xlsx # Fictitious processed output  
│── Automated_Reporting/  
│ └── send_email.py  
│── src/  
│ └── etl_pipeline.py  
│ └── etl_pipeline_Extract_info.py  
│── config.yaml  
│── README.md  

---

## ⚠️ Data Disclaimer

All data included in this repository is fictitious and intended solely for
demonstration and portfolio purposes. No real commercial or ERP data is exposed.
