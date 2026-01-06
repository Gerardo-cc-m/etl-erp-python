# ERP Sales ETL Pipeline (Python)

This project is part of my Data & Business Intelligence portfolio:  
https://github.com/Gerardo-cc-m/portfolio_GerardoCasasCordero

---

## 📌 Project Overview

This project simulates a real-world ETL process used to extract commercial sales
data from an ERP system (SAP), transform and enrich it using business rules,
and load it into a structured repository ready for reporting and
Business Intelligence analysis.

The workflow reflects production scenarios where source files contain
non-standard headers, custom formats, and require data cleansing,
homologation, and aggregation before being consumed by analytics tools.

---

## 🔄 ETL Workflow

### Extract
- Read raw ERP sales files exported from SAP
- Handle non-standard headers and custom file structures
- Manage encoding and delimiter variations

### Transform
- Standardize and clean column names
- Parse and normalize order dates
- Convert quantities and sales values to proper data types
- Apply business rules (brand logic, country-specific conditions)
- Harmonize data using master datasets (products, importers)
- Convert currency values
- Aggregate data at monthly level

### Load
- Store consolidated datasets in Excel format
- Output data ready for consumption in Power BI and Excel

---

## 🧰 Technologies Used
- Python (pandas)
- YAML for configuration management
- Excel as processed data output
- Git & GitHub

---

## 📁 Project Structure

etl-erp-python/   
│   
├── data/   
│ ├── raw/ # Raw ERP extracts (not included)   
│ └── processed/ # Processed fictitious datasets   
│   
├── src/   
│ └── generar_base_pedidos.py   
│   
├── config.yaml   
└── README.md   


---

## ⚠️ Data Disclaimer

All datasets included in this repository are fictitious and were created
exclusively for demonstration and portfolio purposes.

Real production data extracted from SAP is not included due to
confidentiality constraints.

---

## 🎯 Purpose

This repository is intended to demonstrate:
- Practical ETL development in Python
- Handling of real-world ERP data complexity
- Data preparation for Business Intelligence and KPI reporting
