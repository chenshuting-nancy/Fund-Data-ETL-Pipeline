# 📊 Fund Data ETL Pipeline (基金单数据提取自动化)

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![Status](https://img.shields.io/badge/Status-Production%20Ready-green)
![Focus](https://img.shields.io/badge/Focus-Finance%20%26%20Automation-orange)

## Project Overview (项目简介)

**Fund Data ETL Pipeline** is an automated data processing solution designed to tackle the fragmentation of financial data in Asset Management. 

In the mutual fund industry, transaction statements come from diverse sources with inconsistent formats. This tool serves as a **Unified Data Adaptor**, capable of automatically identifying, extracting, and standardizing data from **20+ distribution platforms** (Banks, Brokerages, Third-party agencies) and **5 core document types**.

> **Impact:** This tool transforms a manual reconciliation process that typically takes hours into a sub-minute automated task, ensuring 100% data accuracy for Hundsun valuation systems.

## Key Features (核心功能)

* **Multi-Source Compatibility:** Implemented a scalable strategy pattern to handle distinctive formats from 20+ financial institutions (e.g., ICBC, CMB, Alipay, Tiantian Fund, etc.).
* **Intelligent Classification:** Automatically detects document types based on file signatures and naming conventions (e.g., Confirmation Notes, Dividend Statements, Settlement Sheets).
* **Robust Data Cleaning:** Uses Advanced RegEx and Pandas to normalize "dirty data" (merged cells, irregular headers, non-standard date formats).
* **Batch Processing:** Capable of processing hundreds of files locally in seconds, outputting a unified CSV/Excel standard ready for SQL ingestion.

## Tech Stack (技术栈)

* **Core Logic:** Python 3.x
* **Data Manipulation:** Pandas, NumPy
* **File Parsing:** `pdfplumber` (PDF), `openpyxl`/`xlrd` (Excel), `os`/`shutil` (File I/O)
* **Pattern Matching:** Regular Expressions (Re)

## Workflow Architecture (工作流)

```mermaid
graph TD
    A[Input Folder: Mixed Raw Files] --> B{File Classifier};
    B -->|Platform A| C[Parser Strategy A];
    B -->|Platform B| D[Parser Strategy B];
    B -->|Platform N...| E[Parser Strategy N];
    C --> F[Data Cleaning & Normalization];
    D --> F;
    E --> F;
    F --> G[Validation Rules];
    G --> H[Output: Standardized Master Data];
