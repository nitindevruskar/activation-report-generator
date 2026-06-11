# Activation Report Generator

A Python-based automation project that generates activation reports using Excel inputs and Supabase PostgreSQL. The solution performs data validation, country-wise analysis, year-over-year comparisons, and generates professionally formatted Excel reports.

---

## Project Overview

This project was developed to automate the manual process of preparing activation reports. It reads activation data from Excel files, validates and enriches the data using Supabase PostgreSQL tables, performs analytical calculations, and produces formatted Excel reports automatically.

The objective of this project is to reduce manual effort, improve reporting accuracy, and generate insights quickly.

---

## Technologies Used

* Python
* Pandas
* SQLAlchemy
* Supabase PostgreSQL
* OpenPyXL
* GitHub

---

## Features

* Automated Excel reporting
* Country-wise activation comparison
* Year-over-Year (YOY) analysis
* Summary dashboards
* Grand Total calculations
* Professional Excel formatting
* Automated report generation

---

## Project Architecture

Excel Input File (ATMT.xlsx)

↓

Python Script

↓

Supabase PostgreSQL

↓

Data Validation & Transformation

↓

YOY & Country-wise Analysis

↓

Formatted Excel Report Generation

---

## Database Tables Used

The project uses the following Supabase tables:

### Europe

* public.sim_master_europe_tsim
* public.activation_europe_tsim

### USA

* public.sim_master_usa_tsim
* public.activation_usa_tsim

### Thailand

* public.sim_master_thailand_ais
* public.activation_thailand_ais

---

## Input File Requirements

The Excel input file should be named:

ATMT.xlsx

The workbook must contain the following sheets:

* EU
* USA
* THA

Example columns:

ICCID | GA_DATE

---

## Installation

Clone the repository:

git clone https://github.com/nitindevruskar/activation-report-generator.git

Navigate to the project folder:

cd activation-report-generator

Install dependencies:

pip install -r requirements.txt

---

## Configure Supabase

Update the database configuration in the Python script:

DB = {
"host": "aws-1-ap-south-1.pooler.supabase.com",
"port": 6543,
"database": "postgres",
"user": "YOUR_SUPABASE_USER",
"password": "YOUR_SUPABASE_PASSWORD"
}

Important:

Do not commit actual passwords to GitHub.

---

## Running the Project

Place the Excel file (ATMT.xlsx) in the project folder.

Run:

python "Activation Report Generator.py"

The script will:

* Read Excel data
* Fetch validation data from Supabase
* Generate comparisons and summaries
* Create the final Excel report

---

## Output

The generated report includes:

* EU Activation Details
* Europe Country-wise Comparison
* USA Activation Details
* Thailand Activation Details
* Summary Dashboard
* Year-over-Year Analysis
* Grand Totals

Output example:

Activation Details DD.MM.YYYY.xlsx

---

## Repository Structure

activation-report-generator/

├── Activation Report Generator.py

├── Sample activation file.xlsx

├── requirements.txt

├── README.md

├── LICENSE

└── .gitignore

---

## Limitations

This project cannot run directly on GitHub because GitHub repositories only store code.

To execute the project, users need:

* Python installed
* Required dependencies
* Access to Supabase
* Proper database credentials

---

## Future Enhancements

Planned improvements include:

* Streamlit Web Application
* One-click report generation
* File upload through browser
* Downloadable Excel reports
* GitHub Actions automation

---

## Author

Nitin Devruskar

Python • Excel Automation • PostgreSQL • Reporting Solutions
